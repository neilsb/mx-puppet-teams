import {
	PuppetBridge,
	IRemoteUser,
	IReceiveParams,
	IRemoteRoom,
	IMessageEvent,
	Log,
	IRetList,
	MessageDeduplicator,
	ISendingUser,
	IFileEvent,
} from "mx-puppet-bridge";
import { AuthProvider } from "./auth/auth-provider";
import { TeamsAuthProvider } from "./auth/teams-auth-provider";
import { MSTeamsStore } from "./store";
import { IClientOpts, Message, Chat, TeamsClient } from "./teams-client";
const htmlToFormattedText = require("html-to-formatted-text");


// create our log instance
const log = new Log("TeamsPuppet:teams");

interface ITeamsPuppet {
	client: TeamsClient;
	data: any;
	clientStopped: boolean;
}

// we can hold multiple puppets at once...
interface ITeamsPuppets {
	[puppetId: number]: ITeamsPuppet;
}

export class App {

	private puppets: ITeamsPuppets = {};
	private store: MSTeamsStore;
	private authProvider: AuthProvider;
	private messageDeduplicator: MessageDeduplicator;

	constructor(
		private puppet: PuppetBridge,
	) {
		this.store = new MSTeamsStore(puppet.store);
		this.messageDeduplicator = new MessageDeduplicator();
	}

	public async init(): Promise<void> {
		await this.store.init();
		this.authProvider = new AuthProvider(this.store);
	}

	public async removePuppet(puppetId: number) {
		log.info(`Removing puppet: puppetId=${puppetId}`);
		await this.puppet[puppetId].client.stop();
		delete this.puppets[puppetId];
	}

	public async deletePuppet(puppetId: number) {
		log.info(`Got signal to quit Puppet: puppetId=${puppetId}`);

		// TODO:  Delete Tokens

		await this.puppet[puppetId].client.stop();
		await this.removePuppet(puppetId);
	}

	public async newPuppet(puppetId: number, data: any) {
		log.info(`Adding new Puppet: puppetId=${puppetId}`);
		log.info(`Got data: dat=${JSON.stringify(data)}`);

		// check for updated access token
		if(data.auth_code) {
		await this.authProvider.checkForNewAuthorization(puppetId, data.auth_code);
			delete(data.auth_code);
			this.puppet.setPuppetData(puppetId, data);
		}

		if (this.puppets[puppetId]) {
			await this.removePuppet(puppetId);
		}
		const client = new TeamsClient(<IClientOpts>{
			authProvider: new TeamsAuthProvider(puppetId, this.authProvider),
			ownerUserId: data.userId,
			puppetId,
			knownMessage: async (r: string, e: string): Promise<boolean> =>
				(await this.puppet.eventSync.getMatrix({ roomId: r, puppetId: puppetId }, e)).length > 0
		});

		client.on("connected", async () => {
			await this.puppet.sendStatusMessage(puppetId, "connected");
		});
		client.on("message", async (msg: Message) => {
			try {
				const dedupeKey = `${puppetId};${msg.chat.id}`;

					if(await this.messageDeduplicator.dedupe(dedupeKey, msg.author.id, msg.id, msg.text || "")) {
					return;
				}
				await this.handleTeamsMessage(puppetId, msg);
			} catch (err) {
				log.error("Error handling message event", err);
			}
		});
		client.on("messageChanged", async (msg: Message ) => {
			try {
				log.verbose("Got new message changed event");
				await this.handleTeamsMessageChanged(puppetId, msg);
			} catch (err) {
				log.error("Error handling teams messageChanged event", err);
			}
		});
		client.on("messageDeleted", async (msg: Message ) => {
			try {
				log.verbose("Got new message deleted event");
				await this.handleTeamsMessageDeleted(puppetId, msg);
			} catch (err) {
				log.error("Error handling teams messageDeleted event", err);
			}
		});
		this.puppet.AS.expressAppInstance.post(`/${puppetId}/chatSub`, client.incomingMessage.bind(client));

		this.puppets[puppetId] = {
			client,
			data,
			clientStopped: false,
		};
		try {
			await client.init();
			this.puppet.setUserId(puppetId, data.userId);
		}
		catch (err) {
			log.error("Error starting puppet client", err);
		}
	}

	public async listUsers(puppetId: number): Promise<IRetList[]> {
		const p = this.puppets[puppetId];
		if (!p) {
			return [];
		}
		const reply: IRetList[] = [];
		for (const [,user] of p.client.users) {
			reply.push({
				id: user.id,
				name: user.name,
			});
		}
		return reply;
	}

	public async getDmRoomId(user: IRemoteUser): Promise<string | null> {
		const p = this.puppets[user.puppetId];
		if (!p) {
			return null;
		}

		const room = [...p.client.chats].filter(([,c]) => c.members.has(user.userId))[0][1];
		return room ? room.id : null;
	}

	public async createRoom(room: IRemoteRoom): Promise<IRemoteRoom | null> {
		const p = this.puppets[room.puppetId];
		if (!p) {
			return null;
		}
		log.info(`Received create request for chat update puppetId=${room.puppetId} roomId=${room.roomId}`);
		const chat: Chat | undefined = p.client.chats.get(room.roomId);
		if (!chat) {
			log.warn(`No matching room for ${room.roomId} `);
			return null;
		}

		// Subscriptions
		if (chat.subscriptionId == "" || chat.subscriptionExpiry < new Date()) {
			await p.client.createSubscription(chat);
		}

		return {
			puppetId: room.puppetId,
			roomId: room.roomId,
			isDirect: true,
			name: chat.name + " (Teams)"
		}
	}

	public async getUserIdsInRoom(room: IRemoteRoom): Promise<Set<string> | null> {
		const p = this.puppets[room.puppetId];
		if (!p) {
			return null;
		}
		const chan = p.client.chats.get(room.roomId);
		if (!chan) {
			return null;
		}
		const users = new Set<string>();
		
		for (const [, member] of chan.members) {
			users.add(member.id);
		}
		return users;
	}

	public async handleTeamsMessage(puppetId: number, msg: Message) {
		if (!msg.text) {
			return; // nothing to do
		}

		const params = await this.getSendParams(puppetId, msg);
		const client = this.puppets[puppetId].client;
		log.verbose("Received message.");
		const dedupeKey = `${puppetId};${params.room.roomId}`;

		if (!await this.messageDeduplicator.dedupe(dedupeKey, params.user.userId, params.eventId, msg.text || "")) {
			const opts: IMessageEvent = {
				formattedBody: msg.text.replace(/<[\/]?div>/g, "").replace(/<span[^>]*><img*[^>]+alt="([^\"]*)"[^>]*><\/span>/g, "$1"),
				body: htmlToFormattedText(msg.text),
				emote: false
			};
			await this.puppet.sendMessage(params, opts);
		}

	}

	public async handleTeamsMessageChanged(puppetId: number, msg: Message) {
		if (!msg.text) {
			msg.text = "";
		}
		const params = await this.getSendParams(puppetId, msg);
		const client = this.puppets[puppetId].client;
		log.verbose("Received message.");
		const dedupeKey = `${puppetId};${params.room.roomId}`;

		if (!await this.messageDeduplicator.dedupe(dedupeKey, params.user.userId, params.eventId, msg.text || "")) {
			const opts: IMessageEvent = {
				formattedBody: msg.text.replace(/<[\/]?div>/g, "").replace(/<span[^>]*><img*[^>]+alt="([^\"]*)"[^>]*><\/span>/g, "$1"),
				body: htmlToFormattedText(msg.text),
				emote: false
			};
			await this.puppet.sendEdit(params, msg.id, opts);
		}
	}

	public async handleTeamsMessageDeleted(puppetId: number, msg: Message) {
		const params = await this.getSendParams(puppetId, msg);
		await this.puppet.sendRedact(params, msg.id);
	}

	public async getSendParams(
		puppetId: number,
		msg: Message
	): Promise<IReceiveParams> {
		let eventId: string | undefined;
		eventId = msg.id;

		return {
			room: await this.getRoomParams(puppetId, msg.chat),
			user: {
				puppetId,
				userId: msg.author.id,
				name: msg.author.displayName
			},
			eventId: msg.id
		};
	}

	public async getRoomParams(puppetId: number, chan: Chat): Promise<IRemoteRoom> {
		return {
			puppetId,
			roomId: chan.id,
			name: chan.name + " (Teams)",
			isDirect: true,
		};
	}

	public async handleMatrixMessage(room: IRemoteRoom, data: IMessageEvent, asUser: ISendingUser | null, event: any) {
		const p = this.puppets[room.puppetId];
		if (!p) {
			return;
		}
		const chat = p.client.chats.get(room.roomId);
		if (!chat) {
			log.warn(`Room ${room.roomId} not found!`);
			return;
		}

		const eventId = await p.client.sendMessage(chat, event.content.formatted_body ?? event.content.body);
		if (eventId) {
			await this.puppet.eventSync.insert(room, data.eventId!, eventId);
		}
	}

	public async handleMatrixImage(
		room: IRemoteRoom,
		data: IFileEvent,
		asUser: ISendingUser | null,
		event: any,
	) {
		const p = this.puppets[room.puppetId];
		if (!p) {
			return;
		}
		const chat = p.client.chats.get(room.roomId);
		if (!chat) {
			log.warn(`Room ${room.roomId} not found!`);
			return;
		}

		const eventId = await p.client.sendImage(chat, data);
		if (eventId) {
			await this.puppet.eventSync.insert(room, data.eventId!, eventId);
		}
	}

	public async handleMatrixFile(
		room: IRemoteRoom,
		data: IFileEvent,
		asUser: ISendingUser | null,
		event: any,
	) {
		const p = this.puppets[room.puppetId];
		if (!p) {
			return;
		}
		const chat = p.client.chats.get(room.roomId);
		if (!chat) {
			log.warn(`Room ${room.roomId} not found!`);
			return;
		}

		const eventId = await p.client.sendFile(chat, data);
		if (eventId) {
			await this.puppet.eventSync.insert(room, data.eventId!, eventId);
		}
	}

}