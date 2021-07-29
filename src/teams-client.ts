import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { EventEmitter } from "events";
import * as moment from "moment";
import { Config } from "./index";
import { TeamsAuthProvider } from "./auth/teams-auth-provider";
import { Log } from "mx-puppet-bridge";
import * as urljoin from 'url-join';

const log = new Log("TeamsPuppet:teams-client");


export declare class Chat {
    id: string;
    name: string;
    members: Map<string, User>
    subscriptionId: string;
    subscriptionExpiry: Date;
}

export declare class User {
    id: string;
    name: string;
    displayName: string;
    constructor(id: string, name: string, displayName: string);
}

export declare class Message {
    id: string;
    chat: Chat;
    author: User;
    text: string | null;
}

export interface IClientOpts {
    authProvider: TeamsAuthProvider;
    ownerUserId: string;
    puppetId: number;
    knownMessage: (r: string, e: string) => Promise<boolean>;
}


export class TeamsClient extends EventEmitter {

    private client: Client;
    private teamsAuthProvider: TeamsAuthProvider;
    private ownerUserId: string;
    private subManagement: NodeJS.Timeout;
    private newChatManagement: NodeJS.Timeout;
    private lastNewChatCheck: Date = new Date(0);

    constructor(private opts: IClientOpts) {
        super();
        this.teamsAuthProvider = opts.authProvider;
        this.ownerUserId = opts.ownerUserId
    }

    public chats: Map<string, Chat> = new Map<string, Chat>();
    public users: Map<String, User> = new Map<string, User>();

    public async init() {

        // Create middleware for MS Graph Client
        let clientOptions: ClientOptions = {
            authProvider: this.teamsAuthProvider
        };
        this.client = Client.initWithMiddleware(clientOptions);

        // Recent Chat check limit
        let checkLimit: Date = new Date();
        checkLimit.setHours(checkLimit.getHours() - (Config().teams.recentChatDays * 24));

        var latestChat = await this.LoadChats(checkLimit);
        this.lastNewChatCheck = latestChat ?? checkLimit;

        await this.loadSubscriptions();

        this.emit("connected");

        // Start subsciption Management
        this.subManagement = setInterval(this.subscriptionManagement.bind(this), 300000);
        this.subscriptionManagement();

        // Start new Check Polling 
        this.newChatManagement = setInterval(this.newChatPolling.bind(this), Config().teams.newChatPollingPeriod * 1000);
    }

    public async stop() {
        if (this.subManagement) {
            clearInterval(this.subManagement);
        }
        if (this.newChatManagement) {
            clearInterval(this.newChatManagement);
        }
    }

    private async subscriptionManagement() {

        log.verbose("Checking for any expired/expiring Subscriptions");

        const expiryLimit = new Date();
        expiryLimit.setMinutes(expiryLimit.getMinutes() + 20);

        await [...this.chats].filter(([, v]) => v.subscriptionExpiry < expiryLimit).forEach(async ([, chat]) => {
            if (chat.subscriptionExpiry > new Date()) {
                this.renewSubscription(chat);
            }
            else {
                this.createSubscription(chat);
            }
        })
    }


    private async newChatPolling() {
        const latestChat = await this.LoadChats(this.lastNewChatCheck);
        if (latestChat) {
            this.lastNewChatCheck = latestChat;
        }
    }


    private async LoadChats(modifiedSince: Date): Promise<Date | undefined> {

        try {

            let earliestChat: Date = new Date();
            let latestChat: Date = new Date();

            log.silly(`Loading all chats modified since ${modifiedSince.toISOString()} `);

            let uri = '/me/chats?$expand=members';

            do {

                let chats = await this.client.api(uri)
                    .version('beta')
                    .get();

                await chats.value.filter(x => x.chatType == "oneOnOne").forEach(async chat => {

                    const lastUpdated = new Date(chat.lastUpdatedDateTime);
                    if (lastUpdated < earliestChat) {
                        earliestChat = lastUpdated;
                    }

                    if (lastUpdated > latestChat) {
                        latestChat = lastUpdated;
                    }

                    if (lastUpdated < modifiedSince) {
                        return;
                    }

                    // If this chat is already loaded, ignore it
                    if (this.chats.has(chat.id)) {
                        return;
                    }

                    const members = new Map<string, User>();

                    let otherMember = chat.members.find(x => x.userId != this.ownerUserId);

                    // Need to manually load displayname for members in other tenants
                    if (!otherMember.displayName || !otherMember.email) {

                        try {

                            const chatMembers = await this.client.api(`/chats/${chat.id}/members/`)
                                .version('beta')
                                .get();

                            otherMember = chatMembers.value.filter(x => x.userId == otherMember.userId)[0];
                        }
                        catch (err) {
                            log.warning(`Unable to retrieve details for Chat Member ${otherMember.userId}, Ignoring chat ${chat.id}`);
                            return;
                        }
                    }

                    members.set(
                        otherMember.userId,
                        {
                            id: otherMember.userId,
                            displayName: otherMember.displayName,
                            name: otherMember.email
                        });


                    if (!this.users.has(otherMember.userId)) {
                        this.users.set(otherMember.userId,
                            {
                                id: otherMember.userId,
                                displayName: otherMember.displayName,
                                name: otherMember.email
                            });
                    }
                    const name = otherMember == null ? "unknown" : otherMember.displayName ?? "??";

                    this.chats.set(chat.id,
                        {
                            id: chat.id,
                            name: name,
                            members,
                            subscriptionId: "",
                            subscriptionExpiry: new Date(0)
                        });

                });

                uri = chats["@odata.nextLink"];

            } while (earliestChat > modifiedSince && uri)

            return latestChat;
        }
        catch (err) {
            log.error("Error loading chats", err);
        }
    }

    public async loadSubscriptions() {

        try {
            // Loading subscriptions
            const response = await this.client.api(`/subscriptions`)
                .version('beta')
                .get();

            response.value.forEach(sub => {

                const chat = this.chats.get(sub.resource.match(/\/chats\/([^\/]+)\/messages/)[1]);
                if (chat) {

                    if (sub.notificationUrl != urljoin(Config().oauth.serverBaseUri, `/${this.opts.puppetId}/chatSub`)) {
                        log.warn("Subscription set up for another puppet on this application", sub.notificationUrl);
                        return;
                    }

                    chat.subscriptionId = sub.id;
                    chat.subscriptionExpiry = new Date(sub.expirationDateTime);
                }

            });
        }
        catch (err) {
            log.error("Error loading subscriptions", err);
        }

    }

    public async loadMessages(chat: Chat, limit: number = 20): Promise<Message[]> {

        let messages = await this.client.api(`/chats/${chat.id}/messages?top=${limit}`)
            .version('beta')
            .get();


        const retVal: Message[] = [];

        await messages.value.forEach(msg => {

            let author: User = this.users[msg.from.user.id];

            if (!author) {
                author = <User>{
                    id: msg.from.user.id,
                    displayName: msg.from.user.displayName,
                    name: msg.from.user.displayName,
                };
                this.users.set(author.id, author);
            }

            retVal.push(<Message>{
                id: msg.id,
                text: msg.body.content,
                chat: chat,
                author
            });
        });
        return retVal;
    }

    public async createSubscription(chat: Chat): Promise<void> {

        if (chat.subscriptionExpiry > new Date()) {
            // already have a subscription
            return;
        }

        log.verbose("Creating subscription for " + chat.id);
        try {
            const response = await this.client.api(`/subscriptions`)
                .version('beta')
                .post({
                    "changeType": "created,updated,deleted",
                    "notificationUrl": urljoin(Config().oauth.serverBaseUri, `/${this.opts.puppetId}/chatSub`),
                    "resource": `/chats/${chat.id}/messages`,
                    "expirationDateTime": moment().add(1, 'h').toISOString(),
                    "clientState": "secretClientValue",
                    "latestSupportedTlsVersion": "v1_2"
                });

            chat.subscriptionId = response.id;
            chat.subscriptionExpiry = new Date(response.expirationDateTime);

        } catch (err) {
            log.error("Error subscribing ", err);
        }
    }

    public async renewSubscription(chat: Chat): Promise<void> {

        if (chat.subscriptionExpiry < new Date()) {
            // for expired subscription, re-create
            await this.createSubscription(chat);
            return;
        }

        try {
            const response = await this.client.api(`/subscriptions/${chat.subscriptionId}`)
                .version('beta')
                .patch({
                    "expirationDateTime": moment().add(1, 'h').toISOString(),
                });

            chat.subscriptionExpiry = new Date(response.expirationDateTime);

        } catch (err) {
            log.error("Error renewing subsciption ", err);
        }
    }

    public async incomingMessage(req, res) {
        if (req.query.validationToken) {
            log.debug(`Got validation token request: ${req.query.validationToken}`);
            res.send(req.query.validationToken);
        } else {

            // TODO: Validation (inc secret)

            // Process message
            try {
                for (let i = 0; i < req.body.value.length; i++) {

                    const mdetails = req.body.value[i].resource.match(/chats\('([^']*).*messages\('([^']*)/);

                    if (req.body.value[i].changeType == 'deleted') {
                        log.info("Got message delete event from Teams");
                        const chat = this.chats.get(mdetails[1]);

                        if (chat) {
                            // teams only gives us the deleted message id, not the oringinal author.  The bridge requires
                            // a user id, so just take the 1st member of the room as the redaction author
                            this.emit('messageDeleted', {
                                id: mdetails[2],
                                chat: chat,
                                author: chat.members.get([...chat.members.keys()][0]),
                                text: ""
                            } as Message);
                        }
                        return;
                    }

                    if (req.body.value[i].changeType == 'created') {
                        if (mdetails.length == 3 && await this.opts.knownMessage(mdetails[1], mdetails[2])) {
                            // This is a message created by bridge
                            log.debug(`Ignoring message ${mdetails[1]}/${mdetails[2]} as it was created by us`);
                            continue;
                        }
                    }

                    // Process message from Teams
                    const r = await this.client.api(req.body.value[i].resource)
                        .version('beta')
                        .get();

                    const m = await this.parseTeamsMessage(r);
                    
                    if (m) {
                        if (req.body.value[i].changeType == 'created') {
                            this.emit('message', m); 
                        } else {
                            this.emit('messageChanged', m); 
                        }
                    }
                }
            }
            catch (err) {
                log.error("Error trying to retrieve message from teams", err);
            }

            res.send("OK");
        }
    }


    private async parseTeamsMessage(msg: any): Promise<Message | null> {

        let author: User = this.users[msg.from.user.id];

        if (!author) {
            author = <User>{
                id: msg.from.user.id,
                displayName: msg.from.user.displayName,
                name: msg.from.user.displayName,
            };
            this.users.set(author.id, author);
        }

        let chat: Chat | undefined = this.chats.get(msg.chatId);

        if (!chat) {
            console.log("Chat not found", msg.chatId);
            return null;
        }

        return {
            id: msg.id,
            text: msg.body.content,
            chat: chat,
            author
        };
    }

    public async sendMessage(room: Chat, msg: string): Promise<string> {

        try {
            const response = await this.client.api(`/chats/${room.id}/messages`)
                .version('beta')
                .post({
                    "body": {
                        "contentType": "html",
                        "content": msg,
                    }
                });
            return response.id;
        }
        catch (err) {
            log.error("Error sending message", err);
            return "";
        }
    }

}