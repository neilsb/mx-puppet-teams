import {
	PuppetBridge,
	IProtocolInformation,
	IPuppetBridgeRegOpts,
	Log,
} from "mx-puppet-bridge";

import * as commandLineArgs from "command-line-args";
import * as commandLineUsage from "command-line-usage";
import * as fs from "fs";
import * as yaml from "js-yaml";
import { AuthProvider } from "./auth/auth-provider";
import { App } from "./teams";
import { TeamsConfigWrap } from "./config";
const urljoin = require('url-join');

const log = new Log("MSTeamsPuppet:index");

const commandOptions = [
	{ name: "register", alias: "r", type: Boolean },
	{ name: "registration-file", alias: "f", type: String },
	{ name: "config", alias: "c", type: String },
	{ name: "help", alias: "h", type: Boolean },
];
const options = Object.assign({
	"register": false,
	"registration-file": "msteams-registration.yaml",
	"config": "config.yaml",
	"help": false,
}, commandLineArgs(commandOptions));

if (options.help) {
	console.log(commandLineUsage([
		{
			header: "Matrix Microsft Teams (Chat) Puppet Bridge",
			content: "A matrix puppet bridge for chats in Microsoft Teams",
		},
		{
			header: "Options",
			optionList: commandOptions,
		},
	]));
	process.exit(0);
}

const protocol = {
	features: {
		image: true,
		file: true,
		presence: false,
		edit: true,
		reply: true,
		globalNamespace: true,
	},
	id: "msteams",
	displayname: "MS Teams",
	externalUrl: "https://teams.microsoft.com/",
} as IProtocolInformation;

const puppet = new PuppetBridge(options["registration-file"], options.config, protocol);

if (options.register) {
	// okay, all we have to do is generate a registration file
	puppet.readConfig(false);
	try {
		puppet.generateRegistration({
			prefix: "_teamspuppet_",
			id: "teams-puppet",
			url: `http://${puppet.Config.bridge.bindAddress}:${puppet.Config.bridge.port}`,
		} as IPuppetBridgeRegOpts);
	} catch (err) {
		console.log("Couldn't generate registration file:", err);
	}
	process.exit(0);
}

let config: TeamsConfigWrap = new TeamsConfigWrap();

function readConfig() {
	config = new TeamsConfigWrap();
	config.applyConfig(yaml.safeLoad(fs.readFileSync(options.config)));
}

export function Config(): TeamsConfigWrap {
	return config;
}

export function Puppet(): PuppetBridge {
	return puppet;
}

async function run() {
	await puppet.init();
	readConfig();
	const teams = new App(puppet);
	await teams.init();
	puppet.on("puppetNew", teams.newPuppet.bind(teams));
	puppet.on("puppetDelete", teams.deletePuppet.bind(teams));
	puppet.on("message", teams.handleMatrixMessage.bind(teams));
	puppet.on("image", teams.handleMatrixImage.bind(teams));
	puppet.setCreateRoomHook(teams.createRoom.bind(teams));
	puppet.setGetDmRoomIdHook(teams.getDmRoomId.bind(teams));
	puppet.setListUsersHook(teams.listUsers.bind(teams));
	puppet.setGetUserIdsInRoomHook(teams.getUserIdsInRoom.bind(teams));
	puppet.setGetDataFromStrHook(AuthProvider.getDataFromStrHook);
	puppet.setBotHeaderMsgHook((): string => {
		return "MS Teams Puppet Bridge";
	});

	puppet.AS.expressAppInstance.get('/login', function (req, res) {
		const redirectUrl = urljoin(Config().oauth.serverBaseUri, Config().oauth.redirectPath);
		const authUrl = urljoin(Config().oauth.endPoint, "/authorize");
		res.redirect(`${authUrl}?response_type=code&redirect_uri=${encodeURI(redirectUrl)}&client_id=${Config().oauth.clientId}`);
	});

	puppet.AS.expressAppInstance.get(Config().oauth.redirectPath, AuthProvider.oauthCallback);

	// and finally, we start the puppet
	await puppet.start();
}

run(); // start the bridge
