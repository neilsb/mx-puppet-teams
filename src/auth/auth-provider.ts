import { Request, Response } from "express";
import { IRetData, Log } from "mx-puppet-bridge";
import { Config } from "../index";
import jwt_decode from "jwt-decode";
import "isomorphic-fetch"
import { IStoreToken, MSTeamsStore } from "../store";
import * as moment from "moment";
import * as urljoin from 'url-join';

const request = require('request')
const got = require('got');

const log = new Log("TeamsPuppet:auth_provider");

export class AuthProvider {

	public static tmpTokenStore: Array<any> = []

	constructor(private store: MSTeamsStore) { }

	public async getAccessToken(puppetId: number): Promise<string> {

		let token: IStoreToken;
		try {

			token = await this.store.getToken(puppetId);
		}
		catch (err) {
			log.error(`Unable to load token for puppet ${puppetId} :: ${err}`);
			return Promise.reject("Unable to load token");
		}

		// Check if token has expired
		if (moment.unix(token.accessExpiry) < moment()) {
			token = await this.refreshAccessToken(token);
		}
		else if (moment.unix(token.accessExpiry) < moment().add(-15, 'm')) {
			// If less than 15 mins left, kick off a background refresh
			this.refreshAccessToken(token);
		}

		return token.accessToken;
	}


	private async refreshAccessToken(currentToken: IStoreToken): Promise<IStoreToken> {

		// TODO: User Semaphore to prevent multiple refresh requests

		log.verbose("refreshAccessToken: Requesting new access token");

		const tokenRequestUrl = `${Config().oauth.endPoint}/token`;
		try {
			const response = await got.post(tokenRequestUrl, {
				form: {
					grant_type: 'refresh_token',
					client_id: Config().oauth.clientId,
					client_secret: Config().oauth.clientSecret,
					refresh_token: currentToken.refreshToken
				}
			}).json();

			// Update the store token, save and return
			currentToken.accessToken = response.access_token;
			currentToken.refreshToken = response.refresh_token;
			currentToken.accessExpiry = response.expires_on;

			this.store.storeToken(currentToken.puppetId, currentToken);

			return currentToken;

		} catch (error) {
			log.error("Unable to refresh token", error.response.body);
			return Promise.reject("Unable to refresh token");
		}
	}


	public async checkForNewAuthorization(puppetId: number, authCode: string): Promise<void> {

		if (!AuthProvider.tmpTokenStore[authCode]) {
			return;
		}

		const tokenData = AuthProvider.tmpTokenStore[authCode];

		let token: IStoreToken = {
			puppetId: puppetId,
			userId: (jwt_decode(tokenData.access_token) as any).oid,
			accessExpiry: tokenData.expires_on,
			refreshToken: tokenData.refresh_token,
			accessToken: tokenData.access_token,
			login: tokenData.not_before
		}

		await this.store.storeToken(puppetId, token);
	}


	public static oauthCallback = async (req: Request, res: Response) => {
		if (typeof req.query.code !== "string") {
			res.status(forbidden).send(getHtmlResponse("Error!!"));
			return;
		}

		const _accessCode = req.query.code;

		if (_accessCode) {
			const tokenRequestUrl = `${Config().oauth.endPoint}/token`;

			const tokenRequestBody = {
				grant_type: 'authorization_code',
				client_id: Config().oauth.clientId,
				client_secret: Config().oauth.clientSecret,
				code: _accessCode,
				redirect_uri: urljoin(Config().oauth.serverBaseUri, `/msteams/oauth`),
				resource: 'https://graph.microsoft.com'
			};

			request.post(
				{ url: tokenRequestUrl, form: tokenRequestBody },
				(err, httpResponse, body) => {
					if (!err) {
						let code = Math.random().toString(26).substr(2, 6);

						if (AuthProvider.tmpTokenStore[code] === undefined) {
							AuthProvider.tmpTokenStore[code] = JSON.parse(body);

							res.send(getHtmlResponse(code));
						} else {
							log.error("Clash of random codes - what are the chances... : ", AuthProvider.tmpTokenStore[code])
						}


					} else {
						// Probably throw an error?	
						log.error("Error retrieving acces code", err);
					}
				}
			);
		} else {
			// Probably throw an error?	
			log.error("Missing access code in oauth callback")
			res.send(getHtmlResponse("Error"));
		}
	}


	public static getDataFromStrHook = async (str: string): Promise<IRetData> => {
		const retData = {
			success: false,
		} as IRetData;
		if (!str) {
			retData.error = `Please specify the Auth Code you received after logging in.\n\nIf you have not logged in yet, please visit ${urljoin(Config().oauth.serverBaseUri, "/login")} `;
			return retData;
		}

		if (str.trim().length != 6) {
			retData.error = `The Auth code should be the 6 character code you received after logging in.\n\nIf you have not logged in yet, please visit ${urljoin(Config().oauth.serverBaseUri, "/login")} `;
			return retData;
		}

		if (AuthProvider.tmpTokenStore[str.trim()] === undefined) {
			retData.error = `The Auth code '${str.trim()}' is invalid, or has expired.`;
			return retData;
		}

		const token = AuthProvider.tmpTokenStore[str.trim()]

		// Check that token has required Scopes
		const scopes = token.scope.toLowerCase().split(" ");
		if (!(scopes.includes("chat.readwrite") && scopes.includes("chatmessage.read") && scopes.includes("user.read"))) {
			retData.error = `The received token for auth code '${str.trim()}' does not include the required scopes (Chat.ReadWrite, ChatMessage.Read and ChatMessage.Send).  Please check your Azure Application setup.`;
			return retData;
		}

		// Extract userId from Token
		let userId: string = "";
		try {
			userId = (jwt_decode(token.access_token) as any).oid;
		} catch {
			retData.error = `Unable to retrieve user id from token.`;
			return retData;
		}

		if (userId == "") {
			retData.error = `Token does not contain oid.`;
			return retData;
		}


		retData.data = {
			userId: userId,
			auth_code: str.trim(),
			access_token: token.access_token,
			refresh_token: token.refresh_token,
			login: token.not_before,
			expiry: token.expires_on
		}

		retData.success = true;
		return retData;
	};

}

const forbidden = 403;
const getHtmlResponse = (code) => `<!DOCTYPE html>
<html lang="en">
<head>
	<title>MS Teams Auth token</title>
	<style>
		body {
			margin-top: 20px;
			text-align: center;
		}
	</style>
</head>
<body>
	<h2>Your Auth Code is: ${code}</h2>
	<h4>Use it by talking to the MS Teams Puppet Bridge Bot by saying <pre>link ${code}</pre>
	</h4>
</body>
</html>
`;



export const getNewAccessToken = async (code: string): Promise<any> => {
	const tokenRequestUrl = urljoin(Config().oauth.endPoint, "/token");
	try {
		const response = await got.post(tokenRequestUrl, {
			form: {
				grant_type: 'refresh_token',
				client_id: Config().oauth.clientId,
				client_secret: Config().oauth.clientSecret,
				refresh_token: code
			}
		}).json();
		return response;
	} catch (error) {
		log.error("Error getting access token", error.response.body);
	}

}

export const authSuccess = function (req, res) {
	res.redirect('/');
};