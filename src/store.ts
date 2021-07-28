import { Store } from "mx-puppet-bridge";

const CURRENT_SCHEMA = 1;

export interface IStoreToken {
	puppetId: number,
	accessToken: string,
	refreshToken: string,
	userId: string,
	login: number,
	accessExpiry: number
}

export class MSTeamsStore {
	constructor(
		private store: Store,
	) { }

	public async init(): Promise<void> {
		await this.store.init(CURRENT_SCHEMA, "teams_schema", (version: number) => {
			return require(`./db/schema/v${version}.js`).Schema;
		}, false);
	}

	public async getToken(puppetId: number): Promise<IStoreToken> {
		const rows = await this.store.db.All("SELECT * FROM teams_tokenstore WHERE puppet_id = $p", { p: puppetId });
		let ret: IStoreToken;
		if (rows.length == 1) {
			return {
				puppetId,
				accessToken: rows[0].access_token as string,
				refreshToken: rows[0].refresh_token as string,
				userId: rows[0].user_id as string,
				accessExpiry: rows[0].access_expiry as number,
				login: rows[0].login as number
			}
		}
		return Promise.reject("Token not found");
	}

	public async storeToken(puppetId: number, token: IStoreToken) {
		const exists = await this.store.db.Get("SELECT 1 FROM teams_tokenstore WHERE puppet_id = $p AND user_id = $u",
			{ p: puppetId, u: token.userId });
		
		let sql: string = `INSERT INTO teams_tokenstore (
			puppet_id, access_token, refresh_token, user_id, login, access_expiry
		) VALUES (
			$puppetId, $accessToken, $refreshToken, $userId, $login, $accessExpiry
		)`;

		if (exists) {
			sql = `UPDATE teams_tokenstore SET
				access_token = $accessToken, 
				refresh_token = $refreshToken,
				login = $login,
				access_expiry = $accessExpiry
				WHERE
					puppet_id = $puppetId
				AND
					user_id = $userId`;
		}
		await this.store.db.Run(sql, {
			puppetId,
			accessToken: token.accessToken,
			refreshToken: token.refreshToken,
			userId: token.userId,
			login: token.login,
			accessExpiry: token.accessExpiry
		});
	}
}