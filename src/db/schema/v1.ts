import { IDbSchema, Store } from "mx-puppet-bridge";

export class Schema implements IDbSchema {
	public description = "Schema, Tokenstore";
	public async run(store: Store) {
		await store.createTable(`
			CREATE TABLE teams_schema (
				version	INTEGER UNIQUE NOT NULL
			);`, "teams_schema");
		await store.db.Exec("INSERT INTO teams_schema VALUES (0);");
		await store.createTable(`
			CREATE TABLE teams_tokenstore (
				puppet_id INTEGER NOT NULL,
				access_token TEXT NOT NULL,
				refresh_token TEXT NOT NULL,
				user_id TEXT NOT NULL,
				login INTEGER NOT NULL,
				access_expiry INTEGER NOT NULL
			);`, "teams_tokenstore");
		await store.createTable(`
			CREATE TABLE teams_subscriptions (
				puppet_id INTEGER NOT NULL,
				room_id TEXT NOT NULL,
				subscription_id TEXT NOT NULL,
				expiry INTEGER NOT NULL
			);`, "teams_subscriptions");
	}
	public async rollBack(store: Store) {
		await store.db.Exec("DROP TABLE IF EXISTS teams_schema");
		await store.db.Exec("DROP TABLE IF EXISTS teams_tokenstore");
	}
}
