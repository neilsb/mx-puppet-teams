export class TeamsConfigWrap {
	public oauth: OAuthConfig = new OAuthConfig();
	public teams: TeamsConfig = new TeamsConfig();

	public applyConfig(newConfig: {[key: string]: any}, configLayer: {[key: string]: any} = this) {
		Object.keys(newConfig).forEach((key) => {
			if (configLayer[key] instanceof Object && !(configLayer[key] instanceof Array)) {
				this.applyConfig(newConfig[key], configLayer[key]);
			} else {
				configLayer[key] = newConfig[key];
			}
		});
	}
}

class OAuthConfig {
	public clientId = "";
	public clientSecret = "";
	public redirectPath = "";
	public serverBaseUri = "";
    public endPoint = "https://login.windows.net/common/oauth2"
}

class TeamsConfig {
    public recentChatDays: number = 30;
    public newChatPollingPeriod: number = 300;
	public path: string = "/_matrix/Teams/client";
}
