import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import * as moment from "moment";
import { Log } from "mx-puppet-bridge";
import { AuthProvider } from "./auth-provider";
import jwt_decode from "jwt-decode";

const log = new Log("TeamsPuppet:teams_auth_provider");

/*
 * Auth Provider for MS Graph client
 */
export class TeamsAuthProvider implements AuthenticationProvider {

    private accessToken: string = "";
    private tokenExpiry: moment.Moment;

    constructor(private puppetId: number, private authProvider: AuthProvider) { }

    /**
     * This method will get called before every request to the msgraph server
     * This should return a Promise that resolves to an accessToken (in case of success) or rejects with error (in case of failure)
     * Basically this method will contain the implementation for getting and refreshing accessTokens
     */
    public async getAccessToken(): Promise<string> {

        if (this.accessToken.length > 0 && this.tokenExpiry > moment()) {
            return this.accessToken;
        }

        // Get access token from auth provider
        try {
            this.accessToken = await this.authProvider.getAccessToken(this.puppetId);

            // Get expiry info
            this.tokenExpiry = moment.unix(jwt_decode<any>(this.accessToken).exp);

            return this.accessToken;
        }
        catch (err) {
            log.error(`Unable to get access token for puppet ${this.puppetId} :: ${err}`);
            return Promise.reject("Unable to get access token");
        }
    }
}