import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IAADClientService } from "./IAADClientService";
import { AadHttpClient } from "@microsoft/sp-http";

/**
 * AADClientService class
 * @class Azure AD Http Client Service Class 
 */
export class AADClientService implements IAADClientService {
    private context: WebPartContext = undefined;

    /**
     * AADClientService constructor
     * @constructor Create an instance of AADClientService
     * @param {WebPartContext} context WebPart Context
     */
    constructor(context: WebPartContext) {
        this.context = context;
    }

    /**
     * GetAADClient function
     * @param {string} appId Azure AD Application Id (Client ID)
     * @returns {Promise<AadHttpClient>} AadHttpClient - Returns AADHttpClient Instance
     */
    public GetAADClient = (appId: string): Promise<AadHttpClient> => {
        var aadClient: AadHttpClient;

        return new Promise<AadHttpClient>(
            (resolve: (aadClient: AadHttpClient) => void, reject: (error: any) => void): void => {
                this.context.aadHttpClientFactory
                    .getClient(appId)
                    .then(
                        (client: AadHttpClient): void => {
                            aadClient = client;
                            resolve(aadClient);
                        },
                        (err) => reject(err)
                    );
            }
        );
    }
}