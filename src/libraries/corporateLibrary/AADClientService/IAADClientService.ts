import { AadHttpClient } from "@microsoft/sp-http";

/**
 * IAADClientService interface
 * @interface AAD Http Client Service Interface
 */
export interface IAADClientService {
    /**
     * GetAADClient function
     * @param {string} appId Azure AD Application Id (Client ID) 
     * @returns {Promise<AadHttpClient>} AadHttpClient - Returns AADHttpClient Instance
     */
    GetAADClient(appId: string): Promise<AadHttpClient>;
}