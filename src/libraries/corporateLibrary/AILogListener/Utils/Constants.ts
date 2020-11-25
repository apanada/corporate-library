import { _tenantName } from "./Utilities";

export default class Constants {
    private static _webpartName: string;

    constructor(webpartName: string) {
        Constants._webpartName = webpartName;
    }

    public get ApplicationInsights() {
        return {
            CustomProps: {
                Tenant: _tenantName(), App_Name: Constants._webpartName
            }
        };
    }
}