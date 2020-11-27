import { hOP } from "@pnp/common";
import { ILogListener, Logger, LogLevel } from "@pnp/logging";
import { HttpRequestError } from "@pnp/odata";
import { IExceptionService } from "./IExceptionService";

export class ExceptionService implements IExceptionService {
    constructor(aiLogListener: ILogListener) {
        Logger.activeLogLevel = LogLevel.Info;
        Logger.subscribe(aiLogListener);
    }

    public LogException = async (ex: Error | HttpRequestError): Promise<void> => {
        //Checks to see if the error object has a property called isHttpRequestError. Returns a bool.
        if (hOP(ex, "isHttpRequestError")) {
            const data = await (<HttpRequestError>ex).response.clone().json();
            const message = typeof data["odata.error"] === "object" ? data["odata.error"].message.value : ex.message;
            const level: LogLevel = LogLevel.Error;

            Logger.log({ data, level, message });

        } else {
            // not an HttpRequestError so we just log message
            Logger.error(ex);
        }
    }
}