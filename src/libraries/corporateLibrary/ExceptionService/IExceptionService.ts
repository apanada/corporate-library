import { HttpRequestError } from "@pnp/odata";

export interface IExceptionService {
    LogException(ex: Error | HttpRequestError): Promise<void>;
}