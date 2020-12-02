import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AadHttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { ILogListener, Logger, LogLevel } from "@pnp/logging";

import { AILogListener } from "../AILogListener/AILogListener";
import { IExceptionService } from "../ExceptionService/IExceptionService";
import { ExceptionService } from "../ExceptionService/ExceptionService";
import { Bookmark } from "../models/Bookmark";
import { IBookmarkService } from "./IBookmarkService";
import { IAADClientService } from "../AADClientService/IAADClientService";
import { AADClientService } from "../AADClientService/AADClientService";

import * as AppConfiguration from "../../../config.json";

/**
 * BookmarkService class
 * @class Manage Bookmarks Service Class
 */
export class BookmarkService implements IBookmarkService {
    private readonly context: WebPartContext = undefined;
    private bookmarksClient: AadHttpClient = undefined;
    private exceptionService: IExceptionService;

    /**
     * BookmarkService constructor
     * @constructor Creates an instance of BookmarkService
     * @param {WebPartContext} context WebPart context
     */
    constructor(context: WebPartContext) {
        this.context = context;

        const aiLogListener: ILogListener = new AILogListener(
            AppConfiguration.ApplicationInsightsInstrumentationKey,
            this.context.pageContext.user.email,
            "ManageBookmarks", "1.0.0.0"
        );
        Logger.activeLogLevel = LogLevel.Info;
        Logger.subscribe(aiLogListener);

        this.exceptionService = new ExceptionService(aiLogListener);

        let aadClientService: IAADClientService = new AADClientService(this.context);
        aadClientService.GetAADClient(AppConfiguration.AzureAdAppCliendId)
            .then(
                (client: AadHttpClient): void => {
                    this.bookmarksClient = client;
                    Logger.write("Created new instance of AadHttpClient in BookmarkService constructor", LogLevel.Info);
                },
                (err) => {
                    Logger.write("Exception ocurred in getting AadHttpClient instance", LogLevel.Error);
                    this.exceptionService.LogException(err);
                }
            )
            .catch((err) => {
                Logger.write("Exception ocurred in getting AadHttpClient instance", LogLevel.Error);
                this.exceptionService.LogException(err);
            });
    }

    /**
     * GetBookmarks function
     * @returns {Promise<Bookmark[]>} Bookmark[] - Returns list of all bookmarks
     */
    public GetBookmarks = async (): Promise<Bookmark[]> => {
        const apiUrl: string = `${AppConfiguration.ApiBaseUrl}/api/GetBookmarks?code=KLofLip41yhwRGLh52q9sabeoi7nJxpKVZ9Ds3OSQwtWJFPaV5mqyw==`;

        try {
            Logger.write(`Calling GetBookmarks api for apiUrl: ${apiUrl}`, LogLevel.Info);

            // Get the response
            const response: HttpClientResponse = await this.bookmarksClient
                .get(apiUrl, AadHttpClient.configurations.v1);

            if (response.ok) {
                Logger.write(`Received bookmarks successfully with status ${response.status}`);

                // Read the value from the JSON
                const bookmarks: any = await response.json();

                // Return the value
                return bookmarks.map(
                    (bookmark: any) => ({ Id: bookmark.id, Url: bookmark.url })
                );
            }
        }
        catch (err) {
            Logger.write("Exception ocurred in calling GetBookmarks function", LogLevel.Error);
            this.exceptionService.LogException(err);
        }

        return [] as Bookmark[];
    }

    /**
     * GetBookmarksById function
     * @param {string} id Bookmark Id
     * @returns {Promise<Bookmark>} Bookmark - Returns the bookmark for the specific id
     */
    public GetBookmarksById = async (id: string): Promise<Bookmark> => {
        const apiUrl: string = `${AppConfiguration.ApiBaseUrl}/api/GetBookmarks/${id}?code=4anZC7EuJ4NCIZS4BNDSazGFaBpDHTFkYTcQZvMQHFfagsfsqan2kA==`;

        try {
            Logger.write(`Calling GetBookmarksById api for apiUrl: ${apiUrl}`, LogLevel.Info);

            // Get the response
            const response: HttpClientResponse = await this.bookmarksClient
                .get(apiUrl, AadHttpClient.configurations.v1);

            if (response.ok) {
                Logger.write(`Received bookmark for Id: ${id} successfully with status ${response.status}`);

                // Read the value from the JSON
                const bookmark: any = await response.json();

                // Return the value
                return <Bookmark>{ Id: bookmark.id, Url: bookmark.url };
            }
        }
        catch (err) {
            Logger.write(`Exception ocurred in calling GetBookmarksById function with Id: ${id}`, LogLevel.Error);
            this.exceptionService.LogException(err);
        }

        return undefined as Bookmark;
    }

    /**
     * AddBookmark function
     * @param {Bookmark} bookmark Bookmark model object
     * @returns {Promise<string>} string - Returns response string message
     */
    public AddBookmark = async (bookmark: Bookmark): Promise<string> => {
        const apiUrl: string = `${AppConfiguration.ApiBaseUrl}/api/AddBookmark?code=LoHqdQaRLf2mlYjf4L81jf1l4gwrnMkOivr6IrJs5Wi3Qs82GoOadw==`;

        try {
            Logger.write(`Calling AddBookmark api for apiUrl: ${apiUrl}`, LogLevel.Info);

            // Setup the options with header and body
            const headers: Headers = new Headers({
                'Content-type': 'application/json',
                'Access-Control-Allow-Origin': '*',
                'Accept': 'application/json'
            });

            const newBookmark: any = {
                id: bookmark.Id,
                url: bookmark.Url
            };

            const postOptions: IHttpClientOptions = {
                headers: headers,
                body: JSON.stringify(newBookmark)
            };

            // Get the response
            const response: any = await this.bookmarksClient
                .post(apiUrl, AadHttpClient.configurations.v1, postOptions);

            if (response.ok) {
                Logger.write(`Added bookmark successfully with status ${response.status}`);

                // Read the value from the response
                const responseText: string = await response.text();

                return responseText;
            }
        }
        catch (err) {
            Logger.write(`Exception ocurred in calling AddBookmark function with Bookmark: ${JSON.stringify(bookmark)}`, LogLevel.Error);
            this.exceptionService.LogException(err);
        }

        return undefined as string;
    }
}