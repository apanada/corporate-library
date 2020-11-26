import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AadHttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { Bookmark } from "../models/Bookmark";
import { IBookmarkService } from "./IBookmarkService";
import { IAADClientService } from "../AADClientService/IAADClientService";
import { AADClientService } from "../AADClientService/AADClientService";
import { AppConfiguration } from "read-appsettings-json";

/**
 * BookmarkService class
 * @class Manage Bookmarks Service Class
 */
export class BookmarkService implements IBookmarkService {
    private readonly context: WebPartContext = undefined;
    private bookmarksClient: AadHttpClient = undefined;

    /**
     * BookmarkService constructor
     * @constructor Creates an instance of BookmarkService
     * @param {WebPartContext} context WebPart context
     */
    constructor(context: WebPartContext) {
        this.context = context;

        let aadClientService: IAADClientService = new AADClientService(this.context);
        aadClientService.GetAADClient(AppConfiguration.Setting().AzureAdAppCliendId)
            .then(
                (client: AadHttpClient): void => {
                    this.bookmarksClient = client;
                },
                (err) => console.log(err)
            )
            .catch((err) => console.log(err));
    }

    /**
     * GetBookmarks function
     * @returns {Promise<Bookmark[]>} Bookmark[] - Returns list of all bookmarks
     */
    public GetBookmarks = async (): Promise<Bookmark[]> => {
        const apiUrl: string = `${AppConfiguration.Setting().ApiBaseUrl}/api/GetBookmarks?code=KLofLip41yhwRGLh52q9sabeoi7nJxpKVZ9Ds3OSQwtWJFPaV5mqyw==`;

        // Get the response
        const response: HttpClientResponse = await this.bookmarksClient
            .get(apiUrl, AadHttpClient.configurations.v1);

        if (response.ok) {
            // Read the value from the JSON
            const bookmarks: any = await response.json();

            // Return the value
            return bookmarks.map(
                (bookmark: any) => ({ Id: bookmark.id, Url: bookmark.url })
            );
        }

        return [] as Bookmark[];
    }

    /**
     * GetBookmarksById function
     * @param {string} id Bookmark Id
     * @returns {Promise<Bookmark>} Bookmark - Returns the bookmark for the specific id
     */
    public GetBookmarksById = async (id: string): Promise<Bookmark> => {
        const apiUrl: string = `${AppConfiguration.Setting().ApiBaseUrl}/api/GetBookmarks/${id}?code=4anZC7EuJ4NCIZS4BNDSazGFaBpDHTFkYTcQZvMQHFfagsfsqan2kA==`;

        // Get the response
        const response: HttpClientResponse = await this.bookmarksClient
            .get(apiUrl, AadHttpClient.configurations.v1);

        if (response.ok) {
            // Read the value from the JSON
            const bookmark: any = await response.json();

            // Return the value
            return <Bookmark>{ Id: bookmark.id, Url: bookmark.url };
        }

        return undefined as Bookmark;
    }

    /**
     * AddBookmark function
     * @param {Bookmark} bookmark Bookmark model object
     * @returns {Promise<string>} string - Returns response string message
     */
    public AddBookmark = async (bookmark: Bookmark): Promise<string> => {
        const apiUrl: string = `${AppConfiguration.Setting().ApiBaseUrl}/api/AddBookmark?code=LoHqdQaRLf2mlYjf4L81jf1l4gwrnMkOivr6IrJs5Wi3Qs82GoOadw==`;

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
            // Read the value from the response
            const responseText: string = await response.text();

            return responseText;
        }

        return undefined as string;
    }
}