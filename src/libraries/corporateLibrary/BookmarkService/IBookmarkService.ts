import { Bookmark } from "../models/Bookmark";

/**
 * IBookmarkService interface
 * @interface Manage Bookmarks Service Interface
 */
export interface IBookmarkService {
    /**
     * AddBookmark function
     * @param {Bookmark} bookmark Bookmark model object
     * @returns {Promise<string>} string - Returns response string message
     */
    AddBookmark(bookmark: Bookmark): Promise<string>;

    /**
     * GetBookmarks function
     * @returns {Promise<Bookmark[]>} Bookmark[] - Returns list of all bookmarks
     */
    GetBookmarks(): Promise<Bookmark[]>;

    /**
     * GetBookmarksById function
     * @param {string} id Bookmark Id
     * @returns {Promise<Bookmark>} Bookmark - Returns the bookmark for the specific id
     */
    GetBookmarksById(id: string): Promise<Bookmark>;
}