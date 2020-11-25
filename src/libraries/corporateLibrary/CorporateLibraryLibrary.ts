/**
 * CorporateLibraryLibrary class
 * @class Generic Service Library
 */
export class CorporateLibraryLibrary {
  public name(): string {
    return 'CorporateLibraryLibrary';
  }

  public getCurrentTime(): string {
    return 'The current time as returned from the corporate library is ' + new Date().toTimeString();
  }
}
