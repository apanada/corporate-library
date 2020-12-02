import { LogLevel, ILogListener, ILogEntry, Logger } from "@pnp/logging";
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ReactPlugin, withAITracking } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from 'history';

import { _hashUser, _logEventFormat, _logMessageFormat } from "./Utils/Utilities";
import Constants from "./Utils/Constants";

/**
 * AILogListener Class
 * @class Application Insights Log Service Listner Class
 */
export class AILogListener implements ILogListener {
    private static _instrumentationKey: string;
    private static _webpartName: string;
    private static _webpartVersion: string;
    private static _appInsightsInstance: ApplicationInsights;
    private static _reactPluginInstance: ReactPlugin;

    private _constants: Constants;

    /**
     * AILogListener Constructor
     * @constructor Creates an instance of AILogListener
     * @param {string} instrumentationKey Application Insights InstrumentationKey
     * @param {string} currentUser Current SignIn User
     * @param {string} webpartName Webpart Name
     * @param {string} webpartversion Webpart Version
     */
    constructor(instrumentationKey: string, currentUser: string, webpartName: string, webpartVersion: string = "1.0.0.0") {
        AILogListener._instrumentationKey = instrumentationKey;
        AILogListener._webpartName = webpartName;
        AILogListener._webpartVersion = webpartVersion;

        this._constants = new Constants(AILogListener._webpartName);

        if (!AILogListener._appInsightsInstance) {
            AILogListener._appInsightsInstance = AILogListener.initializeApplicationInsights(currentUser);
        }
    }

    /**
     * initializeApplicationInsights function
     * @function Inializes and return the ApplicationInsights object
     * @param {string} currentUser Current SignIn User
     * @returns {ApplicationInsights} ApplicationInsights - ApplicationInsights Instance
     */
    private static initializeApplicationInsights = (currentUser?: string): ApplicationInsights => {
        try {
            if (!AILogListener._instrumentationKey) {
                throw new Error('Instrumentation key not provided');
            }

            const browserHistory = createBrowserHistory({ basename: '' });
            AILogListener._reactPluginInstance = new ReactPlugin();
            const appInsights = new ApplicationInsights({
                config: {
                    maxBatchInterval: 0,
                    instrumentationKey: AILogListener._instrumentationKey,
                    namePrefix: AILogListener._webpartName,             // Used as Postfix for cookie and localStorage 
                    disableFetchTracking: false,                        // To avoid tracking on all fetch
                    disableAjaxTracking: true,                          // Not to autocollect Ajax calls
                    autoTrackPageVisitTime: true,
                    extensions: [AILogListener._reactPluginInstance],
                    extensionConfig: {
                        [AILogListener._reactPluginInstance.identifier]: { history: browserHistory }
                    }
                }
            });

            appInsights.loadAppInsights();
            appInsights.trackPageView();
            appInsights.context.application.ver = AILogListener._webpartVersion;    // application_Version
            appInsights.setAuthenticatedUserContext(_hashUser(currentUser));        // user_AuthenticateId
            return appInsights;
        }
        catch (ex) {
            console.error(ex);
        }

        return undefined;
    }

    /**
     * getReactPluginInstance function
     * @returns {ReactPlugin} ReactPlugin - ReactPlugin Instance
     */
    public static getReactPluginInstance(): ReactPlugin {
        if (!AILogListener._reactPluginInstance) {
            AILogListener._reactPluginInstance = new ReactPlugin();
        }
        return AILogListener._reactPluginInstance;
    }

    /**
     * getAppInsights function
     * @returns {ApplicationInsights} ApplicationInsights - ApplicationInsights Instance
     */
    public static getAppInsights(): ApplicationInsights {
        if (!AILogListener._appInsightsInstance) {
            AILogListener._appInsightsInstance = AILogListener.initializeApplicationInsights();
        }
        return AILogListener._appInsightsInstance;
    }

    /**
     * trackEvent function
     * @param {string} name Event Name
     */
    public trackEvent(name: string): void {
        if (AILogListener._appInsightsInstance)
            AILogListener._appInsightsInstance.trackEvent(
                _logEventFormat(name),
                this._constants.ApplicationInsights.CustomProps
            );
    }

    /**
     * log function
     * @override ILogListener log method
     * @param {ILogEntry} entry ILogEntry Instance
     */
    public log(entry: ILogEntry): void {
        const msg = _logMessageFormat(entry);
        if (entry.level === LogLevel.Off) {
            // No log required since the level is Off
            return;
        }

        if (AILogListener._appInsightsInstance)
            switch (entry.level) {
                case LogLevel.Verbose:
                    AILogListener._appInsightsInstance.trackTrace({ message: msg, severityLevel: SeverityLevel.Verbose }, this._constants.ApplicationInsights.CustomProps);
                    break;
                case LogLevel.Info:
                    console.log({ ...this._constants.ApplicationInsights.CustomProps, Message: msg });
                    AILogListener._appInsightsInstance.trackTrace({ message: msg, severityLevel: SeverityLevel.Information }, this._constants.ApplicationInsights.CustomProps);
                    break;
                case LogLevel.Warning:
                    console.warn({ ...this._constants.ApplicationInsights.CustomProps, Message: msg });
                    AILogListener._appInsightsInstance.trackTrace({ message: msg, severityLevel: SeverityLevel.Warning }, this._constants.ApplicationInsights.CustomProps);
                    break;
                case LogLevel.Error:
                    console.error({ ...this._constants.ApplicationInsights.CustomProps, Message: msg });
                    AILogListener._appInsightsInstance.trackException({ exception: new Error(msg), severityLevel: SeverityLevel.Error, properties: this._constants.ApplicationInsights.CustomProps });
                    break;
            }
    }
}

export default (Component: any) =>
    withAITracking(AILogListener.getReactPluginInstance(), Component);