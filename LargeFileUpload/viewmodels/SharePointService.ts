import { sp, containsInvalidFileFolderChars, IFolderAddResult } from "@pnp/sp/presets/all";
import { HttpRequestError } from "@pnp/odata";
import { Web, IWeb } from "@pnp/sp/webs";
import * as Msal from "msal";
import { SPFetchClient } from "@pnp/nodejs";
import { PnPFetchClient } from "../PnPFetchClient";


//sp.setup({
//    sp: {
//        fetchClientFactory: () => {
//            return new SPFetchClient("https://msftnbu.sharepoint.com/", "a80fe795-7a8c-49cb-8135-cac6230902fe", "dd453110-1f1c-462c-b6f0-26206a2ab1e5");
//        },
//    },
//});
export class SharePointService {
    static serviceProviderName = "SharePointService";
    context: ComponentFramework.Context<unknown>;
    config: {
        sharepointSiteId: string;
        sharePointStructureEntity: string;
        clientId: string;
        loginHint: string;
    };
    msalInstance: Msal.UserAgentApplication;
    web: IWeb;
    sharePointRelativeUrl: string;
    sharePointStructureEntity: string;
    sharePointAboluteUrl: string;
    msalConfig: Msal.Configuration;
    ssoRequest: Msal.AuthenticationParameters;
    constructor(
        context: ComponentFramework.Context<unknown>,
        config: {
            sharepointSiteId: string;
            sharePointStructureEntity: string;
            clientId: string;
            loginHint: string;
        },
    ) {
        this.context = context;
        this.config = config;
    }

    async setupSharePoint(
        sharePointAboluteUrl: string,
        sharePointStructureEntity: string,
    ): Promise<void> {
        const msalConfig = {
            auth: {
                clientId: "a80fe795-7a8c-49cb-8135-cac6230902fe",
                redirectURi: "https://orged9fe8b2.crm.dynamics.com/main.aspx",
                authority: "https://login.microsoftonline.com/common",
            },
            cache: {
                cahcheLocation: "localStorate",
                storeAuthStateInCookie: true, // Set this to "true" if you are having issues on IE11 or Edge
            },
        };
        const ssoRequest: Msal.AuthenticationParameters = {
            loginHint: "dev1@msftnbu.com"
                //this.config.loginHint
        };
        this.msalConfig = msalConfig;
        this.ssoRequest = ssoRequest;
        this.sharePointAboluteUrl = sharePointAboluteUrl;
        const msalInstance = new Msal.UserAgentApplication(msalConfig);
        await msalInstance.ssoSilent(ssoRequest).then((response) => {
            sp.setup({
                sp: {
                    // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
                    fetchClientFactory: () => {
                        return new PnPFetchClient(msalInstance);
                    },
                },
            });
        })
            .catch(error => {
                // handle error by invoking an interactive login method
                msalInstance.loginPopup(ssoRequest);
            });
        this.web = Web(sharePointAboluteUrl);
        this.sharePointStructureEntity = sharePointStructureEntity;
        this.sharePointRelativeUrl = sharePointAboluteUrl.replace(/^(?:\/\/|[^\/]+)*\//, "");
    }	

    async uploadFileToSharePoint(sharePointFolderName: string, file: any): Promise<void> {
        if (file.size <= 10485760) {
            await sp.web.getFolderByServerRelativePath(sharePointFolderName).files.add(file.name, file, true);
        } else {
            await sp.web.getFolderByServerRelativePath(sharePointFolderName).files.addChunked(
                file.name,
                file,
                (data) => {
                    console.log(data);
                },
                true,
            );
        }
    }

    async sharepointFolderExists(folderName: string): Promise<boolean> {
        try {
            const folderExists = await this.web.getFolderByServerRelativePath(folderName)();
            return true;
        } catch {
            console.log("SharePoint Folder Not Found - Will be Created");
            return false;
        }
    }

    async createSharePointFolder(folderName: string, newFolder: string): Promise<void> {
        try {
            const folder = await this.web.getFolderByServerRelativePath(folderName).folders.add(newFolder);
        } catch (e) {
            console.log(e);
            if (e?.isHttpRequestError) {

                // we can read the json from the response
                const json = await (<HttpRequestError>e).response.json();

                // if we have a value property we can show it
                console.log(typeof json["odata.error"] === "object" ? json["odata.error"].message.value : e.message);

                // add of course you have access to the other properties and can make choices on how to act
                if ((<HttpRequestError>e).status === 404) {
                    console.error((<HttpRequestError>e).statusText);
                    // maybe create the resource, or redirect, or fallback to a secondary data source
                    // just ideas, handle any of the status codes uniquely as needed
                }

            } else {
                // not an HttpRequestError so we just log message
                console.log(e.message);
            }
        }
    }
}