import {
    SPHttpClient,
    // SPHttpClientConfiguration,
    ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { SPItemsResponse } from "../interfaces/SPItemsResponse";

export class ListService {
    private readonly listTitle: string;
    private readonly spHttpClient: SPHttpClient;
    private readonly siteUrl: string;

    constructor(
        listTitle: string,
        spHttpClient: SPHttpClient,
        siteUrl: string
    ) {
        this.listTitle = listTitle;
        this.spHttpClient = spHttpClient;
        this.siteUrl = siteUrl;
    }

    public async getListItems(
        filter: string = ""
    ): Promise<SPItemsResponse> {
        const url = `${this.siteUrl}/_api/web/lists/getbytitle('${this.listTitle}')/items?${filter}`;
        const requestOptions: ISPHttpClientOptions = {
            headers: {
                Accept: "application/json",
            },
        };

        return this.spHttpClient
            .get(url, SPHttpClient.configurations.v1, requestOptions)
            .then((response) => {
                if (response.ok) {
                    return response.json();
                } else {
                    console.log(`${response.status}: ${response.statusText}`);
                    
                    return Promise.reject(
                        `${response.status}: ${response.statusText}`
                    );
                }
            })
            .catch((error) => {
                console.log(error);
                return Promise.reject(error);
            });
    }
}
