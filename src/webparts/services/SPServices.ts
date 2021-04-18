import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IDropdownOption } from "office-ui-fabric-react";

export default class SPOperations {

  public GetAllListsTitles(context: WebPartContext): Promise<IDropdownOption[]> {
    let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists?select=Title";
    let listTitles: IDropdownOption[] = [];

    return new Promise<IDropdownOption[]>(async (resolve, reject) => {
      context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          response.json().then((results: any) => {
            console.log(results);
            results.value.map((result: any) => {
              listTitles.push({ key: result.Title, text: result.Title });
            });
          });
          resolve(listTitles);
        }, (error: any): void => {
          reject("error occured " + error);
        });
    });
  }

  public CreateListItem(context: WebPartContext, listTitle: string): Promise<string> {
    let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listTitle + "')/items";
    const body: string = JSON.stringify({ Title: "New Item created" });
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=nometadata",
        "mode": "cors",
        "content-type": "application/json;odata=nometadata",
        "odata-version": ""
      },
      body: body
    };
    return new Promise<string>(async (resolve, reject) => {
      context.spHttpClient.post(restApiUrl, SPHttpClient
        .configurations.v1, options)
    });
  }
}
