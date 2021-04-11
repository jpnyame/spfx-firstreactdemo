import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IDropdownOption } from "office-ui-fabric-react";

export default class SPOperations {

    public GetAllLists(context:WebPartContext): Promise<IDropdownOption[]>{
      let restApiUrl:string = context.pageContext.web.absoluteUrl + "/_api/web/lists?select=Title";
      let listTitles: IDropdownOption[] = [];

      return new Promise<IDropdownOption[]>(async(resolve,reject) => {
        context.spHttpClient.get(restApiUrl,SPHttpClient.configurations.v1)
        .then((response:SPHttpClientResponse) => {
          response.json().then((results:any) => {
              console.log(results);
              results.value.map((result:any) => {
                  listTitles.push({key:result.Title, text:result.Title});
              });
          });
          resolve(listTitles);
        },(error:any):void => {
          reject("error occured " + error);
        });
      });
    }
}
