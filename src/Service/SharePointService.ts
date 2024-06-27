/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPFI } from "@pnp/sp";
import { ISharePointService } from "./ISharePointService";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { getGraphClient, getSP } from "./pnpjsConfig";
import "@pnp/sp/files";
import { IItemAddResult } from "@pnp/sp/items";

export class SharePointService implements ISharePointService {
  private _sp: SPFI;
  private _graphClient: MSGraphClientV3;
  constructor() {
    this._sp = getSP();
    getGraphClient()
      .then((value: any) => {
        this._graphClient = value;
        console.log(this._graphClient);
      })
      .catch((error) => console.error(error));
  }

  public getListItems = async (
    listTitle: string,
    orderBy: string
  ): Promise<any> => {
    const ls = this._sp.web.lists
      .getByTitle(listTitle)
      .items.top(4000)
      .orderBy(orderBy)
      .getPaged();
    return await ls;
  };

  public getFilteredListItems = async (
    listName: string,
    filterQuery: string,
    orderBy: string,
    selects: string[],
    expands: string[]
  ): Promise<any> => {
    const ls = this._sp.web.lists
      .getByTitle(listName)
      .items.select(...selects)
      .filter(filterQuery)
      .expand(...expands)
      .top(4000)
      .orderBy(orderBy, false)
      .getPaged();
    return await ls;
  };

  public getFileBlob = async (path: string): Promise<Blob> => {
    return await this._sp.web.getFileByServerRelativePath(path).getBlob();
  };

  public saveItemToList = async(listName: string, body: any): Promise<IItemAddResult> =>{
    return await this._sp.web.lists.getByTitle(listName).items.add(body);
  }
}
