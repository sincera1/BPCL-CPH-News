import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import ErrorService from "./ErrorService";

export interface IAuditLog {
  action: string;
  operation: "GET" | "POST" | "PUT" | "DELETE";
  statusCode: number;
  listName: string;
  itemId?: number;
  endpoint?: string; 
  userRole?: string; 
} 
   
export default class AuditService {

  private _sp: SPFI;
  private _siteUrl: string;

  constructor(sp: SPFI, siteUrl: string) {
    this._sp = sp;
    this._siteUrl = siteUrl;
  }

  //  Get Public IP (fallback safe)
  private async getUserIP(): Promise<string> {
    try {
      const res = await fetch("https://api.ipify.org?format=json");
      const data = await res.json();
      return data.ip;
    } catch {
      return "Unknown";
    }
  }

  //  Get Current User
  private async getCurrentUser() {
    return await this._sp.web.currentUser();
  }

  //  Build API Endpoint
  private buildEndpoint(listName: string, itemId?: number): string {
    return itemId
      ? `/_api/web/lists/getbytitle('${listName}')/items(${itemId})`
      : `/_api/web/lists/getbytitle('${listName}')/items`;
  }

  //  Build Display URL (Clickable)
  private buildItemUrl(listName: string, itemId?: number): string {
    if (!itemId) return "";
    return `${this._siteUrl}/Lists/${listName}/DispForm.aspx?ID=${itemId}`;
  }

  //  MAIN LOGGER
  public async log(logData: IAuditLog): Promise<void> {
    try {

      const user = await this.getCurrentUser();
      const ip = await this.getUserIP();

      const endpoint = this.buildEndpoint(
        logData.listName,
        logData.itemId
      );

      const itemUrl = this.buildItemUrl(
        logData.listName,
        logData.itemId
      );

      await this._sp.web.lists
        .getByTitle("AuditLogs")
        .items.add({
          Title: logData.action,
          UserEmail: user.Email,
          UserName: user.Title,
          Operation: logData.operation,
          StatusCode: logData.statusCode,
          ItemID: logData.itemId || 0,
          ListName: logData.listName,
          IPAddress: ip,
          UserRole: logData.userRole || "User",

          APIEndpoint: {
            Url: `${this._siteUrl}${endpoint}`,
            Description: "API Endpoint"
          },

          ItemURL: itemUrl
            ? {
                Url: itemUrl,
                Description: "View Item"
              }
            : null,
        });

    } catch (error) {

      ErrorService.logError(
        "AuditService.log",
        error
      );

    }
  }
 
}
