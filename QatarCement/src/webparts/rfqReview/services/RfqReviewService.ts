import { BaseService } from "../../../shared/services/BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../shared/PnP/pnpjsConfig";

export class RfqReviewService extends BaseService {
  private spfi: SPFI;
  constructor(context: WebPartContext, siteUrl: string) {
    super(context, siteUrl);
    this.spfi = getSP(context);
  }

  public getCurrentUser() {
    return this.spfi.web.currentUser();
  }
  public async isUserInGroup(groupName: string): Promise<boolean> {
    try {
      const groups = await this.spfi.web.currentUser.groups();

      return groups.some(g => g.Title.toLowerCase() === groupName.toLowerCase());
    } catch (error) {
      console.error("Error checking group membership:", error);
      return false;
    }
  }
}