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

}