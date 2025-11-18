import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IQatarCementDashboardProps {
  wpproperties: any;
  context: WebPartContext;
}
export interface IQatarCementDashboardWebPartProps {
  webpartTitle: string;
  PRDetailsListName: string;
  PRItemSpecficationsListName: string;
}
export interface IQatarCementDashboardState {
  modalOverlay: {
    isOpen: boolean;
    Text: string;
  };
}