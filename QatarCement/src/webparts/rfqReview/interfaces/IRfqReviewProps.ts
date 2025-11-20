import { IDropdownOption } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRfqReviewProps {
  wpproperties: any;
  context: WebPartContext;
}
export interface IRfqReviewWebPartProps {
  webpartTitle: string;
  PRDetailsListName: string;
  PRItemSpecficationsListName: string;
  vendorListName: string;
  FlowConnectionsListName: string;
}
export interface IRfqReviewState {
  modalOverlay: {
    isOpen: boolean;
    Text: string;
  },
  prNumber: string;
  department: string;
  priority: string;
  dueDate: string;
  prInitiator: string;
  businessJustification: string;
  itemDetails: IItemData[]
  vendorOptions: IDropdownOption[];
  masterid: string;
}
export interface IItemData {
  index: string;
  Description: string;
  ItemCode: string;
  Quantity: string;
  UOM: string;
  Title: string;
  vendors: string;
}