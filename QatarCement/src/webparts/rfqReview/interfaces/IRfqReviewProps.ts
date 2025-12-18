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
  workflowTaskListName: string;
}

export interface IVendorResponse {
  status?: string;
  price?: string;
  comments?: string; 
}

export interface IRfqReviewState {
  modalOverlay: {
    isOpen: boolean;
    Text: string;
  },
  currentUser: IUser;
  userType: string;
  prNumber: string;
  department: string;
  priority: string;
  dueDate: string;
  prInitiator: string;
  businessJustification: string;
  itemDetails: IItemData[]
  vendorOptions: IDropdownOption[];
  masterid: string;
  taskID: any;
  vendorResponses: Record<number, IVendorResponse>;

}


export interface IItemData {
  index: string;
  Id: number;
  Description: string;
  ItemCode: string;
  Quantity: string;
  UOM: string;
  Title: string;
  vendors: string;
  WorkflowDetailsId: number;
}
export interface IUser {
  id: any;
  email: string;
  title: string;
}