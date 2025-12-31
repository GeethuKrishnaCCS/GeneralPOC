// import { IDropdownOption } from "@fluentui/react";
// import { WebPartContext } from "@microsoft/sp-webpart-base";

// export interface IRfqReviewProps {
//   wpproperties: any;
//   context: WebPartContext;
// }
// export interface IRfqReviewWebPartProps {
//   webpartTitle: string;
//   PRDetailsListName: string;
//   PRItemSpecficationsListName: string;
//   vendorListName: string;
//   FlowConnectionsListName: string;
//   workflowTaskListName: string;
// }

// export interface IVendorResponse {
//   status?: string;
//   price?: string;
//   comments?: string; 
// }



// export interface IRfqReviewState {
//   modalOverlay: {
//     isOpen: boolean;
//     Text: string;
//   },
//   currentUser: IUser;
//   userType: string;
//   prNumber: string;
//   department: string;
//   priority: string;
//   dueDate: string;
//   prInitiator: string;
//   businessJustification: string;
//   itemDetails: IItemData[]
//   vendorOptions: IDropdownOption[];
//   masterid: string;
//   taskID: any;
//   vendorResponses: Record<number, IVendorResponse>;

// }


// export interface IItemData {
//   index: string;
//   Id: number;
//   Description: string;
//   ItemCode: string;
//   Quantity: string;
//   UOM: string;
//   Title: string;
//   vendors: string;
//   WorkflowDetailsId: number;
//   vendorResponses?: IVendorResponse[]; // ⭐ ADD THIS for initiator view
// }
// export interface IUser {
//   id: any;
//   email: string;
//   title: string;
// }


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
  WorkflowTasksListName: string; // ⭐ Fixed property name
}

// ⭐ For WorkflowDetails list items
export interface IVendorResponse {
  Id: number;
  Vendor: string;
  Status: string;
  Price: string;
  Comments: string;
  PRItemID: string;
}

// ⭐ For vendor input tracking
export interface IVendorResponseInput {
  status?: string;
  price?: string;
  comments?: string; 
}

// ⭐ For initiator technical review
export interface IInitiatorResponse {
  status?: string;
  comments?: string;
}

export interface IManagerResponse {
  status?: string;
  comments?: string;
}

export interface IProcurementManagerResponse {
  status?: string;
  comments?: string;
}

export interface IRfqReviewState {
  modalOverlay: {
    isOpen: boolean;
    Text: string;
  };
  currentUser: IUser;
  userType: string;
  prNumber: string;
  department: string;
  priority: string;
  dueDate: string;
  prInitiator: string;
  businessJustification: string;
  itemDetails: IItemData[];
  vendorOptions: IDropdownOption[];
  masterid: string;
  taskID: any;
  vendorResponses: Record<number, IVendorResponseInput>; 
  initiatorResponses: Record<number, IInitiatorResponse>; 
  managerResponses: Record<number, IManagerResponse>;
  procurementManagerResponses: Record<number, IProcurementManagerResponse>;
}

export interface IItemData {
  index: number;
  Id: number ;
  Description: string;
  ItemCode: string;
  Quantity: string;
  UOM: string;
  Title: string;
  vendors: string;
  WorkflowDetailsId?: number | null;
  
  // ✅ Add these new properties for Initiator view
  Vendor?: string;           // Individual vendor email (from WorkflowDetails)
  Price?: string | number;   // Vendor's quoted price
  Comments?: string;         // Vendor's comments
  TaskID?: number | string;  // Associated task ID
  
  // For vendor responses (grouped view - if you still need it)
  vendorResponses?: any[];
   InitiatorStatus?: string;
  InitiatorComments?: string;
  ManagerStatus?: string;
  ManagerComments?: string;
}

export interface IUser {
  id: any;
  email: string;
  title: string;
}