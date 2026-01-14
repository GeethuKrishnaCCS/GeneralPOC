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
  WorkflowTasksListName: string;
  DocumentLibraryName: string;
}

export interface IVendorResponse {
  Id: number;
  Vendor: string;
  Status: string;
  Price: string;
  Comments: string;
  PRItemID: string;
}

export interface IVendorResponseInput {
  status?: string;
  price?: string;
  comments?: string;
  termsAndConditions?: string;
  technicalSupport?: string;
  warrantySupport?: string;
  attachments?: File[];
}

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

export interface IAttachment {
  name: string;
  url: string;
  size?: number;
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
  selectedFiles: File[];
  uploadedFileUrls: string[];
  attachments: IAttachment[];
  isLoadingAttachments: boolean;
  commonManagerStatus: string;
  commonManagerComments: string;
  commonProcurementStatus: string;
  commonProcurementComments: string;
  workflowDetailsAttachments?: Record<number, Array<{ name: string; url: string }>>;
  vendorTermsAndConditions: string;
  vendorTechnicalSupport: string;
  vendorWarrantySupport: string;
} 

export interface IItemData {
  index: number;
  Id: number;
  Description: string;
  ItemCode: string;
  Quantity: string;
  UOM: string;
  Title: string;
  vendors: string;
  WorkflowDetailsId?: number | null;
  Vendor?: string;
  Price?: string | number;
  Comments?: string;
  TaskID?: number | string;
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