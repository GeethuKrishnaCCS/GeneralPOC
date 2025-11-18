import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRfqReviewProps {
  wpproperties: any;
  context: WebPartContext;
}
export interface IRfqReviewWebPartProps {
  webpartTitle: string;
  PRDetailsListName: string;
  PRItemSpecficationsListName: string;
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
}