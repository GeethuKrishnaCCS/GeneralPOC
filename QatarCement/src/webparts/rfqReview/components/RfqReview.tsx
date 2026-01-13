import * as React from 'react';
import styles from './RfqReview.module.scss';
import type { IInitiatorResponse, IItemData, IManagerResponse, IProcurementManagerResponse, IRfqReviewProps, IRfqReviewState, IVendorResponseInput } from '../interfaces/IRfqReviewProps';
import { RfqReviewService } from '../services/RfqReviewService';
import ModalOverlay from '../../../shared/controls/Overlay/Overlay';
import { TextField } from '@fluentui/react';
import ToastService from '../../../shared/controls/Toast/Toast';
import * as strings from 'RfqReviewWebPartStrings';
import { HttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import * as moment from 'moment';
import RFQDeptDetailsTable from './RFQDeptDetailsTable';
import VendorDetailsTable from './VendorDetailsTable';
import InitiatorDetailsTable from './InitiatorDetailsTable';
import ManagerDetailsTable from './ManagerDetailsTable';
import ProcurementManagerDetailsTable from './ProcurementManagerDetailsTable';
import InitiatorDetailsTableForm from './InitiatorDetailsTableForm';


export default class RfqReview extends React.Component<IRfqReviewProps, IRfqReviewState, {}> {
  private service: RfqReviewService;
  constructor(props: IRfqReviewProps) {
    super(props);
    this.state = {
      modalOverlay: {
        isOpen: false,
        Text: ''
      },
      currentUser: { id: '', email: '', title: '' },
      userType: '', // Can now be: 'RFQDept', 'Vendor', or 'Initiator'
      prNumber: '',
      department: '',
      priority: '',
      dueDate: '',
      prInitiator: '',
      businessJustification: '',
      itemDetails: [],
      vendorOptions: [],
      masterid: '',
      taskID: null,
      vendorResponses: {} as Record<number, IVendorResponseInput>,
      initiatorResponses: {} as Record<number, IInitiatorResponse>,
      managerResponses: {} as Record<number, IManagerResponse>,
      procurementManagerResponses: {} as Record<number, IProcurementManagerResponse>,
      selectedFiles: [],
      uploadedFileUrls: [],
      attachments: [],
      isLoadingAttachments: false,
      commonManagerStatus: '',
      commonManagerComments: '',
      commonProcurementStatus: '',
      commonProcurementComments: '',
      workflowDetailsAttachments: {},
      vendorTermsAndConditions: '',
      vendorTechnicalSupport: '',
      vendorWarrantySupport: '',
    };

    this.service = new RfqReviewService(this.props.context, this.props.context.pageContext.web.absoluteUrl);
    this.validateURLParams = this.validateURLParams.bind(this);
    this.bindRFQData = this.bindRFQData.bind(this);
    this.bindWorkflowData = this.bindWorkflowData.bind(this);
    this.checkUserGroup = this.checkUserGroup.bind(this);
    this.bindMasterData = this.bindMasterData.bind(this);
    this.handleChange = this.handleChange.bind(this);
    this.triggerSubmit = this.triggerSubmit.bind(this);
    this.submitRFQDept = this.submitRFQDept.bind(this);
    this.handleCommonManagerStatusChange = this.handleCommonManagerStatusChange.bind(this);
    this.handleCommonManagerCommentsChange = this.handleCommonManagerCommentsChange.bind(this);
    this.handleCommonProcurementStatusChange = this.handleCommonProcurementStatusChange.bind(this);
    this.handleCommonProcurementCommentsChange = this.handleCommonProcurementCommentsChange.bind(this);
    this.handleVendorResponseChange = this.handleVendorResponseChange.bind(this);
    this.submitVendor = this.submitVendor.bind(this);
    this.triggerVendorSubmit = this.triggerVendorSubmit.bind(this);
    this.closeWindow = this.closeWindow.bind(this);
    this.handleInitiatorResponseChange = this.handleInitiatorResponseChange.bind(this);
    this.submitInitiator = this.submitInitiator.bind(this);
    this.triggerInitiatorSubmit = this.triggerInitiatorSubmit.bind(this);
    this.bindManagerData = this.bindManagerData.bind(this);
    this.bindProcurementManagerData = this.bindProcurementManagerData.bind(this);
    this.handleMaintenanceManagerResponseChange = this.handleMaintenanceManagerResponseChange.bind(this);
    this.handleProcurementManagerResponseChange = this.handleProcurementManagerResponseChange.bind(this);
    this.submitMaintenanceManager = this.submitMaintenanceManager.bind(this);
    this.submitProcurementManager = this.submitProcurementManager.bind(this);
    this.triggerMaintenanceManagerSubmit = this.triggerMaintenanceManagerSubmit.bind(this);
    this.triggerProcurementManagerSubmit = this.triggerProcurementManagerSubmit.bind(this);
    this.fetchAttachments = this.fetchAttachments.bind(this);
    this.bindMasterDataForInitiator = this.bindMasterDataForInitiator.bind(this);
    this.bindMasterDataForManager = this.bindMasterDataForManager.bind(this);
    this.bindMasterDataForProcurementManager = this.bindMasterDataForProcurementManager.bind(this);


  }
  public async componentDidMount(): Promise<void> {
    const user = await this.service.getCurrentUser();
    this.setState({
      modalOverlay: { isOpen: true, Text: 'Loading...' },
      currentUser: { id: user.Id, email: user.Email, title: user.Title }
    });
    /* Check queryparameter */
    await this.validateURLParams();

  }
  //validate url parameters
  public async validateURLParams() {
    const params = new URLSearchParams(window.location.search);
    const masterid = params.get('MID');
    const taskid = params.get('TID');
    const userParam = params.get('User'); // User type parameter

    if (masterid !== "" && masterid !== null && masterid !== undefined) {
      this.setState({ masterid: masterid });

      // 👉 CASE 1: MID + TID + User=Initiator
      if (userParam && userParam.toLowerCase() === 'initiator') {
        if (taskid !== "" && taskid !== null && taskid !== undefined) {
          this.setState({ taskID: taskid });
          await this.bindInitiatorData(masterid, taskid);
        } else {
          ToastService.error("Task ID is required for Initiator view.");
        }
      }
      // 👉 NEW CASE: MID + User=InitiatorForm (no TID required)
      else if (userParam && userParam.toLowerCase() === 'initiatorform') {
        this.setState({ userType: "InitiatorForm" });
        await this.bindMasterDataForInitiator(masterid);

      }
      // 👉 CASE 2: MID + TID + User=MaintenanceManager
      else if (userParam && userParam.toLowerCase() === 'maintenancemanager') {
        if (taskid !== "" && taskid !== null && taskid !== undefined) {
          this.setState({ taskID: taskid });
          await this.bindManagerData(masterid, taskid);
        } else {
          ToastService.error("Task ID is required for Maintenance Manager view.");
        }
      }
      // 👉 CASE 3: MID + TID + User=ProcurementManager
      else if (userParam && userParam.toLowerCase() === 'procurementmanager') {
        if (taskid !== "" && taskid !== null && taskid !== undefined) {
          this.setState({ taskID: taskid });
          await this.bindProcurementManagerData(masterid, taskid);
        } else {
          ToastService.error("Task ID is required for Procurement Manager view.");
        }
      }
      // 👉 CASE 4: MID + TID (Vendor)
      else if (taskid !== "" && taskid !== null && taskid !== undefined) {
        this.setState({ taskID: taskid });
        await this.bindWorkflowData(masterid, taskid);
      }
      // 👉 CASE 5: Only MID (RFQDept)
      else {
        await this.checkUserGroup(masterid);
      }
    } else {
      ToastService.error("Invalid URL parameters. Please check and try again.");
    }
  }

  // Check user in RFQDept group
  public async checkUserGroup(masterid: any) {
    const isInGroup = await this.service.isUserInGroup("RFQDept");
    if (!isInGroup) {
      ToastService.error("You do not have permission to access this form.");
      this.setState({ modalOverlay: { isOpen: true, Text: 'Access Denied' } });
    }
    else {
      /* Bind data */
      await this.bindRFQData();
      /*  Bind Master Data */
      await this.bindMasterData(masterid);

    }
  }
  // Bind Vendor Data
  public async bindRFQData() {
    // Bind Document Type
    const vendorqueryurl = this.props.context.pageContext.web.serverRelativeUrl + strings.queryList + this.props.wpproperties.vendorListName;
    const getVendorchoice = await this.service.getPagedListItems(vendorqueryurl);
    const VendorsName: { key: string, text: string }[] = [];
    getVendorchoice.map((item: any) => {
      VendorsName.push({ key: item.Title, text: item.Title });
    });
    this.setState({ vendorOptions: VendorsName, userType: "RFQDept" });
  }

  // Check current user from workflow task data
  public async bindWorkflowData(masterid: any, taskid: any) {
    try {
      const taskqueryurl = this.props.context.pageContext.web.serverRelativeUrl + strings.queryList + this.props.wpproperties.WorkflowTasksListName;
      const select = "*,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail";
      const expand = "AssignedTo";
      const workflowData = await this.service.getItemsByIdSelectExpand(taskqueryurl, Number(taskid), select, expand);
      console.log("Workflow Data:", workflowData);
      if (workflowData) {
        if (workflowData.AssignedTo.EMail.toLowerCase() === this.state.currentUser.email.toLowerCase()) {
          this.setState({ userType: "Vendor" })
          // Process workflow data as needed
          await this.bindMasterData(masterid);
        }
        else {
          this.setState({ modalOverlay: { isOpen: true, Text: 'Access Denied' } });
          ToastService.error("You are not authorized to access this task.");
          return;
        }

      } else {
        ToastService.error("No workflow data found for the provided Task ID and Master ID.");
      }
    } catch (error) {
      console.error("Error fetching workflow data:", error);
      ToastService.error("Failed to fetch workflow data. Please try again later.");
    }
  }

  public async bindMasterData(masterid: any) {
    let masterdata: any;
    let itemdetaildata: any[];
    let itemdetaildataitems: any[] = [];

    const masterqueryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.PRDetailsListName;
    const select = "*,PRInitiator/ID,PRInitiator/Title,PRInitiator/EMail";
    const expand = "PRInitiator";

    const itemdetailqueryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.PRItemSpecficationsListName;
    const itemfilter = "PRDetailsIDId eq '" + Number(masterid) + "'";

    try {
      masterdata = await this.service.getItemsByIdSelectExpand(masterqueryurl, Number(masterid), select, expand);
      console.log("masterdata", masterdata);

      itemdetaildata = await this.service.getPagedFilterListItems(itemdetailqueryurl, itemfilter);
      console.log("itemdetaildata", itemdetaildata);

      // ✅ Fetch WorkflowDetails only for Vendor
      let workflowDetailsMap: Record<string, number> = {};

      if (this.state.userType === "Vendor" && this.state.taskID) {
        const workflowDetailsQuery =
          this.props.context.pageContext.web.serverRelativeUrl +
          strings.queryList +
          "WorkflowDetails";

        const workflowFilter =
          `PRDetailID eq '${this.state.masterid}' and ` +
          `TaskID eq '${this.state.taskID}' and ` +
          `Vendor eq '${this.state.currentUser.email}'`;

        console.log("WorkflowDetails Filter:", workflowFilter);

        const workflowDetailsData =
          await this.service.getItemsFilter(workflowDetailsQuery, workflowFilter);

        console.log("workflowDetailsData fetched:", workflowDetailsData);

        // 🚨 CRITICAL VALIDATION: No workflow records found
        if (!workflowDetailsData || workflowDetailsData.length === 0) {
          this.setState({
            modalOverlay: { isOpen: true, Text: 'No Items Assigned' }
          });
          ToastService.error(
            "No items have been assigned to you for this RFQ. " +
            "Please contact the RFQ Department if you believe this is an error."
          );
          return; // ✅ Stop processing
        }

        // 🔑 Map PRItemID → WorkflowDetails ID
        workflowDetailsData.forEach((wfItem: any) => {
          if (wfItem.PRItemID) {
            const prItemIdKey = String(wfItem.PRItemID).trim();
            workflowDetailsMap[prItemIdKey] = wfItem.Id;
            console.log(`Mapped PRItemID ${prItemIdKey} -> WorkflowDetailsId ${wfItem.Id}`);
          }
        });

        console.log("WorkflowDetails Map:", workflowDetailsMap);
      }

      // ✅ Only process items that have WorkflowDetailsId (for vendors)
      if (itemdetaildata.length > 0) {
        itemdetaildata.forEach((item: any, index: any) => {
          const itemIdKey = String(item.Id).trim();
          const workflowDetailsId = workflowDetailsMap[itemIdKey] || null;

          // For vendors, only include items assigned to them
          if (this.state.userType === "Vendor" && !workflowDetailsId) {
            console.warn(`Skipping Item ${item.Id} - not assigned to vendor`);
            return; // Skip this item
          }

          console.log(`Item ${item.Id} (${item.ItemCode}) -> WorkflowDetailsId: ${workflowDetailsId}`);

          itemdetaildataitems.push({
            index: index + 1,
            Id: item.Id,
            Description: item.Description,
            ItemCode: item.ItemCode,
            Quantity: item.Qty,
            UOM: item.UoM,
            Title: item.Title,
            vendors: item.Vendors,
            WorkflowDetailsId: workflowDetailsId
          });
        });
      }

      // 🚨 FINAL VALIDATION: No items after filtering
      if (this.state.userType === "Vendor" && itemdetaildataitems.length === 0) {
        this.setState({
          modalOverlay: { isOpen: true, Text: 'No Items Available' }
        });
        ToastService.error("No items are currently available for you to respond to.");
        return;
      }

      console.log("Final itemDetails with WorkflowDetailsId:", itemdetaildataitems);

      this.setState({
        masterid: masterid,
        prNumber: masterdata.PRNumber,
        department: masterdata.Department,
        priority: masterdata.Priority,
        dueDate: masterdata.DueDate,
        prInitiator: masterdata.PRInitiator.Title,
        businessJustification: masterdata.BusinessJustification,
        modalOverlay: { isOpen: false, Text: '' },
        itemDetails: itemdetaildataitems
      });

      // Fetch attachments for vendor view
      if (this.state.userType === "Vendor") {
        await this.fetchAttachments(masterid);
      }

    } catch (error) {
      console.error("Error fetching data:", error);
      this.setState({ modalOverlay: { isOpen: true, Text: 'Error Loading Data' } });
      ToastService.error("Failed to fetch data. Please try again later.");
    }
  }

  //  fetch attachments:
  public async fetchAttachments(masterid: string) {
    this.setState({ isLoadingAttachments: true });

    try {
      const libraryName = "Shared Documents"; // or use this.props.wpproperties.DocumentLibraryName
      const attachments = await this.service.getAttachmentsByPRDetailID(libraryName, masterid);

      console.log("Fetched attachments for PRDetailID:", masterid, attachments);

      this.setState({
        attachments: attachments,
        isLoadingAttachments: false
      });
    } catch (error) {
      console.error("Error fetching attachments:", error);
      ToastService.error("Failed to load attachments.");
      this.setState({ isLoadingAttachments: false });
    }
  }

  // Bind Initiator Data
  public async bindInitiatorData(masterid: any, taskid: any) {
    try {
      const taskqueryurl = this.props.context.pageContext.web.serverRelativeUrl +
        strings.queryList + this.props.wpproperties.WorkflowTasksListName;
      const select = "*,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail";
      const expand = "AssignedTo";

      const workflowData = await this.service.getItemsByIdSelectExpand(
        taskqueryurl,
        Number(taskid),
        select,
        expand
      );

      console.log("Initiator Workflow Data:", workflowData);

      if (!workflowData) {
        ToastService.error("No workflow data found for the provided Task ID.");
        this.setState({ modalOverlay: { isOpen: true, Text: 'Task Not Found' } });
        return;
      }

      // Verify the current user is the PR Initiator
      if (workflowData.AssignedTo.EMail.toLowerCase() !== this.state.currentUser.email.toLowerCase()) {
        this.setState({ modalOverlay: { isOpen: true, Text: 'Access Denied' } });
        ToastService.error("You are not authorized to access this task as Initiator.");
        return;
      }

      this.setState({ userType: "Initiator" });
      await this.bindMasterDataForInitiator(masterid);

    } catch (error) {
      console.error("Error fetching initiator workflow data:", error);
      this.setState({ modalOverlay: { isOpen: true, Text: 'Error Loading Data' } });
      ToastService.error("Failed to fetch workflow data. Please try again later.");
    }
  }

  // ✅ IMPROVED: bindMasterDataForInitiator with validation
  public async bindMasterDataForInitiator(masterid: any) {
    let masterdata: any;
    let itemdetaildataitems: any[] = [];

    const masterqueryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.PRDetailsListName;
    const select = "*,PRInitiator/ID,PRInitiator/Title,PRInitiator/EMail";
    const expand = "PRInitiator";

    try {
      // Fetch master PR data
      masterdata = await this.service.getItemsByIdSelectExpand(masterqueryurl, Number(masterid), select, expand);
      console.log("masterdata", masterdata);

      // Fetch WorkflowDetails for all items related to this PR
      const workflowDetailsQuery = this.props.context.pageContext.web.serverRelativeUrl +
        strings.queryList + "WorkflowDetails";
      // const workflowFilter = `PRDetailID eq ${masterid}`;
      const workflowFilter = `PRDetailID eq ${masterid} and Status eq 'Available'`;

      const workflowDetailsData = await this.service.getItemsFilter(workflowDetailsQuery, workflowFilter);
      console.log("Initiator WorkflowDetails:", workflowDetailsData);

      // 🚨 CRITICAL VALIDATION: No workflow records (no vendor responses yet)
      if (!workflowDetailsData || workflowDetailsData.length === 0) {
        this.setState({
          modalOverlay: { isOpen: true, Text: 'No Vendor Responses Yet' }
        });
        ToastService.error(
          "No vendor responses have been submitted for this RFQ yet. " +
          "Please wait for vendors to respond before performing technical review."
        );
        return; // ✅ Stop processing
      }

      // Create separate row for each vendor response (no grouping)
      workflowDetailsData.forEach((wf: any, index: number) => {
        itemdetaildataitems.push({
          index: index + 1,
          Id: wf.PRItemID,
          WorkflowDetailsId: wf.Id,
          Description: wf.Description,
          ItemCode: wf.ItemCode,
          Quantity: wf.Qty,
          UOM: wf.UoM,
          Title: wf.Title,
          Vendor: wf.Vendor,
          Status: wf.Status,
          Price: wf.Price || '',
          Comments: wf.Comments || '',
          TaskID: wf.TaskID
        });
      });

      // 🚨 SECONDARY VALIDATION: Ensure items were processed
      if (itemdetaildataitems.length === 0) {
        this.setState({
          modalOverlay: { isOpen: true, Text: 'No Items Available' }
        });
        ToastService.error("No items are currently available for initiator review.");
        return;
      }

      console.log("Final itemDetails for Initiator:", itemdetaildataitems);

      this.setState({
        masterid: masterid,
        prNumber: masterdata.PRNumber,
        department: masterdata.Department,
        priority: masterdata.Priority,
        dueDate: masterdata.DueDate,
        prInitiator: masterdata.PRInitiator.Title,
        businessJustification: masterdata.BusinessJustification,
        vendorTermsAndConditions: masterdata.TermsAndConditions || '',
        vendorTechnicalSupport: masterdata.TechnicalSupport || '',
        vendorWarrantySupport: masterdata.WarrantySupport || '',
        modalOverlay: { isOpen: false, Text: '' },
        itemDetails: itemdetaildataitems
      });
      // Fetch attachments for initiator view (add after setState in bindMasterDataForInitiator)
      if (this.state.userType === "Initiator" || this.state.userType === "InitiatorForm") {
        await this.fetchWorkflowDetailsAttachments();
      }

    } catch (error) {
      console.error("Error fetching initiator data:", error);
      this.setState({ modalOverlay: { isOpen: true, Text: 'Error Loading Data' } });
      ToastService.error("Failed to fetch data. Please try again later.");
    }
  }

  // Add this new method to RfqReview class
  public async fetchWorkflowDetailsAttachments() {
    try {
      const attachmentsMap: Record<number, Array<{ name: string; url: string }>> = {};

      // Fetch attachments for each item
      for (const item of this.state.itemDetails) {
        if (item.WorkflowDetailsId) {
          const attachments = await this.service.getWorkflowDetailsAttachments(item.WorkflowDetailsId);
          if (attachments.length > 0) {
            attachmentsMap[item.WorkflowDetailsId] = attachments;
          }
        }
      }

      this.setState({ workflowDetailsAttachments: attachmentsMap });
      console.log("WorkflowDetails attachments fetched:", attachmentsMap);

    } catch (error) {
      console.error("Error fetching WorkflowDetails attachments:", error);
    }
  }
  public async bindManagerData(masterid: string, taskid: string) {
    try {
      // 1. Check group membership first (extra security layer)
      const isInMaintenanceManager = await this.service.isUserInGroup("MaintenanceManager");
      if (!isInMaintenanceManager) {
        ToastService.error("You must be a member of the Maintenance Manager group to access this view.");
        this.setState({ modalOverlay: { isOpen: true, Text: 'Access Denied' } });
        return;
      }

      // 2. Then check workflow task assignment
      const taskqueryurl = `${this.props.context.pageContext.web.serverRelativeUrl}${strings.queryList}${this.props.wpproperties.WorkflowTasksListName}`;
      const select = "*,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail";
      const expand = "AssignedTo";

      const workflowData = await this.service.getItemsByIdSelectExpand(
        taskqueryurl,
        Number(taskid),
        select,
        expand
      );

      if (!workflowData) {
        ToastService.error("No workflow data found for the provided Task ID.");
        this.setState({ modalOverlay: { isOpen: true, Text: 'Task Not Found' } });
        return;
      }

      if (workflowData.AssignedTo.EMail.toLowerCase() !== this.state.currentUser.email.toLowerCase()) {
        ToastService.error("You are not assigned to this task as Maintenance Manager.");
        this.setState({ modalOverlay: { isOpen: true, Text: 'Access Denied' } });
        return;
      }

      // If both checks pass
      this.setState({ userType: "MaintenanceManager" });
      await this.bindMasterDataForManager(masterid);

    } catch (error) {
      console.error("Error in bindManagerData:", error);
      this.setState({ modalOverlay: { isOpen: true, Text: 'Error Loading Data' } });
      ToastService.error("Failed to validate access. Please try again.");
    }
  }

  // Bind Master Data For Manager 
  public async bindMasterDataForManager(masterid: any) {
    let masterdata: any;
    let itemdetaildataitems: any[] = [];

    const masterqueryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.PRDetailsListName;
    const select = "*,PRInitiator/ID,PRInitiator/Title,PRInitiator/EMail";
    const expand = "PRInitiator";

    try {
      // Fetch master PR data
      masterdata = await this.service.getItemsByIdSelectExpand(masterqueryurl, Number(masterid), select, expand);
      console.log("masterdata", masterdata);

      // Fetch WorkflowDetails where InitiatorStatus = 'Technically Accepted'
      const workflowDetailsQuery = this.props.context.pageContext.web.serverRelativeUrl +
        strings.queryList + "WorkflowDetails";
      const workflowFilter = `PRDetailID eq ${masterid}`;
      // const workflowFilter = `PRDetailID eq ${masterid} and InitiatorStatus eq 'Technically Accepted'`;

      const workflowDetailsData = await this.service.getItemsFilter(workflowDetailsQuery, workflowFilter);
      console.log("Maintenance Manager WorkflowDetails:", workflowDetailsData);

      // 🚨 CRITICAL VALIDATION: No technically accepted items found
      if (!workflowDetailsData || workflowDetailsData.length === 0) {
        this.setState({
          modalOverlay: { isOpen: true, Text: 'No Items for Review' }
        });
        ToastService.error(
          "No items have been technically accepted by the Initiator yet. " +
          "Please wait for the Initiator to complete their technical review before proceeding."
        );
        return; // ✅ Stop processing
      }

      // Create separate row for each technically accepted item
      workflowDetailsData.forEach((wf: any, index: number) => {
        itemdetaildataitems.push({
          index: index + 1,
          Id: wf.PRItemID,
          WorkflowDetailsId: wf.Id,
          Description: wf.Description,
          ItemCode: wf.ItemCode,
          Quantity: String(wf.Qty),
          UOM: wf.UoM,
          Title: wf.Title,
          vendors: wf.Vendor || '',
          Vendor: wf.Vendor,
          Status: wf.Status,
          Price: wf.Price || '',
          Comments: wf.Comments || '',
          InitiatorStatus: wf.InitiatorStatus,
          InitiatorComments: wf.InitiatorComments || '',
          TaskID: wf.TaskID
        });
      });

      // 🚨 SECONDARY VALIDATION: Ensure items were processed
      if (itemdetaildataitems.length === 0) {
        this.setState({
          modalOverlay: { isOpen: true, Text: 'No Items Available' }
        });
        ToastService.error("No items are currently available for maintenance manager review.");
        return;
      }

      console.log("Final itemDetails for Manager:", itemdetaildataitems);

      this.setState({
        masterid: masterid,
        prNumber: masterdata.PRNumber,
        department: masterdata.Department,
        priority: masterdata.Priority,
        dueDate: masterdata.DueDate,
        prInitiator: masterdata.PRInitiator.Title,
        businessJustification: masterdata.BusinessJustification,
        modalOverlay: { isOpen: false, Text: '' },
        vendorTermsAndConditions: masterdata.TermsAndConditions || '',
        vendorTechnicalSupport: masterdata.TechnicalSupport || '',
        vendorWarrantySupport: masterdata.WarrantySupport || '',
        itemDetails: itemdetaildataitems
      });
      if (this.state.userType === "MaintenanceManager") {
        await this.fetchWorkflowDetailsAttachments();
      }

    } catch (error) {
      console.error("Error fetching manager data:", error);
      this.setState({ modalOverlay: { isOpen: true, Text: 'Error Loading Data' } });
      ToastService.error("Failed to fetch data. Please try again later.");
    }
  }


  // Bind Procurement Manager Data

  public async bindProcurementManagerData(masterid: any, taskid: any) {
    try {
      // 1. Check group membership first (optional extra security)
      const isInProcurementManager = await this.service.isUserInGroup("ProcurementManager");
      if (!isInProcurementManager) {
        ToastService.error("You must be a member of the Procurement Manager group to access this view.");
        this.setState({ modalOverlay: { isOpen: true, Text: 'Access Denied' } });
        return;
      }

      // 2. Check workflow task assignment
      const taskqueryurl = this.props.context.pageContext.web.serverRelativeUrl +
        strings.queryList + this.props.wpproperties.WorkflowTasksListName;
      const select = "*,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail";
      const expand = "AssignedTo";

      const workflowData = await this.service.getItemsByIdSelectExpand(
        taskqueryurl,
        Number(taskid),
        select,
        expand
      );

      console.log("Procurement Manager Workflow Data:", workflowData);

      if (!workflowData) {
        ToastService.error("No workflow data found for the provided Task ID.");
        this.setState({ modalOverlay: { isOpen: true, Text: 'Task Not Found' } });
        return;
      }

      // Verify the current user is the Procurement Manager
      if (workflowData.AssignedTo.EMail.toLowerCase() !== this.state.currentUser.email.toLowerCase()) {
        this.setState({ modalOverlay: { isOpen: true, Text: 'Access Denied' } });
        ToastService.error("You are not authorized to access this task as Procurement Manager.");
        return;
      }

      // If checks pass
      this.setState({ userType: "ProcurementManager" });
      await this.bindMasterDataForProcurementManager(masterid);

    } catch (error) {
      console.error("Error fetching procurement manager workflow data:", error);
      this.setState({ modalOverlay: { isOpen: true, Text: 'Error Loading Data' } });
      ToastService.error("Failed to fetch workflow data. Please try again later.");
    }
  }

  // ✅ IMPROVED: bindMasterDataForProcurementManager with validation
  public async bindMasterDataForProcurementManager(masterid: any) {
    let masterdata: any;
    let itemdetaildataitems: any[] = [];

    const masterqueryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.PRDetailsListName;
    const select = "*,PRInitiator/ID,PRInitiator/Title,PRInitiator/EMail";
    const expand = "PRInitiator";

    try {
      // Fetch master PR data
      masterdata = await this.service.getItemsByIdSelectExpand(masterqueryurl, Number(masterid), select, expand);
      console.log("masterdata", masterdata);

      // Fetch WorkflowDetails where ManagerStatus = 'Approve'
      const workflowDetailsQuery = this.props.context.pageContext.web.serverRelativeUrl +
        strings.queryList + "WorkflowDetails";
      const workflowFilter = `PRDetailID eq ${masterid} and MaintenanceManagerStatus eq 'Approve'`;

      const workflowDetailsData = await this.service.getItemsFilter(workflowDetailsQuery, workflowFilter);
      console.log("Procurement Manager WorkflowDetails:", workflowDetailsData);

      // 🚨 CRITICAL VALIDATION: No approve items found
      if (!workflowDetailsData || workflowDetailsData.length === 0) {
        this.setState({
          modalOverlay: { isOpen: true, Text: 'No Items for Review' }
        });
        ToastService.error(
          "No items have been approved by the Maintenance Manager yet. " +
          "Please wait for the Maintenance Manager to complete their review before proceeding."
        );
        return; // ✅ Stop processing
      }

      // Create separate row for each approved item
      workflowDetailsData.forEach((wf: any, index: number) => {
        itemdetaildataitems.push({
          index: index + 1,
          Id: wf.PRItemID,
          WorkflowDetailsId: wf.Id,
          Description: wf.Description,
          ItemCode: wf.ItemCode,
          Quantity: String(wf.Qty),
          UOM: wf.UoM,
          Title: wf.Title,
          vendors: wf.Vendor || '',
          Vendor: wf.Vendor,
          Status: wf.Status,
          Price: wf.Price || '',
          Comments: wf.Comments || '',
          InitiatorStatus: wf.InitiatorStatus,
          InitiatorComments: wf.InitiatorComments || '',
          ManagerStatus: wf.MaintenanceManagerStatus,
          ManagerComments: wf.MaintenanceManagerComments || '',
          TaskID: wf.TaskID
        });
      });

      // 🚨 SECONDARY VALIDATION: Ensure items were processed
      if (itemdetaildataitems.length === 0) {
        this.setState({
          modalOverlay: { isOpen: true, Text: 'No Items Available' }
        });
        ToastService.error("No items are currently available for procurement manager review.");
        return;
      }

      console.log("Final itemDetails for Procurement Manager:", itemdetaildataitems);

      this.setState({
        masterid: masterid,
        prNumber: masterdata.PRNumber,
        department: masterdata.Department,
        priority: masterdata.Priority,
        dueDate: masterdata.DueDate,
        prInitiator: masterdata.PRInitiator.Title,
        businessJustification: masterdata.BusinessJustification,
        modalOverlay: { isOpen: false, Text: '' },
        itemDetails: itemdetaildataitems
      });
      if (this.state.userType === "ProcurementManager") {
        await this.fetchWorkflowDetailsAttachments();
      }

    } catch (error) {
      console.error("Error fetching procurement manager data:", error);
      this.setState({ modalOverlay: { isOpen: true, Text: 'Error Loading Data' } });
      ToastService.error("Failed to fetch data. Please try again later.");
    }
  }


  // Handle file selection - append new files to existing ones
  public handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files = event.target.files;
    if (files) {
      const newFiles = Array.from(files);

      // Append new files to existing selected files
      this.setState(prevState => ({
        selectedFiles: [...prevState.selectedFiles, ...newFiles]
      }));

      // Reset the input value so the same file can be selected again if needed
      event.target.value = '';
    }
  };

  // Remove a selected file
  public removeFile = (index: number): void => {
    this.setState(prevState => ({
      selectedFiles: prevState.selectedFiles.filter((_, i) => i !== index)
    }));
  };


  // Upload files to SharePoint document library
  public uploadFilesToSharePoint = async (): Promise<string[]> => {
    const { selectedFiles, masterid } = this.state;
    const uploadedUrls: string[] = [];

    if (selectedFiles.length === 0) {
      return uploadedUrls;
    }

    try {
      // Construct the full server relative path for the document library
      // const libraryPath = `${this.props.context.pageContext.web.serverRelativeUrl}/${this.props.wpproperties.DocumentLibraryName}`;
      const libraryPath = `${this.props.context.pageContext.web.serverRelativeUrl}/Shared Documents`;

      // Prepare metadata to set on uploaded files
      const metadata = {
        PRDetailID: masterid  // Set the PRDetailID column value
      };

      // Upload each file
      for (const file of selectedFiles) {
        const fileName = `${masterid}_${file.name}`;
        const fileUrl = await this.service.uploadFile(
          libraryPath,
          fileName,
          file,
          metadata  // Pass metadata to set column values
        );
        uploadedUrls.push(fileUrl);
      }

      console.log("Files uploaded successfully:", uploadedUrls);
      // ToastService.success(`${selectedFiles.length} file(s) uploaded successfully`);

      return uploadedUrls;
    } catch (error) {
      console.error("Error uploading files:", error);
      ToastService.error("Failed to upload files. Please try again.");
      throw error;
    }
  };

  public handleChange = (index: number, field: keyof IItemData, value: string): void => {
    const vendorData = [...this.state.itemDetails];

    (vendorData[index][field] as any) = value;

    this.setState({ itemDetails: vendorData });
  };



  // Modified submitRFQDept to include file upload
  public submitRFQDept = async () => {
    this.setState({ modalOverlay: { isOpen: true, Text: 'Uploading files and submitting...' } });

    try {
      // Upload files first
      const uploadedUrls = await this.uploadFilesToSharePoint();
      // Store uploaded URLs in state
      this.setState({ uploadedFileUrls: uploadedUrls });
      // Then trigger the flow with file URLs included
      await this.triggerSubmit(uploadedUrls);

      this.setState({
        modalOverlay: { isOpen: false, Text: '' },
        selectedFiles: [] // Clear selected files after successful submission
      });
    } catch (error) {
      console.error("Submission error:", error);
      this.setState({ modalOverlay: { isOpen: false, Text: '' } });
    }
  };

  // Update triggerSubmit to accept file URLs
  public async triggerSubmit(fileUrls: string[] = []) {
    const queryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.FlowConnectionsListName;
    const flowName = "QatarCement_RFQSubmit";
    const filter = "Title eq '" + flowName + "'";
    const laUrl = await this.service.getItemsFilter(queryurl, filter);
    const postURL = laUrl[0].AppURL;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");

    const body: string = JSON.stringify({
      'MasterID': String(this.state.masterid),
      'ItemDetails': this.state.itemDetails,
      'AttachmentUrls': fileUrls // Include file URLs
    });

    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };

    const response = await this.props.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      postOptions
    );

    if (response) {
      const responseJSON = await response.json();
      if (response.ok) {
        console.log("Response from Flow:", responseJSON);
        ToastService.success("RFQ Submitted successfully.");
        this.setState({ modalOverlay: { isOpen: false, Text: '' } });
        // setTimeout(() => this.closeWindow(), 2000);
        setTimeout(() => {
          window.location.replace(this.props.context.pageContext.web.serverRelativeUrl);
        }, 5000);
      }
      else {
        ToastService.error("Failed to submit review. Please try again.");
      }
    }
  }

  public handleVendorResponseChange = (
    itemId: number,
    field: 'status' | 'price' | 'comments' | 'termsAndConditions' | 'technicalSupport' | 'warrantySupport',
    value: string
  ): void => {
    this.setState(prevState => ({
      vendorResponses: {
        ...prevState.vendorResponses,
        [itemId]: {
          ...(prevState.vendorResponses[itemId] || {}),
          [field]: value
        }
      }
    }));
  };
  // Handler for file changes per item
  public handleVendorFileChange = (itemId: number, files: File[]): void => {
    this.setState(prevState => ({
      vendorResponses: {
        ...prevState.vendorResponses,
        [itemId]: {
          ...(prevState.vendorResponses[itemId] || {}),
          attachments: files
        }
      }
    }));
  };

  // Handler for removing a file from an item
  public handleVendorFileRemove = (itemId: number, fileIndex: number): void => {
    this.setState(prevState => {
      const currentFiles = prevState.vendorResponses[itemId]?.attachments || [];
      const updatedFiles = currentFiles.filter((_, index) => index !== fileIndex);

      return {
        vendorResponses: {
          ...prevState.vendorResponses,
          [itemId]: {
            ...(prevState.vendorResponses[itemId] || {}),
            attachments: updatedFiles
          }
        }
      };
    });
  };

  public submitVendor = async () => {
    this.setState({ modalOverlay: { isOpen: true, Text: 'Submitting Vendor Response...' } });

    // call a NEW vendor-specific flow
    await this.triggerVendorSubmit();

    this.setState({ modalOverlay: { isOpen: false, Text: '' } });
  };

  public async triggerVendorSubmit() {
    const vendorFlow = "QatarCement_RFQVendorSubmit";

    const queryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.FlowConnectionsListName;
    const filter = `Title eq '${vendorFlow}'`;

    const laUrl = await this.service.getItemsFilter(queryurl, filter);
    const postURL = laUrl[0].AppURL;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");

    const userEmailLower = this.state.currentUser.email.toLowerCase().trim();

    const vendorAssignedItems = this.state.itemDetails.filter(item => {
      if (!item.vendors) return false;
      const vendorList = item.vendors.toLowerCase();
      const emails = vendorList.split(/[,;]/).map(e => e.trim());
      return emails.some(email => email === userEmailLower || email.includes(userEmailLower));
    });

    console.log("Vendor Assigned Items:", vendorAssignedItems);

    // 🔥 Upload files to WorkflowDetails ATTACHMENTS (not Shared Documents)
    const itemDetailsWithResponses = await Promise.all(
      vendorAssignedItems.map(async (item) => {
        let attachmentSuccess = false;
        let attachmentCount = 0;

        // ✅ Check if WorkflowDetailsId exists and files are selected
        const itemFiles = this.state.vendorResponses[item.Id]?.attachments;

        console.log(`Processing Item ${item.Id} (${item.ItemCode}):`, {
          WorkflowDetailsId: item.WorkflowDetailsId,
          HasFiles: itemFiles && itemFiles.length > 0,
          FileCount: itemFiles?.length || 0
        });

        if (!item.WorkflowDetailsId) {
          console.error(`❌ Item ${item.Id} (${item.ItemCode}) has NULL WorkflowDetailsId - Cannot attach files`);

          if (itemFiles && itemFiles.length > 0) {
            ToastService.error(`Cannot attach files for item ${item.ItemCode} - WorkflowDetailsId is missing`);
          }
        } else if (itemFiles && itemFiles.length > 0) {
          try {
            console.log(`✅ Attaching ${itemFiles.length} files to WorkflowDetails ID: ${item.WorkflowDetailsId}`);

            // 🔥 Use native SharePoint attachments
            await this.service.addAttachmentsToWorkflowDetails(
              item.WorkflowDetailsId,
              itemFiles
            );

            attachmentSuccess = true;
            attachmentCount = itemFiles.length;

            console.log(`✅ Successfully attached ${itemFiles.length} files to WorkflowDetails ID ${item.WorkflowDetailsId}`);
            // ToastService.success(`${itemFiles.length} file(s) attached for item ${item.ItemCode}`);

          } catch (error) {
            console.error(`❌ Error uploading files for item ${item.Id}:`, error);
            ToastService.error(`Failed to upload files for item ${item.ItemCode}: ${error.message}`);
          }
        }

        return {
          ...item,
          WorkflowDetailsId: item.WorkflowDetailsId,
          vendorStatus: this.state.vendorResponses[item.Id]?.status || '',
          vendorPrice: this.state.vendorResponses[item.Id]?.price || '',
          vendorComments: this.state.vendorResponses[item.Id]?.comments || '',
          attachmentSuccess: attachmentSuccess,
          attachmentCount: attachmentCount
        };
      })
    );

    console.log("Final Item Details with Responses:", itemDetailsWithResponses);

    const body: string = JSON.stringify({
      'TaskID': String(this.state.taskID),
      'VendorEmail': this.state.currentUser.email,
      'MasterID': String(this.state.masterid),
      'TermsAndConditions': this.state.vendorResponses[0]?.termsAndConditions || '',
      'TechnicalSupport': this.state.vendorResponses[0]?.technicalSupport || '',
      'WarrantySupport': this.state.vendorResponses[0]?.warrantySupport || '',
      'ItemDetails': itemDetailsWithResponses
    });

    console.log("Submitting Vendor Data to Flow:", JSON.parse(body));

    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };

    const response = await this.props.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      postOptions
    );

    if (response) {
      const responseJSON = await response.json();
      if (response.ok) {
        console.log("Response from Flow:", responseJSON);
        ToastService.success("Vendor response submitted successfully!");
        this.setState({
          vendorResponses: {},
          modalOverlay: { isOpen: false, Text: '' }
        });
        // setTimeout(() => this.closeWindow(), 2000);
        setTimeout(() => {
          window.location.replace(this.props.context.pageContext.web.serverRelativeUrl);
        }, 5000);
      }
      else {
        ToastService.error("Failed to submit vendor response. Please try again.");
      }

    }
  }


  public handleInitiatorResponseChange = (workflowDetailsId: number | null | undefined, field: 'status' | 'comments', value: string): void => {
    if (!workflowDetailsId) {
      console.warn("WorkflowDetailsId is null or undefined, cannot save response");
      return;
    }

    this.setState(prevState => ({
      initiatorResponses: {
        ...prevState.initiatorResponses,
        [workflowDetailsId]: {
          ...(prevState.initiatorResponses[workflowDetailsId] || {}),
          [field]: value
        }
      }
    }));
  };

  public submitInitiator = async () => {
    this.setState({ modalOverlay: { isOpen: true, Text: 'Submitting Initiator Response...' } });

    await this.triggerInitiatorSubmit();

    this.setState({ modalOverlay: { isOpen: false, Text: '' } });
  };


  public async triggerInitiatorSubmit() {
    const flowName = "QatarCement_RFQInitiatorSubmit";

    const queryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.FlowConnectionsListName;
    const filter = "Title eq '" + flowName + "'";

    const laUrl = await this.service.getItemsFilter(queryurl, filter);
    const postURL = laUrl[0].AppURL;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");

    // Prepare items with initiator technical review responses
    const itemDetailsWithInitiatorResponse = this.state.itemDetails.map((item) => {
      const workflowDetailsId = item.WorkflowDetailsId ?? 0; // Use 0 as fallback if null

      return {
        WorkflowDetailsId: item.WorkflowDetailsId,
        PRItemID: item.Id,
        ItemCode: item.ItemCode,
        Description: item.Description,
        Quantity: item.Quantity,
        UOM: item.UOM,
        Vendor: item.Vendor || '',
        VendorPrice: item.Price || '',
        VendorComments: item.Comments || '',
        TaskID: item.TaskID || '',
        technicalStatus: this.state.initiatorResponses[workflowDetailsId]?.status || '',
        technicalComments: this.state.initiatorResponses[workflowDetailsId]?.comments || '',
      };
    });

    const body: string = JSON.stringify({
      'TaskID': String(this.state.taskID),
      'InitiatorEmail': this.state.currentUser.email,
      'MasterID': String(this.state.masterid),
      'ItemDetails': itemDetailsWithInitiatorResponse,
    });

    console.log("Submitting Initiator Data:", itemDetailsWithInitiatorResponse);

    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };

    const response = await this.props.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      postOptions
    );

    if (response) {
      const responseJSON = await response.json();
      if (response.ok) {
        console.log("Response from Flow:", responseJSON);
        ToastService.success("Initiator Review Submitted successfully.");
        this.setState({ modalOverlay: { isOpen: false, Text: '' } });
        // setTimeout(() => this.closeWindow(), 2000);
        setTimeout(() => {
          window.location.replace(this.props.context.pageContext.web.serverRelativeUrl);
        }, 5000);
      }
      else {
        ToastService.error("Failed to submit initiator review. Please try again.");
      }
    }
  }
  public handleMaintenanceManagerResponseChange = (workflowDetailsId: number | null | undefined, field: 'status' | 'comments', value: string): void => {
    if (!workflowDetailsId) {
      console.warn("WorkflowDetailsId is null or undefined, cannot save response");
      return;
    }

    this.setState(prevState => ({
      managerResponses: {
        ...prevState.managerResponses,
        [workflowDetailsId]: {
          ...(prevState.managerResponses[workflowDetailsId] || {}),
          [field]: value
        }
      }
    }));
  };

  // Handler for common manager status
  public handleCommonManagerStatusChange = (value: string): void => {
    this.setState({ commonManagerStatus: value });
  };

  // Handler for common manager comments
  public handleCommonManagerCommentsChange = (value: string): void => {
    this.setState({ commonManagerComments: value });
  };

  // Procurement Manager Response Handler
  public handleProcurementManagerResponseChange = (workflowDetailsId: number | null | undefined, field: 'status' | 'comments', value: string): void => {
    if (!workflowDetailsId) {
      console.warn("WorkflowDetailsId is null or undefined, cannot save response");
      return;
    }

    this.setState(prevState => ({
      procurementManagerResponses: {
        ...prevState.procurementManagerResponses,
        [workflowDetailsId]: {
          ...(prevState.procurementManagerResponses[workflowDetailsId] || {}),
          [field]: value
        }
      }
    }));
  };

  // Submit Manager
  public submitMaintenanceManager = async () => {
    this.setState({ modalOverlay: { isOpen: true, Text: 'Submitting Maintenance Manager Response...' } });
    await this.triggerMaintenanceManagerSubmit();
    this.setState({ modalOverlay: { isOpen: false, Text: '' } });
  };


  public async triggerMaintenanceManagerSubmit() {
    const managerFlow = "QatarCement_RFQMaintenanceManagerSubmit";

    const queryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.FlowConnectionsListName;
    const filter = `Title eq '${managerFlow}'`;

    const laUrl = await this.service.getItemsFilter(queryurl, filter);
    const postURL = laUrl[0].AppURL;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");

    const itemDetailsWithManagerResponse = this.state.itemDetails.map((item) => {
      const workflowDetailsId = item.WorkflowDetailsId ?? 0;

      return {
        WorkflowDetailsId: item.WorkflowDetailsId,
        PRItemID: item.Id,
        ItemCode: item.ItemCode,
        Description: item.Description,
        Quantity: item.Quantity,
        UOM: item.UOM,
        Vendor: item.Vendor || '',
        VendorPrice: item.Price || '',
        VendorComments: item.Comments || '',
        InitiatorStatus: item.InitiatorStatus || '',
        InitiatorComments: item.InitiatorComments || '',
        TaskID: item.TaskID || '',
        MaintenanceManagerStatus: this.state.commonManagerStatus, // ✅ Use common status
        MaintenanceManagerComments: this.state.managerResponses[workflowDetailsId]?.comments || '', // ✅ Per-item comments
      };
    });

    const body: string = JSON.stringify({
      'TaskID': String(this.state.taskID),
      'ManagerEmail': this.state.currentUser.email,
      'MasterID': String(this.state.masterid),
      'CommonManagerStatus': this.state.commonManagerStatus, // ✅ Add common status
      'CommonManagerComments': this.state.commonManagerComments, // ✅ Add common comments
      'ItemDetails': itemDetailsWithManagerResponse,
    });

    console.log("Submitting Maintenance Manager Data:", itemDetailsWithManagerResponse);

    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };

    const response = await this.props.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      postOptions
    );

    if (response) {
      console.log("Maintenance Manager Response:", response);

      const responseJSON = await response.json();
      if (response.ok) {
        console.log("Response from Flow:", responseJSON);
        ToastService.success("Maintenance Manager Submitted successfully.");
        this.setState({ modalOverlay: { isOpen: false, Text: '' } });
        // setTimeout(() => this.closeWindow(), 2000);
        setTimeout(() => {
          window.location.replace(this.props.context.pageContext.web.serverRelativeUrl);
        }, 5000);
      } else {
        ToastService.error("Failed to submit maintenance manager review. Please try again.");
      }
    }
  }
  // Handler for common procurement status
  public handleCommonProcurementStatusChange = (value: string): void => {
    this.setState({ commonProcurementStatus: value });
  };

  // Handler for common procurement comments
  public handleCommonProcurementCommentsChange = (value: string): void => {
    this.setState({ commonProcurementComments: value });
  };
  // Submit Procurement Manager
  public submitProcurementManager = async () => {
    this.setState({ modalOverlay: { isOpen: true, Text: 'Submitting Procurement Manager Response...' } });
    await this.triggerProcurementManagerSubmit();
    this.setState({ modalOverlay: { isOpen: false, Text: '' } });
  };

  // Trigger Procurement Manager Submit
  // public async triggerProcurementManagerSubmit() {
  //   const procurementManagerFlow = "QatarCement_RFQProcurementManagerSubmit";

  //   const queryurl = this.props.context.pageContext.web.serverRelativeUrl +
  //     strings.queryList + this.props.wpproperties.FlowConnectionsListName;
  //   const filter = `Title eq '${procurementManagerFlow}'`;

  //   const laUrl = await this.service.getItemsFilter(queryurl, filter);
  //   const postURL = laUrl[0].AppURL;

  //   const requestHeaders: Headers = new Headers();
  //   requestHeaders.append("Content-type", "application/json");

  //   const itemDetailsWithProcurementManagerResponse = this.state.itemDetails.map((item) => {
  //     const workflowDetailsId = item.WorkflowDetailsId ?? 0;

  //     return {
  //       WorkflowDetailsId: item.WorkflowDetailsId,
  //       PRItemID: item.Id,
  //       ItemCode: item.ItemCode,
  //       Description: item.Description,
  //       Quantity: item.Quantity,
  //       UOM: item.UOM,
  //       Vendor: item.Vendor || '',
  //       VendorPrice: item.Price || '',
  //       VendorComments: item.Comments || '',
  //       InitiatorStatus: item.InitiatorStatus || '',
  //       InitiatorComments: item.InitiatorComments || '',
  //       ManagerStatus: item.ManagerStatus || '',
  //       ManagerComments: item.ManagerComments || '',
  //       TaskID: item.TaskID || '',
  //       procurementManagerStatus: this.state.procurementManagerResponses[workflowDetailsId]?.status || '',
  //       procurementManagerComments: this.state.procurementManagerResponses[workflowDetailsId]?.comments || '',
  //     };
  //   });

  //   const body: string = JSON.stringify({
  //     'TaskID': String(this.state.taskID),
  //     'ProcurementManagerEmail': this.state.currentUser.email,
  //     'MasterID': String(this.state.masterid),
  //     'ItemDetails': itemDetailsWithProcurementManagerResponse,
  //   });

  //   console.log("Submitting Procurement Manager Data:", itemDetailsWithProcurementManagerResponse);

  //   const postOptions: IHttpClientOptions = {
  //     headers: requestHeaders,
  //     body: body
  //   };

  //   const response = await this.props.context.httpClient.post(
  //     postURL,
  //     HttpClient.configurations.v1,
  //     postOptions
  //   );

  //   if (response) {
  //     console.log("Procurement Manager Response:", response);

  //     const responseJSON = await response.json();
  //     if (response.ok) {
  //       console.log("Response from Flow:", responseJSON);
  //       ToastService.success("Procurement Manager Submitted successfully.");
  //       this.setState({ modalOverlay: { isOpen: false, Text: '' } });
  //       setTimeout(() => this.closeWindow(), 2000);
  //     } else {
  //       ToastService.error("Failed to submit procurement manager review. Please try again.");
  //     }
  //   }
  // }

  public async triggerProcurementManagerSubmit() {
    const procurementManagerFlow = "QatarCement_RFQProcurementManagerSubmit";

    const queryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.FlowConnectionsListName;
    const filter = `Title eq '${procurementManagerFlow}'`;

    const laUrl = await this.service.getItemsFilter(queryurl, filter);
    const postURL = laUrl[0].AppURL;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");

    const itemDetailsWithProcurementManagerResponse = this.state.itemDetails.map((item) => {
      const workflowDetailsId = item.WorkflowDetailsId ?? 0;

      return {
        WorkflowDetailsId: item.WorkflowDetailsId,
        PRItemID: item.Id,
        ItemCode: item.ItemCode,
        Description: item.Description,
        Quantity: item.Quantity,
        UOM: item.UOM,
        Vendor: item.Vendor || '',
        VendorPrice: item.Price || '',
        VendorComments: item.Comments || '',
        InitiatorStatus: item.InitiatorStatus || '',
        InitiatorComments: item.InitiatorComments || '',
        ManagerStatus: item.ManagerStatus || '',
        ManagerComments: item.ManagerComments || '',
        TaskID: item.TaskID || '',
        procurementManagerStatus: this.state.commonProcurementStatus, // ✅ Use common status
        procurementManagerComments: this.state.procurementManagerResponses[workflowDetailsId]?.comments || '', // ✅ Per-item comments
      };
    });

    const body: string = JSON.stringify({
      'TaskID': String(this.state.taskID),
      'ProcurementManagerEmail': this.state.currentUser.email,
      'MasterID': String(this.state.masterid),
      'CommonProcurementStatus': this.state.commonProcurementStatus, // ✅ Add common status
      'CommonProcurementComments': this.state.commonProcurementComments, // ✅ Add common comments
      'ItemDetails': itemDetailsWithProcurementManagerResponse,
    });

    console.log("Submitting Procurement Manager Data:", itemDetailsWithProcurementManagerResponse);

    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };

    const response = await this.props.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      postOptions
    );

    if (response) {
      console.log("Procurement Manager Response:", response);

      const responseJSON = await response.json();
      if (response.ok) {
        console.log("Response from Flow:", responseJSON);
        ToastService.success("Procurement Manager Submitted successfully.");
        this.setState({ modalOverlay: { isOpen: false, Text: '' } });
        setTimeout(() => this.closeWindow(), 2000);
      } else {
        ToastService.error("Failed to submit procurement manager review. Please try again.");
      }
    }
  }

  public closeWindow = () => {
    ToastService.success("Redirecting to homepage...");
    setTimeout(() => {
      window.location.replace(this.props.context.pageContext.web.serverRelativeUrl);
    }, 2000);
  };

  public render(): React.ReactElement<IRfqReviewProps> {

    // Common styles for TextField
    const textFieldStyles = {
      field: {
        backgroundColor: '#f0f0f0', // Light grey for readonly, white otherwise
      },
    };
    return (
      <section className={styles.container}>
        <div className={styles.formpopup}>
          <div className={styles.formheader}>
            <div className={styles.formtitle}>{this.props.wpproperties.webpartTitle}</div>
          </div>
          <div className={styles.formbody}>
            <div className={styles.row}>
              <div className={styles.col6}>
                <TextField label="PR Number" value={this.state.prNumber} readOnly styles={textFieldStyles} />
              </div>
              <div className={styles.col6}>
                <TextField label="Department" value={this.state.department} readOnly styles={textFieldStyles} />
              </div>
            </div>
            <div className={styles.row}>
              <div className={styles.col4}>
                <TextField label="Priority" value={this.state.priority} readOnly styles={textFieldStyles} />
              </div>
              <div className={styles.col4}>
                <TextField label="Due Date" value={moment(this.state.dueDate).format(strings.DateFormat)} readOnly styles={textFieldStyles} />
              </div>
              <div className={styles.col4}>
                <TextField label="PR Initiator" value={this.state.prInitiator} readOnly styles={textFieldStyles} />
              </div>
            </div>
            <div className={styles.row}>
              <div className={styles.col12}>
                <TextField label="Business Justification" multiline rows={3} autoAdjustHeight value={this.state.businessJustification} readOnly styles={textFieldStyles} />
              </div>
            </div>
            <div className={styles.row}>
              <div className={styles.col12}>
                <h3 >Item Details</h3>
              </div>
            </div>

            {/* RFQDept View */}
            {this.state.userType === "RFQDept" && <>
              <RFQDeptDetailsTable
                itemDetails={this.state.itemDetails}
                vendorOptions={this.state.vendorOptions}
                handleChange={this.handleChange}
                onSubmitRFQDept={this.submitRFQDept}
                onCancel={this.closeWindow}
                selectedFiles={this.state.selectedFiles}
                onFileSelect={this.handleFileSelect}
                onRemoveFile={this.removeFile}
              />
            </>}
            {/* Vendor View */}
            {this.state.userType === "Vendor" && (<>
              <VendorDetailsTable
                itemDetails={this.state.itemDetails}
                currentUserEmail={this.state.currentUser.email}
                onSubmitVendor={this.submitVendor}
                onCancel={this.closeWindow}
                onResponseChange={this.handleVendorResponseChange}
                onFileChange={this.handleVendorFileChange}
                onRemoveFile={this.handleVendorFileRemove}
                vendorResponses={this.state.vendorResponses}
                attachments={this.state.attachments}
                isLoadingAttachments={this.state.isLoadingAttachments}
              />
            </>)}
            {/* Initiator View - NEW */}
            {this.state.userType === "Initiator" && (
              <InitiatorDetailsTable
                itemDetails={this.state.itemDetails}
                masterid={this.state.masterid}
                taskID={this.state.taskID}
                initiatorResponses={this.state.initiatorResponses}
                workflowDetailsAttachments={this.state.workflowDetailsAttachments || {}}  // ✅ Add this
                vendorTermsAndConditions={this.state.vendorTermsAndConditions}
                vendorTechnicalSupport={this.state.vendorTechnicalSupport}
                vendorWarrantySupport={this.state.vendorWarrantySupport}
                onResponseChange={this.handleInitiatorResponseChange}
                onSubmitInitiator={this.submitInitiator}
                onCancel={this.closeWindow}
              />
            )}

            {/* Initiator Form*/}
            {this.state.userType === "InitiatorForm" && (
              <InitiatorDetailsTableForm
                itemDetails={this.state.itemDetails}
                masterid={this.state.masterid}
                taskID={this.state.taskID}
                vendorTermsAndConditions={this.state.vendorTermsAndConditions}
                vendorTechnicalSupport={this.state.vendorTechnicalSupport}
                vendorWarrantySupport={this.state.vendorWarrantySupport}
                workflowDetailsAttachments={this.state.workflowDetailsAttachments || {}}  // ✅ Add this

              />
            )}

            {/* Manager View */}
            {this.state.userType === "MaintenanceManager" && (
              <ManagerDetailsTable
                itemDetails={this.state.itemDetails}
                masterid={this.state.masterid}
                taskID={this.state.taskID}
                managerResponses={this.state.managerResponses}
                commonManagerStatus={this.state.commonManagerStatus}
                commonManagerComments={this.state.commonManagerComments}
                onResponseChange={this.handleMaintenanceManagerResponseChange}
                onCommonStatusChange={this.handleCommonManagerStatusChange}
                onCommonCommentsChange={this.handleCommonManagerCommentsChange}
                onSubmitManager={this.submitMaintenanceManager}
                onCancel={this.closeWindow}
                workflowDetailsAttachments={this.state.workflowDetailsAttachments || {}}
                vendorTermsAndConditions={this.state.vendorTermsAndConditions}
                vendorTechnicalSupport={this.state.vendorTechnicalSupport}
                vendorWarrantySupport={this.state.vendorWarrantySupport}
              />
            )}

            {/* Procurement Manager View */}
            {this.state.userType === "ProcurementManager" && (
              <ProcurementManagerDetailsTable
                itemDetails={this.state.itemDetails}
                masterid={this.state.masterid}
                taskID={this.state.taskID}
                procurementManagerResponses={this.state.procurementManagerResponses}
                commonProcurementStatus={this.state.commonProcurementStatus}
                commonProcurementComments={this.state.commonProcurementComments}
                onResponseChange={this.handleProcurementManagerResponseChange}
                onCommonStatusChange={this.handleCommonProcurementStatusChange}
                onCommonCommentsChange={this.handleCommonProcurementCommentsChange}
                onSubmitProcurementManager={this.submitProcurementManager}
                onCancel={this.closeWindow}
                workflowDetailsAttachments={this.state.workflowDetailsAttachments || {}}  // ✅ Add this
                vendorTermsAndConditions={this.state.vendorTermsAndConditions}
                vendorTechnicalSupport={this.state.vendorTechnicalSupport}
                vendorWarrantySupport={this.state.vendorWarrantySupport}
              />
            )}
          </div>
        </div>
        {ToastService.container()}
        <ModalOverlay
          isModalOpen={this.state.modalOverlay.isOpen}
          modalText={this.state.modalOverlay.Text}
        />
      </section>
    );
  }
}
