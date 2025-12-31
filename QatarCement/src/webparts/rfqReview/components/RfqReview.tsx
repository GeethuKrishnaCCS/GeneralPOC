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
      // vendorResponses: {} as Record<number, { status?: string; price?: string; comments?: string }>,
      // initiatorResponses: {} as Record<number, { status?: string; comments?: string }> // ⭐ ADD THIS
      vendorResponses: {} as Record<number, IVendorResponseInput>,
      initiatorResponses: {} as Record<number, IInitiatorResponse>,
      managerResponses: {} as Record<number, IManagerResponse>,
      procurementManagerResponses: {} as Record<number, IProcurementManagerResponse>
    };

    this.service = new RfqReviewService(this.props.context, this.props.context.pageContext.web.absoluteUrl);
    this.validateURLParams = this.validateURLParams.bind(this);
    this.bindVendorData = this.bindVendorData.bind(this);
    this.bindWorkflowData = this.bindWorkflowData.bind(this);
    this.checkUserGroup = this.checkUserGroup.bind(this);
    this.bindMasterData = this.bindMasterData.bind(this);
    this.handleChange = this.handleChange.bind(this);
    this.triggerSubmit = this.triggerSubmit.bind(this);
    this.submitRFQDept = this.submitRFQDept.bind(this);
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


  // Add this new method to handle Initiator case
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

      if (workflowData) {
        // Verify the current user is the PR Initiator
        if (workflowData.AssignedTo.EMail.toLowerCase() === this.state.currentUser.email.toLowerCase()) {
          this.setState({ userType: "Initiator" });
          await this.bindMasterDataForInitiator(masterid);
        } else {
          this.setState({ modalOverlay: { isOpen: true, Text: 'Access Denied' } });
          ToastService.error("You are not authorized to access this task as Initiator.");
          return;
        }
      } else {
        ToastService.error("No workflow data found for the provided Task ID.");
      }
    } catch (error) {
      console.error("Error fetching initiator workflow data:", error);
      ToastService.error("Failed to fetch workflow data. Please try again later.");
    }
  }

  // Bind Maintenance Manager Data
  // public async bindManagerData(masterid: any, taskid: any) {
  //   try {
  //     const taskqueryurl = this.props.context.pageContext.web.serverRelativeUrl +
  //       strings.queryList + this.props.wpproperties.WorkflowTasksListName;
  //     const select = "*,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail";
  //     const expand = "AssignedTo";

  //     const workflowData = await this.service.getItemsByIdSelectExpand(
  //       taskqueryurl,
  //       Number(taskid),
  //       select,
  //       expand
  //     );

  //     console.log("Maintenance Manager Workflow Data:", workflowData);

  //     if (workflowData) {
  //       // Verify the current user is the Maintenance Manager
  //       if (workflowData.AssignedTo.EMail.toLowerCase() === this.state.currentUser.email.toLowerCase()) {
  //         this.setState({ userType: "MaintenanceManager" });
  //         await this.bindMasterDataForManager(masterid);
  //       } else {
  //         this.setState({ modalOverlay: { isOpen: true, Text: 'Access Denied' } });
  //         ToastService.error("You are not authorized to access this task as Maintenance Manager.");
  //         return;
  //       }
  //     } else {
  //       ToastService.error("No workflow data found for the provided Task ID.");
  //     }
  //   } catch (error) {
  //     console.error("Error fetching manager workflow data:", error);
  //     ToastService.error("Failed to fetch workflow data. Please try again later.");
  //   }
  // }

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
      const workflowFilter = `PRDetailID eq ${masterid} and InitiatorStatus eq 'Technically Accepted'`;

      const workflowDetailsData = await this.service.getItemsFilter(workflowDetailsQuery, workflowFilter);
      console.log("Maintenance Manager WorkflowDetails:", workflowDetailsData);

      if (workflowDetailsData.length > 0) {
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
      }

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

    } catch (error) {
      console.error("Error fetching manager data:", error);
      ToastService.error("Failed to fetch data. Please try again later.");
    }
  }

  // Bind Procurement Manager Data
  public async bindProcurementManagerData(masterid: any, taskid: any) {
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

      console.log("Procurement Manager Workflow Data:", workflowData);

      if (workflowData) {
        // Verify the current user is the Procurement Manager
        if (workflowData.AssignedTo.EMail.toLowerCase() === this.state.currentUser.email.toLowerCase()) {
          this.setState({ userType: "ProcurementManager" });
          await this.bindMasterDataForProcurementManager(masterid);
        } else {
          this.setState({ modalOverlay: { isOpen: true, Text: 'Access Denied' } });
          ToastService.error("You are not authorized to access this task as Procurement Manager.");
          return;
        }
      } else {
        ToastService.error("No workflow data found for the provided Task ID.");
      }
    } catch (error) {
      console.error("Error fetching procurement manager workflow data:", error);
      ToastService.error("Failed to fetch workflow data. Please try again later.");
    }
  }

  // Bind Master Data For Procurement Manager
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

      // Fetch WorkflowDetails where ManagerStatus = 'Approved'
      const workflowDetailsQuery = this.props.context.pageContext.web.serverRelativeUrl +
        strings.queryList + "WorkflowDetails";
      const workflowFilter = `PRDetailID eq ${masterid} and ManagerStatus eq 'Approved'`;

      const workflowDetailsData = await this.service.getItemsFilter(workflowDetailsQuery, workflowFilter);
      console.log("Procurement Manager WorkflowDetails:", workflowDetailsData);

      if (workflowDetailsData.length > 0) {
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
            ManagerStatus: wf.ManagerStatus,
            ManagerComments: wf.ManagerComments || '',
            TaskID: wf.TaskID
          });
        });
      }

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

    } catch (error) {
      console.error("Error fetching procurement manager data:", error);
      ToastService.error("Failed to fetch data. Please try again later.");
    }
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
  // Check user in RFQDept group
  public async checkUserGroup(masterid: any) {
    const isInGroup = await this.service.isUserInGroup("RFQDept");
    if (!isInGroup) {
      ToastService.error("You do not have permission to access this form.");
      this.setState({ modalOverlay: { isOpen: true, Text: 'Access Denied' } });
    }
    else {
      /* Bind data */
      await this.bindVendorData();
      /*  Bind Master Data */
      await this.bindMasterData(masterid);

    }
  }
  // Bind Vendor Data
  public async bindVendorData() {
    // Bind Document Type
    const vendorqueryurl = this.props.context.pageContext.web.serverRelativeUrl + strings.queryList + this.props.wpproperties.vendorListName;
    const getVendorchoice = await this.service.getPagedListItems(vendorqueryurl);
    const VendorsName: { key: string, text: string }[] = [];
    getVendorchoice.map((item: any) => {
      VendorsName.push({ key: item.Title, text: item.Title });
    });
    this.setState({ vendorOptions: VendorsName, userType: "RFQDept" });
  }



  // bindMasterData method in RfqReview.tsx with this updated version:
  public async bindMasterData(masterid: any) {
    let masterdata: any;
    let itemdetaildata: any[];
    let itemdetaildataitems: any[] = [];

    // Fetch master index item
    const masterqueryurl = this.props.context.pageContext.web.serverRelativeUrl + strings.queryList + this.props.wpproperties.PRDetailsListName;
    const select = "*,PRInitiator/ID,PRInitiator/Title,PRInitiator/EMail";
    const expand = "PRInitiator";

    // Fetch master item details
    const itemdetailqueryurl = this.props.context.pageContext.web.serverRelativeUrl + strings.queryList + this.props.wpproperties.PRItemSpecficationsListName;
    const itemfilter = "PRDetailsIDId eq '" + Number(masterid) + "'";

    try {
      masterdata = await this.service.getItemsByIdSelectExpand(masterqueryurl, Number(masterid), select, expand);
      console.log("masterdata", masterdata);

      itemdetaildata = await this.service.getPagedFilterListItems(itemdetailqueryurl, itemfilter);
      console.log("itemdetaildata", itemdetaildata);

      // ✓ Fetch WorkflowDetails only for Vendor
      let workflowDetailsMap: Record<string, number> = {};

      if (this.state.userType === "Vendor" && this.state.taskID) {

        const workflowDetailsQuery =
          this.props.context.pageContext.web.serverRelativeUrl +
          strings.queryList +
          "WorkflowDetails";

        // Remove quotes from numeric fields
        const workflowFilter =
          `PRDetailID eq ${this.state.masterid} and ` +
          `TaskID eq ${this.state.taskID} and ` +
          `Vendor eq '${this.state.currentUser.email}'`;

        const workflowDetailsData =
          await this.service.getItemsFilter(workflowDetailsQuery, workflowFilter);

        console.log("workflowDetailsData", workflowDetailsData);

        // 🔑 Map PRItemID (text) → WorkflowDetails ID
        workflowDetailsData.forEach((wfItem: any) => {
          if (wfItem.PRItemID) {
            workflowDetailsMap[String(wfItem.PRItemID)] = wfItem.Id;
          }
        });
      }

      if (itemdetaildata.length > 0) {
        itemdetaildata.forEach((item: any, index: any) => {
          itemdetaildataitems.push({
            index: index + 1,
            Id: item.Id,
            Description: item.Description,
            ItemCode: item.ItemCode,
            Quantity: item.Qty,
            UOM: item.UoM,
            Title: item.Title,
            vendors: item.Vendors,

            // ✅ THIS IS THE KEY LINE
            WorkflowDetailsId: workflowDetailsMap[String(item.Id)] || null
          });
        });
      }

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

    } catch (error) {
      console.error("Error fetching data:", error);
      ToastService.error("Failed to fetch data. Please try again later.");
    }
  }

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
      const workflowFilter = `PRDetailID eq ${masterid}`;

      const workflowDetailsData = await this.service.getItemsFilter(workflowDetailsQuery, workflowFilter);
      console.log("Initiator WorkflowDetails:", workflowDetailsData);

      if (workflowDetailsData.length > 0) {
        // Create separate row for each vendor response (no grouping)
        workflowDetailsData.forEach((wf: any, index: number) => {
          itemdetaildataitems.push({
            index: index + 1,
            Id: wf.PRItemID,                    // Item ID from WorkflowDetails
            WorkflowDetailsId: wf.Id,           // WorkflowDetails record ID
            Description: wf.Description,
            ItemCode: wf.ItemCode,
            Quantity: wf.Qty,
            UOM: wf.UoM,
            Title: wf.Title,
            Vendor: wf.Vendor,                  // Individual vendor email
            Status: wf.Status,                  // Individual vendor email
            Price: wf.Price || '',              // Vendor's quoted price
            Comments: wf.Comments || '',        // Vendor's comments
            TaskID: wf.TaskID                   // Associated task ID
          });
        });
      }

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

    } catch (error) {
      console.error("Error fetching data:", error);
      ToastService.error("Failed to fetch data. Please try again later.");
    }
  }

  public handleChange = (index: number, field: keyof IItemData, value: string): void => {
    const vendorData = [...this.state.itemDetails];

    (vendorData[index][field] as any) = value;

    this.setState({ itemDetails: vendorData });
  };

  // // Trigger IndexCreation 
  public async triggerSubmit() {
    const queryurl = this.props.context.pageContext.web.serverRelativeUrl + strings.queryList + this.props.wpproperties.FlowConnectionsListName;
    const flowName = "QatarCement_RFQSubmit"
    const filter = "Title eq '" + flowName + "'";
    const laUrl = await this.service.getItemsFilter(queryurl, filter);
    const postURL = laUrl[0].AppURL;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'MasterID': String(this.state.masterid),
      'ItemDetails': this.state.itemDetails


    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    const response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    if (response) {
      const responseJSON = await response.json();
      if (response.ok) {
        console.log("Response from Flow:", responseJSON);
        ToastService.success("RFQ Submitted successfully.");
        this.setState({ modalOverlay: { isOpen: false, Text: '' } });
      }
    }
  }


  public submitRFQDept = async () => {
    this.setState({ modalOverlay: { isOpen: true, Text: 'Submitting RFQ Dept Data...' } });

    await this.triggerSubmit();   // use your existing flow submit logic

    this.setState({ modalOverlay: { isOpen: false, Text: '' } });
  };



  public handleVendorResponseChange = (itemId: number, field: 'status' | 'price' | 'comments', value: string): void => {
    this.setState(prevState => ({
      vendorResponses: {
        ...prevState.vendorResponses,
        [itemId]: {
          // Preserve all existing fields for this itemId
          ...(prevState.vendorResponses[itemId] || {}),
          // Only update the specific field being changed
          [field]: value
        }
      }
    }));
  };


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
    const initiatorFlow = "QatarCement_RFQInitiatorSubmit";

    const queryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.FlowConnectionsListName;
    const filter = `Title eq '${initiatorFlow}'`;

    const laUrl = await this.service.getItemsFilter(queryurl, filter);
    const postURL = laUrl[0].AppURL;

    const headers = new Headers();
    headers.append("Content-type", "application/json");

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

    // try {
    const response = await this.props.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      { headers, body }
    );

    const json = await response.json();
    console.log("Initiator Response:", json);

    if (response.ok) {
      ToastService.success("Technical review submitted successfully!");
      // Optionally close the window after successful submission
      // setTimeout(() => this.closeWindow(), 2000);
    } else {
      ToastService.error("Failed to submit technical review. Please try again.");
    }
    // } catch (error) {
    //   console.error("Error submitting initiator data:", error);
    //   ToastService.error("An error occurred while submitting. Please try again.");
    // }
  }


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

    const headers = new Headers();
    headers.append("Content-type", "application/json");

    // Filter items assigned to current vendor
    const userEmailLower = this.state.currentUser.email.toLowerCase().trim();

    const vendorAssignedItems = this.state.itemDetails.filter(item => {
      if (!item.vendors) return false;
      const vendorList = item.vendors.toLowerCase();
      const emails = vendorList.split(/[,;]/).map(e => e.trim());
      return emails.some(email => email === userEmailLower || email.includes(userEmailLower));
    });

    // Add vendor responses + WorkflowDetailsId to filtered items
    const itemDetailsWithResponses = vendorAssignedItems.map((item) => ({
      ...item,
      WorkflowDetailsId: item.WorkflowDetailsId, // ✓ Include WorkflowDetailsId
      vendorStatus: this.state.vendorResponses[item.Id]?.status || '',
      vendorPrice: this.state.vendorResponses[item.Id]?.price || '',
      vendorComments: this.state.vendorResponses[item.Id]?.comments || '',
    }));

    const body: string = JSON.stringify({
      'TaskID': String(this.state.taskID),
      'VendorEmail': this.state.currentUser.email,
      'MasterID': String(this.state.masterid),
      'ItemDetails': itemDetailsWithResponses, // ✓ Now includes WorkflowDetailsId
    });

    console.log("Submitting Vendor Data:", itemDetailsWithResponses);

    const response = await this.props.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      { headers, body }
    );

    const json = await response.json();
    console.log("Vendor Response:", json);

    if (response.ok) {
      ToastService.success("Vendor response submitted!");
    } else {
      ToastService.error("Failed to submit vendor response. Please try again.");
    }
  }

  // Maintenance Manager Response Handler
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

  // Trigger Manager Submit
  public async triggerMaintenanceManagerSubmit() {
    const managerFlow = "QatarCement_RFQMaintenanceManagerSubmit";

    const queryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.FlowConnectionsListName;
    const filter = `Title eq '${managerFlow}'`;

    const laUrl = await this.service.getItemsFilter(queryurl, filter);
    const postURL = laUrl[0].AppURL;

    const headers = new Headers();
    headers.append("Content-type", "application/json");

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
        managerStatus: this.state.managerResponses[workflowDetailsId]?.status || '',
        managerComments: this.state.managerResponses[workflowDetailsId]?.comments || '',
      };
    });

    const body: string = JSON.stringify({
      'TaskID': String(this.state.taskID),
      'ManagerEmail': this.state.currentUser.email,
      'MasterID': String(this.state.masterid),
      'ItemDetails': itemDetailsWithManagerResponse,
    });

    console.log("Submitting Maintenance Manager Data:", itemDetailsWithManagerResponse);

    const response = await this.props.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      { headers, body }
    );

    const json = await response.json();
    console.log("Maintenance Manager Response:", json);

    if (response.ok) {
      ToastService.success("Maintenance Manager review submitted successfully!");
    } else {
      ToastService.error("Failed to submit maintenance manager review. Please try again.");
    }
  }

  // Submit Procurement Manager
  public submitProcurementManager = async () => {
    this.setState({ modalOverlay: { isOpen: true, Text: 'Submitting Procurement Manager Response...' } });
    await this.triggerProcurementManagerSubmit();
    this.setState({ modalOverlay: { isOpen: false, Text: '' } });
  };

  // Trigger Procurement Manager Submit
  public async triggerProcurementManagerSubmit() {
    const procurementManagerFlow = "QatarCement_RFQProcurementManagerSubmit";

    const queryurl = this.props.context.pageContext.web.serverRelativeUrl +
      strings.queryList + this.props.wpproperties.FlowConnectionsListName;
    const filter = `Title eq '${procurementManagerFlow}'`;

    const laUrl = await this.service.getItemsFilter(queryurl, filter);
    const postURL = laUrl[0].AppURL;

    const headers = new Headers();
    headers.append("Content-type", "application/json");

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
        procurementManagerStatus: this.state.procurementManagerResponses[workflowDetailsId]?.status || '',
        procurementManagerComments: this.state.procurementManagerResponses[workflowDetailsId]?.comments || '',
      };
    });

    const body: string = JSON.stringify({
      'TaskID': String(this.state.taskID),
      'ProcurementManagerEmail': this.state.currentUser.email,
      'MasterID': String(this.state.masterid),
      'ItemDetails': itemDetailsWithProcurementManagerResponse,
    });

    console.log("Submitting Procurement Manager Data:", itemDetailsWithProcurementManagerResponse);

    const response = await this.props.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      { headers, body }
    );

    const json = await response.json();
    console.log("Procurement Manager Response:", json);

    if (response.ok) {
      ToastService.success("Procurement Manager review submitted successfully!");
    } else {
      ToastService.error("Failed to submit procurement manager review. Please try again.");
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
              />
            </>}
            {/* Vendor View */}
            {this.state.userType === "Vendor" && (<>
              <VendorDetailsTable itemDetails={this.state.itemDetails}
                currentUserEmail={this.state.currentUser.email}
                onSubmitVendor={this.submitVendor}
                onCancel={this.closeWindow}
                onResponseChange={this.handleVendorResponseChange} vendorResponses={this.state.vendorResponses} />
            </>)}
            {/* Initiator View - NEW */}
            {this.state.userType === "Initiator" && (
              <InitiatorDetailsTable
                itemDetails={this.state.itemDetails}
                masterid={this.state.masterid}
                taskID={this.state.taskID}
                initiatorResponses={this.state.initiatorResponses}  // ✅
                onResponseChange={this.handleInitiatorResponseChange}  // ✅
                onSubmitInitiator={this.submitInitiator}  // ✅
                onCancel={this.closeWindow}  // ✅
              />
            )}
            {/* Manager View */}
            {this.state.userType === "Manager" && (
              <ManagerDetailsTable
                itemDetails={this.state.itemDetails}
                masterid={this.state.masterid}
                taskID={this.state.taskID}
                managerResponses={this.state.managerResponses}
                onResponseChange={this.handleMaintenanceManagerResponseChange}
                onSubmitManager={this.submitMaintenanceManager}
                onCancel={this.closeWindow}
              />
            )}

            {/* Procurement Manager View */}
            {/* {this.state.userType === "ProcurementManager" && (
              <ProcurementManagerDetailsTable
                itemDetails={this.state.itemDetails}
                masterid={this.state.masterid}
                taskID={this.state.taskID}
                procurementManagerResponses={this.state.procurementManagerResponses}
                onResponseChange={this.handleProcurementManagerResponseChange}
                onSubmitProcurementManager={this.submitProcurementManager}
                onCancel={this.closeWindow}
              />
            )} */}
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
