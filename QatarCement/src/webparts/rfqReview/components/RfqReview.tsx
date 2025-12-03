import * as React from 'react';
import styles from './RfqReview.module.scss';
import type { IItemData, IRfqReviewProps, IRfqReviewState } from '../interfaces/IRfqReviewProps';
import { RfqReviewService } from '../services/RfqReviewService';
import ModalOverlay from '../../../shared/controls/Overlay/Overlay';
import { PrimaryButton, TextField } from '@fluentui/react';
import ToastService from '../../../shared/controls/Toast/Toast';
import * as strings from 'RfqReviewWebPartStrings';
import { HttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import * as moment from 'moment';
import RFQDeptDetailsTable from './RFQDeptDetailsTable';
import VendorDetailsTable from './VendorDetailsTable';
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
      userType: '',
      prNumber: '',
      department: '',
      priority: '',
      dueDate: '',
      prInitiator: '',
      businessJustification: '',
      itemDetails: [],
      vendorOptions: [],
      masterid: ''
    };
    this.service = new RfqReviewService(this.props.context, this.props.context.pageContext.web.absoluteUrl);
    this.validateURLParams = this.validateURLParams.bind(this);
    this.bindVendorData = this.bindVendorData.bind(this);
    this.bindWorkflowData = this.bindWorkflowData.bind(this);
    this.checkUserGroup = this.checkUserGroup.bind(this);
    this.bindMasterData = this.bindMasterData.bind(this);
    this.handleChange = this.handleChange.bind(this);
    this.onsubmit = this.onsubmit.bind(this);
    this.onCancel = this.onCancel.bind(this);
    this.triggerSubmit = this.triggerSubmit.bind(this);

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
    const taskid = params.get('TID'); // NEW

    if (masterid !== "" && masterid !== null && masterid !== undefined) {
      if (taskid !== "" && taskid !== null && taskid !== undefined) {
        // 👉 CASE 2: MID + TID
        await this.bindWorkflowData(masterid, taskid);

      } else {
        // 👉 CASE 1: Only MID
        await this.checkUserGroup(masterid);

      }
    } else {
      ToastService.error("Invalid URL parameters. Please check and try again.");
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
          this.setState({ userType: "Vendors" })
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

  // Bind Master Data
  public async bindMasterData(masterid: any) {
    let masterdata: any;
    let itemdetaildata: any[];
    let itemdetaildataitems: any[] = [];
    //Fetch master index item
    const masterqueryurl = this.props.context.pageContext.web.serverRelativeUrl + strings.queryList + this.props.wpproperties.PRDetailsListName;
    const select = "*,PRInitiator/ID,PRInitiator/Title,PRInitiator/EMail";
    const expand = "PRInitiator";
    // Fetch master item details
    const itemdetailqueryurl = this.props.context.pageContext.web.serverRelativeUrl + strings.queryList + this.props.wpproperties.PRItemSpecficationsListName;
    const itemfilter = "PRDetailsIDId eq '" + Number(masterid) + "'"; // Filter to get the specific DMS ID
    try {
      masterdata = await this.service.getItemsByIdSelectExpand(masterqueryurl, Number(masterid), select, expand);
      console.log("masterdata" + masterdata);
      itemdetaildata = await this.service.getPagedFilterListItems(itemdetailqueryurl, itemfilter);
      console.log("itemdetaildata" + itemdetaildata);
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
            vendors: item.Vendors
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
      console.error("Error fetching DMS data:", error);
      ToastService.error("Failed to fetch data. Please try again later.");
    }
  }
  // Handle change for vendor table data
  public handleChange = (index: number, field: keyof IItemData, value: string): void => {
    const vendorData = [...this.state.itemDetails];

    vendorData[index][field] = value;

    this.setState({ itemDetails: vendorData });
  };
  //on submit
  public onsubmit = async (): Promise<void> => {
    this.setState({ modalOverlay: { isOpen: true, Text: 'Submitting...' } });
    await this.triggerSubmit();
  };
  // Trigger IndexCreation 
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
  //on cancel
  public onCancel = (): void => {
    window.close();
  }
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
            {/* Inside the render method, replace the table code with this: */}
            {this.state.userType === "RFQDept" && <RFQDeptDetailsTable
              itemDetails={this.state.itemDetails}
              vendorOptions={this.state.vendorOptions}
              handleChange={this.handleChange}
            />}
            <VendorDetailsTable
              itemDetails={this.state.itemDetails}
            />
            {/* <div className={styles.row}>
              <div className={styles.col12}>
                {this.state.itemDetails.length > 0 &&
                  <div className={styles.doctable}>
                    <table className={styles.table} >
                      <tr className={styles.tr}>
                        <th className={styles.th}>Sl No</th>
                        <th className={styles.th}>ItemCode</th>
                        <th className={styles.th}>Description</th>
                        <th className={styles.th}>Quantity</th>
                        <th className={styles.th}>UOM</th>
                        <th className={styles.th}>Vendors</th>
                      </tr>
                      {this.state.itemDetails.map((item, key) => {
                        return (
                          <tr key={key} className={styles.tr}>
                            <td className={styles.th}>{key + 1}</td>
                            <td className={styles.th}><TextField value={item.ItemCode} readOnly /></td>
                            <td className={styles.th}><TooltipHost content={item.Description}><TextField value={item.Description} readOnly /></TooltipHost></td>
                            <td className={styles.th}><TextField value={item.Quantity} readOnly /></td>
                            <td className={styles.th}><TextField value={item.UOM} readOnly /></td>
                            <td className={styles.th}>
                              <div className={styles.vendorCell}>
                                <Dropdown
                                  placeholder="Select Vendors"
                                  multiSelect
                                  options={this.state.vendorOptions}
                                  selectedKeys={item.vendors ? item.vendors.split(',') : []}
                                  onChange={(e, option) => {
                                    let updated = [...(item.vendors ? item.vendors.split(',') : [])];

                                    if (option?.selected) {
                                      updated.push(option.key as string);
                                    } else {
                                      updated = updated.filter(v => v !== option?.key);
                                    }

                                    this.handleChange(key, 'vendors', updated.join(','));
                                  }}
                                />
                              </div>
                            </td>
                          </tr>
                        )
                      })}
                    </table>
                  </div>
                }
              </div>
            </div> */}
            <div className={styles.row}>
              <div className={styles.col12}>
                <div className={styles.rgtalign}>
                  <PrimaryButton className={styles.btn} onClick={this.onsubmit}>Submit</PrimaryButton >
                  <PrimaryButton className={styles.btn} onClick={this.onCancel}>Close</PrimaryButton >
                </div>
              </div>
            </div>
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
