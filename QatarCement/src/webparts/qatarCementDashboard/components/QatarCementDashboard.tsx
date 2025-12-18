import * as React from 'react';
import styles from './QatarCementDashboard.module.scss';
import type { IQatarCementDashboardProps, IQatarCementDashboardState } from '../interfaces/IQatarCementDashboardProps';

interface QatarCementDashboardStateWithExpand extends IQatarCementDashboardState {
  listItems?: any[];
  groups?: any[];
  itemsPerPage?: number;
  tabularTotalPages?: number;
  groupedTotalPages?: number;
  expandedPRs?: { [prId: string]: boolean };
  prItemCache?: { [prId: string]: any[] };
}
import { QatarCementDashboardService } from '../services/QatarCementDashboardService';
import ModalOverlay from '../../../shared/controls/Overlay/Overlay';
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';

import * as _ from 'lodash';

export default class QatarCementDashboard extends React.Component<IQatarCementDashboardProps, QatarCementDashboardStateWithExpand, {}> {
  private service: QatarCementDashboardService;
  constructor(props: IQatarCementDashboardProps) {
    super(props);
    this.state = {
      modalOverlay: {
        isOpen: false,
        Text: ''
      }
    };
    this.service = new QatarCementDashboardService(this.props.context, this.props.context.pageContext.web.absoluteUrl);
  }
  public async componentDidMount(): Promise<void> {
    this.setState({ modalOverlay: { isOpen: true, Text: 'Loading...' } });
    const user = await this.service.getCurrentUser();
    console.log(user);
    // load PR lists and grouped items
    try {
      await this.getPRDetailsAndItems();
    } catch (err) {
      console.error('Error loading PR lists', err);
    }
    this.setState({ modalOverlay: { isOpen: false, Text: '' } });
  }

  private formatDate(value?: string | Date): string {
    if (!value) return '';
    const d = new Date(value as any);
    if (isNaN(d.getTime())) return '';
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  }



  /**
   * Fetch PR Details and PR Item Specifications and build grouped items for DetailsList
   */
  private async getPRDetailsAndItems(): Promise<void> {
    // list names are expected from web part properties, fallback to defaults
    const prDetailsListName = this.props.wpproperties?.PRDetailsListName || 'PRDetails';
    //const prItemsListName = this.props.wpproperties?.PRItemSpecficationsListName || 'PRItemSpecifications';
    const prItemsListName = this.props.wpproperties?.WorkflowDetailsListName || 'WorkflowDetails';

    const webUrl: string = this.props.context.pageContext.web.serverRelativeUrl || this.props.context.pageContext.web.absoluteUrl;
    const prDetailsQuery = `${webUrl}/Lists/${prDetailsListName}`;
    const prItemsQuery = `${webUrl}/Lists/${prItemsListName}`;

    // fetch PR Details (include extra fields to display in parent row)
    const prDetailsSelect = "ID,PRNumber,Department,Priority,DueDate,PRInitiator/ID,PRInitiator/Title,BusinessJustification,Created,Modified";
    const prDetailsExpand = "PRInitiator";
    const prDetails: any[] = await this.service.getSelectExpand(prDetailsQuery, prDetailsSelect, prDetailsExpand) || [];

    // fetch PR Item Specifications and include PRDetails lookup; Vendors is a plain text field
    //const prItemsSelect = "ID,Title,ItemCode,Description,Qty,UoM,PRDetailsID/ID,PRDetailsID/PRNumber,Created,Modified,Vendors";
    //const prItemsExpand = "PRDetailsID";

    const prItemsSelect =
      "ID,Title,Vendor,PRDetailID,PRItemID,TaskID,Status,Comments,Price,VendorDescription,ItemCode,Description,Qty,UoM,Created,Modified";

    const prItemsExpand = ""; // WorkflowDetails has NO lookup fields to expand


    const prItemsRaw: any[] = await this.service.getSelectExpand(prItemsQuery, prItemsSelect, prItemsExpand) || [];

    const prItems: any[] = prItemsRaw.map((it: any) => {
      return {
        ID: it.ID,
        Title: it.Title,
        Vendor: it.Vendor,
        PRDetailID: it.PRDetailID,
        PRItemID: it.PRItemID,
        TaskID: it.TaskID,
        Status: it.Status,
        Comments: it.Comments,
        Price: it.Price,
        VendorDescription: it.VendorDescription,
        ItemCode: it.ItemCode,
        Description: it.Description,
        Qty: it.Qty,
        UoM: it.UoM,
        prDetailsId: it.PRDetailID,  // 🔥 THIS CONNECTS CHILD TO PARENT
        Created: it.Created ? this.formatDate(it.Created) : '',
        Modified: it.Modified ? this.formatDate(it.Modified) : ''
      };
    });


    // Build combined list: for each PRDetail create a parent row then its child rows
    const combinedItems: any[] = [];
    const groups: any[] = [];
    let cumulativeCount = 0;
    for (const pd of prDetails) {
      const parent = {
        isParent: true,
        ID: pd.ID,
        PRNumber: pd.PRNumber,
        Department: pd.Department,
        Priority: pd.Priority,
        DueDate: pd.DueDate ? this.formatDate(pd.DueDate) : '',
        PRInitiator: pd.PRInitiator ? (pd.PRInitiator.Title || '') : '',
        BusinessJustification: pd.BusinessJustification || '',
        Created: pd.Created ? this.formatDate(pd.Created) : '',
        Modified: pd.Modified ? this.formatDate(pd.Modified) : ''
      };

      // children for this PR
      const children = prItems.filter(pi => String(pi.prDetailsId) === String(pd.ID));

      // PR top-level group (level 0) — includes parent + all its children
      const prGroupCount = 1 + children.length;
      groups.push({
        key: `pr-${pd.ID}`,
        name: `${pd.PRNumber || `PR ${pd.ID}`} (${children.length})`,
        startIndex: cumulativeCount,
        count: prGroupCount,
        level: 0,
        isCollapsed: true
      });

      // push parent row
      combinedItems.push(parent);
      cumulativeCount += 1;

      // group children by Vendor (normalize empty -> 'Unknown')
      const vendorGroups = _.groupBy(children, (c: any) => {
        const v = (c.Vendor || c.Vendors || '').toString().trim();
        return v.length ? v : 'Unknown Vendor';
      });

      // for each vendor group add a level-1 group and push the child rows
      for (const vendorName of Object.keys(vendorGroups)) {
        const vendorItems = vendorGroups[vendorName];
        const vendorKey = `pr-${pd.ID}-vendor-${encodeURIComponent(vendorName)}`;
        groups.push({
          key: vendorKey,
          name: vendorName,
          startIndex: cumulativeCount,
          count: vendorItems.length,
          level: 1,
          isCollapsed: true
        });

        for (const ci of vendorItems) {
          combinedItems.push({ ...ci, isParent: false });
          cumulativeCount += 1;
        }
      }
    }
    // ...existing code...
    this.setState({ listItems: combinedItems, groups: groups, itemsPerPage: 10 });
  }

  private getPRDetailsColumns(): IColumn[] {
    return [
      { key: 'title', name: 'Title', fieldName: 'Title', minWidth: 100, maxWidth: 120, isResizable: true },
      /*{
        key: 'priority', name: 'Priority', fieldName: 'Priority', minWidth: 80, maxWidth: 100, isResizable: true,
        onRender: (item: any) => {
          const value = (item.Priority || '').toString();
          let bg = '#107c10';
          let color = '#ffffff';
          if (/high/i.test(value)) { bg = '#e81123'; color = '#ffffff'; }
          else if (/medium/i.test(value)) { bg = '#ff8c00'; color = '#ffffff'; }
          else if (/low/i.test(value)) { bg = '#107c10'; color = '#ffffff'; }
          return (
            <span style={{ background: bg, color: color, padding: '4px 10px', borderRadius: 12, fontSize: 12, display: 'inline-block' }}>
              {value}
            </span>
          );
        }
      },*/
      { key: 'itemCode', name: 'ItemCode', fieldName: 'ItemCode', minWidth: 100, maxWidth: 120, isResizable: true },
      { key: 'description', name: 'Description', fieldName: 'Description', minWidth: 100, maxWidth: 120, isResizable: true },
      { key: 'qty', name: 'Qty', fieldName: 'Qty', minWidth: 150, maxWidth: 200, isResizable: true },
      { key: 'uoM', name: 'UoM', fieldName: 'UoM', minWidth: 100, maxWidth: 120, isResizable: true },
      { key: 'comments', name: 'Comments', fieldName: 'Comments', minWidth: 100, maxWidth: 120, isResizable: true },
      { key: 'price', name: 'Price', fieldName: 'Price', minWidth: 40, maxWidth: 60, isResizable: true },
      { key: 'vendor', name: 'Vendor', fieldName: 'Vendor', minWidth: 100, maxWidth: 120, isResizable: true },
    ];
  }

  /*private getPRItemSpecColumns(): IColumn[] {
      return [
        { key: 'itemcode', name: 'ItemCode', fieldName: 'ItemCode', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'description', name: 'Description', fieldName: 'Description', minWidth: 200, maxWidth: 300, isResizable: true },
        { key: 'qty', name: 'Qty', fieldName: 'Qty', minWidth: 60, maxWidth: 80, isResizable: true },
        { key: 'uom', name: 'UoM', fieldName: 'UoM', minWidth: 60, maxWidth: 80, isResizable: true },
        //{ key: 'id', name: 'ID', fieldName: 'ID', minWidth: 40, maxWidth: 60, isResizable: true },
        { key: 'vendors', name: 'Vendors', fieldName: 'Vendors', minWidth: 150, maxWidth: 300, isResizable: true },
      ];
    } */

  private getPRItemSpecColumns(): IColumn[] {
    return [
      //{ key: 'pritemid', name: 'PRItemID', fieldName: 'PRItemID', minWidth: 80, maxWidth: 100, isResizable: true },
      { key: 'title', name: 'Title', fieldName: 'Title', minWidth: 80, maxWidth: 100, isResizable: true },
      { key: 'itemcode', name: 'Item Code', fieldName: 'ItemCode', minWidth: 80, maxWidth: 120, isResizable: true },
      { key: 'description', name: 'Description', fieldName: 'Description', minWidth: 200, maxWidth: 300, isResizable: true },
      { key: 'qty', name: 'Qty', fieldName: 'Qty', minWidth: 60, maxWidth: 80, isResizable: true },
      { key: 'uom', name: 'UoM', fieldName: 'UoM', minWidth: 60, maxWidth: 80, isResizable: true },
      { key: 'comments', name: 'Comments', fieldName: 'Comments', minWidth: 200, maxWidth: 350, isResizable: true },
      { key: 'price', name: 'Price', fieldName: 'Price', minWidth: 80, maxWidth: 120, isResizable: true },
      { key: 'vendor', name: 'Vendor', fieldName: 'Vendor', minWidth: 150, maxWidth: 200, isResizable: true },

      //{ key: 'vendordesc', name: 'Vendor Description', fieldName: 'VendorDescription', minWidth: 150, maxWidth: 250, isResizable: true },


      //{ key: 'status', name: 'Status', fieldName: 'Status', minWidth: 100, maxWidth: 150, isResizable: true },
    ];
  }



  public render(): React.ReactElement<IQatarCementDashboardProps> {
    return (
      <section className={styles.container}>
        <div className={styles.formpopup}>
          <div className={styles.formheader}>
            <div className={styles.formtitle}>{this.props.wpproperties.webpartTitle}</div>
          </div>
          <div className={styles.formbody}>
            {this.state?.listItems && this.state.listItems.length > 0 ? (
              <DetailsList
                items={this.state.listItems}
                columns={this.getPRDetailsColumns()}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                selectionMode={SelectionMode.none}
                groups={this.state.groups}
                groupProps={{
                  showEmptyGroups: false,
                  // ensure all groups are collapsed on first render; clicking the group arrow will toggle expansion
                  isAllGroupsCollapsed: true
                }}
                onRenderRow={(props, defaultRender) => {
                  if (!props) return null;
                  const item = props.item;
                  if (item.isParent) {
                    return defaultRender ? defaultRender(props) : null;
                  } else {
                    // Render child row with PRItemSpecifications columns
                    return (
                      <DetailsList
                        items={[item]}
                        columns={this.getPRItemSpecColumns()}
                        setKey={`sub-${item.ID}`}
                        layoutMode={DetailsListLayoutMode.justified}
                        isHeaderVisible={false}
                        selectionMode={SelectionMode.none}
                        styles={{ root: { background: '#f9f9f9' } }}
                      />
                    );
                  }
                }}
              />
            ) : (
              <div>No PR items to display</div>
            )}
          </div>
        </div>
        <ModalOverlay
          isModalOpen={this.state.modalOverlay.isOpen}
          modalText={this.state.modalOverlay.Text}
        />
      </section>
    );
  }
}




