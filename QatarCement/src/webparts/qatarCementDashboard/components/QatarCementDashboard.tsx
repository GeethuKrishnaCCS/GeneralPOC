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
    const prItemsListName = this.props.wpproperties?.PRItemSpecficationsListName || 'PRItemSpecifications';

    const webUrl: string = this.props.context.pageContext.web.serverRelativeUrl || this.props.context.pageContext.web.absoluteUrl;
    const prDetailsQuery = `${webUrl}/Lists/${prDetailsListName}`;
    const prItemsQuery = `${webUrl}/Lists/${prItemsListName}`;

    // fetch PR Details (include extra fields to display in parent row)
    const prDetailsSelect = "ID,PRNumber,Department,Priority,DueDate,PRInitiator/ID,PRInitiator/Title,BusinessJustification,Created,Modified";
    const prDetailsExpand = "PRInitiator";
    const prDetails: any[] = await this.service.getSelectExpand(prDetailsQuery, prDetailsSelect, prDetailsExpand) || [];

    // fetch PR Item Specifications and include PRDetails lookup; Vendors is a plain text field
    const prItemsSelect = "ID,Title,ItemCode,Description,Qty,UoM,PRDetailsID/ID,PRDetailsID/PRNumber,Created,Modified,Vendors";
    const prItemsExpand = "PRDetailsID";
    const prItemsRaw: any[] = await this.service.getSelectExpand(prItemsQuery, prItemsSelect, prItemsExpand) || [];

    // normalize PR item children
    const prItems: any[] = prItemsRaw.map((it: any) => {
      const lookup = it.PRDetailsID || it.PRDetails || null;
      const prDetailsId = lookup ? (lookup.ID || lookup.Id || lookup.Id || lookup.ID) : null;
      // normalize Vendors (could be array or single object)
      let vendors = '';
      const v = it.Vendors || null;
      if (Array.isArray(v)) {
        vendors = v.map((vv: any) => vv.Title || vv.EMail || vv.Email || '').filter((x: string) => x).join(', ');
      } else if (v && typeof v === 'object') {
        vendors = v.Title || v.EMail || v.Email || '';
      } else if (typeof it.Vendors === 'string') {
        vendors = it.Vendors;
      }

      return {
        ID: it.ID,
        Title: it.Title,
        ItemCode: it.ItemCode,
        Description: it.Description,
        Qty: it.Qty,
        UoM: it.UoM,
        prDetailsId: prDetailsId,
        Vendors: vendors,
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
      // children
      const children = prItems.filter(pi => String(pi.prDetailsId) === String(pd.ID));
      combinedItems.push(parent);
      for (const c of children) {
        combinedItems.push({ ...c, isParent: false });
      }
      const count = 1 + children.length;
      // default groups collapsed so sub items are hidden until user clicks the toggle
      groups.push({ key: String(pd.ID), name: pd.PRNumber || `PR ${pd.ID}`, startIndex: cumulativeCount, count: count, level: 0, isCollapsed: true });
      cumulativeCount += count;
    }

    this.setState({ listItems: combinedItems, groups: groups, itemsPerPage: 10 });
  }

  private getPRDetailsColumns(): IColumn[] {
    return [
      { key: 'prnumber', name: 'PRNumber', fieldName: 'PRNumber', minWidth: 100, maxWidth: 120, isResizable: true },
      { key: 'department', name: 'Department', fieldName: 'Department', minWidth: 100, maxWidth: 120, isResizable: true },
      {
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
      },
      { key: 'duedate', name: 'DueDate', fieldName: 'DueDate', minWidth: 100, maxWidth: 120, isResizable: true },
      { key: 'prinitiator', name: 'PRInitiator', fieldName: 'PRInitiator', minWidth: 100, maxWidth: 120, isResizable: true },
      { key: 'businessjustification', name: 'BusinessJustification', fieldName: 'BusinessJustification', minWidth: 150, maxWidth: 200, isResizable: true },
      //{ key: 'created', name: 'Created', fieldName: 'Created', minWidth: 100, maxWidth: 120, isResizable: true },
      //{ key: 'modified', name: 'Modified', fieldName: 'Modified', minWidth: 100, maxWidth: 120, isResizable: true },
      //{ key: 'id', name: 'ID', fieldName: 'ID', minWidth: 40, maxWidth: 60, isResizable: true },
    ];
  }

  private getPRItemSpecColumns(): IColumn[] {
    return [
      { key: 'itemcode', name: 'ItemCode', fieldName: 'ItemCode', minWidth: 80, maxWidth: 120, isResizable: true },
      { key: 'description', name: 'Description', fieldName: 'Description', minWidth: 200, maxWidth: 300, isResizable: true },
      { key: 'qty', name: 'Qty', fieldName: 'Qty', minWidth: 60, maxWidth: 80, isResizable: true },
      { key: 'uom', name: 'UoM', fieldName: 'UoM', minWidth: 60, maxWidth: 80, isResizable: true },
      //{ key: 'id', name: 'ID', fieldName: 'ID', minWidth: 40, maxWidth: 60, isResizable: true },
      { key: 'vendors', name: 'Vendors', fieldName: 'Vendors', minWidth: 150, maxWidth: 300, isResizable: true },
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




