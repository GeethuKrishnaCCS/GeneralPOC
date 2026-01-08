import * as React from 'react';

import styles from './QatarCementDashboard.module.scss';

import type {

  IQatarCementDashboardProps,

  IQatarCementDashboardState

} from '../interfaces/IQatarCementDashboardProps';
import * as XLSX from 'xlsx';


import { QatarCementDashboardService } from '../services/QatarCementDashboardService';

import ModalOverlay from '../../../shared/controls/Overlay/Overlay';

import {

  DetailsList,

  DetailsListLayoutMode,

  SelectionMode,

  IColumn,
  DetailsRow,
  IDetailsRowProps

} from '@fluentui/react/lib/DetailsList';

import * as _ from 'lodash';

interface QatarCementDashboardStateWithExpand

  extends IQatarCementDashboardState {

  listItems?: any[];

  groups?: any[];

}

export default class QatarCementDashboard extends React.Component<

  IQatarCementDashboardProps,

  QatarCementDashboardStateWithExpand
> {

  private service: QatarCementDashboardService;

  constructor(props: IQatarCementDashboardProps) {

    super(props);

    this.state = {

      listItems: [],

      groups: [],

      modalOverlay: {

        isOpen: false,

        Text: ''

      }

    };

    this.service = new QatarCementDashboardService(

      this.props.context,

      this.props.context.pageContext.web.absoluteUrl

    );

  }

  public async componentDidMount(): Promise<void> {

    this.setState({ modalOverlay: { isOpen: true, Text: 'Loading...' } });

    await this.getPRDetailsAndItems();

    this.setState({ modalOverlay: { isOpen: false, Text: '' } });

  }

  /**

   * CORE LOGIC – PR → Item Code → Individual Rows

   */

  // ...existing code...
  /**
   * CORE LOGIC – PR → Item Code → Individual Rows
   */
  private async getPRDetailsAndItems(): Promise<void> {
    const prDetailsListName =
      this.props.wpproperties?.PRDetailsListName || 'PRDetails';
    const workflowListName =
      this.props.wpproperties?.WorkflowDetailsListName || 'WorkflowDetails';

    const webUrl = this.props.context.pageContext.web.serverRelativeUrl;

    const prDetails = await this.service.getSelectExpand(
      `${webUrl}/Lists/${prDetailsListName}`,
      'ID,PRNumber',
      ''
    );

    const workflowItemsRaw = await this.service.getSelectExpand(
      `${webUrl}/Lists/${workflowListName}`,
      'ID,Vendor,ItemCode,Description,Qty,UoM,Comments,Price,PRDetailID,InitiatorStatus',
      ''
    );

    const workflowItems = workflowItemsRaw.map((i: any) => ({
      ...i,
      prDetailsId: i.PRDetailID
    }));

    const combinedItems: any[] = []; // only actual item rows live here (no placeholders)
    const groups: any[] = [];
    let startIndex = 0;

    for (const pr of prDetails) {
      const prItems = workflowItems.filter(
        (w: any) => String(w.prDetailsId) === String(pr.ID)
      );

      if (!prItems.length) continue;

      // Group by ItemCode
      const itemGroups = _.groupBy(prItems, (v: any) =>
        (v.ItemCode || 'Unknown Item').trim()
      );

      const parentStart = startIndex;
      const children: any[] = [];
      let parentItemCount = 0;

      Object.keys(itemGroups).forEach((itemCode) => {
        const items = itemGroups[itemCode];

        // Sort by price (LOWEST first)
        const sortedItems = [...items].sort(
          (a, b) => Number(a.Price) - Number(b.Price)
        );

        // child group's startIndex must point to where its items will be inserted
        const childStart = startIndex;

        // push child group (level 1)
        children.push({
          key: `pr-${pr.ID}-item-${itemCode}`,
          name: itemCode,
          startIndex: childStart,
          count: sortedItems.length,
          level: 1,
          isCollapsed: false
        });

        // push actual item rows into combinedItems
        sortedItems.forEach((item, idx) => {
          combinedItems.push({
            ...item,
            priceRank: idx // 0 = lowest, 1 = second, 2+ = third
          });
        });

        // advance startIndex and parent counters
        startIndex += sortedItems.length;
        parentItemCount += sortedItems.length;
      });

      // parent group (level 0) — children array so expanding PR shows only ItemCode headers
      groups.push({
        key: `pr-${pr.ID}`,
        name: `${pr.PRNumber}`,
        startIndex: parentStart,
        // count must equal total number of items under this PR (sum of child counts)
        count: parentItemCount,
        level: 0,
        isCollapsed: false,
        children
      });
    }

    // ensure groups ordered by startIndex
    groups.sort((a: any, b: any) => a.startIndex - b.startIndex);

    this.setState({ listItems: combinedItems, groups });
  }
  // ...existing code...

  /**

   * Columns - only shown at the detail level

   */

  private getColumns(): IColumn[] {

    return [



      {

        key: 'itemcode',

        name: 'Item Code',

        fieldName: 'ItemCode',

        minWidth: 100,

        onRender: (item: any) => {

          if (item.isPlaceholder) return null;

          return <span>{item.ItemCode}</span>;

        }

      },

      {

        key: 'description',

        name: 'Description',

        fieldName: 'Description',

        minWidth: 200,

        onRender: (item: any) => {

          if (item.isPlaceholder) return null;

          return <span>{item.Description}</span>;

        }

      },

      {

        key: 'qty',

        name: 'Qty',

        fieldName: 'Qty',

        minWidth: 60,

        onRender: (item: any) => {

          if (item.isPlaceholder) return null;

          return <span>{item.Qty}</span>;

        }

      },

      {

        key: 'uom',

        name: 'UoM',

        fieldName: 'UoM',

        minWidth: 60,

        onRender: (item: any) => {

          if (item.isPlaceholder) return null;

          return <span>{item.UoM}</span>;

        }

      },

      {

        key: 'comments',

        name: 'Comments',

        fieldName: 'Comments',

        minWidth: 150,

        onRender: (item: any) => {

          if (item.isPlaceholder) return null;

          return <span>{item.Comments}</span>;

        }

      },

      {

        key: 'price',

        name: 'Price',

        fieldName: 'Price',

        minWidth: 100,

        onRender: (item: any) => {

          if (item.isPlaceholder) return null;

          let color = '#323130';

          if (item.priceRank === 0) {

            color = '#107c10'; // 🟢 Lowest

          } else if (item.priceRank === 1) {

            color = '#ffb900'; // 🟡 Second

          } else {

            color = '#a80000'; // 🔴 Third & above

          }

          return (
            <span style={{ color, fontWeight: 600 }}>

              {item.Price}
            </span>

          );

        }

      },
      {

        key: 'vendor',

        name: 'Vendor',

        fieldName: 'Vendor',

        minWidth: 150,

        onRender: (item: any) => {

          // Hide placeholder rows

          if (item.isPlaceholder) return null;

          return <span>{item.Vendor}</span>;

        }

      },
      {
        key: 'initiatorStatus',
        name: 'Status',
        fieldName: 'InitiatorStatus',
        minWidth: 170,
        onRender: (item: any) => {
          if (item.isPlaceholder) return null;

          const isRejected =
            item.InitiatorStatus === 'Technically Not Accepted';

          return (
            <span
              style={{
                color: isRejected ? '#a80000' : '#323130',
                fontWeight: isRejected ? 600 : 400
              }}
            >
              {item.InitiatorStatus}
            </span>
          );
        }
      }



    ];

  }

  private onRenderRow = (props?: IDetailsRowProps): JSX.Element | null => {
    if (!props) return null;

    const { itemIndex, group } = props;

    // Reset zebra inside each ItemCode group
    const relativeIndex =
      group && group.level === 1
        ? itemIndex - group.startIndex
        : itemIndex;

    const isEvenRow = relativeIndex % 2 === 0;

    const backgroundColor = isEvenRow
      ? '#ffffff'   // white
      : '#f4f4f4';  // light ash

    return (
      <DetailsRow
        {...props}
        styles={{
          root: {
            backgroundColor,
            selectors: {
              ':hover': {
                backgroundColor: '#eaeaea'
              }
            }
          }
        }}
      />
    );
  };
  private renderPriceLegend(): JSX.Element {
    const legendItem = (color: string, text: string) => (
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          marginRight: 24
        }}
      >
        <span
          style={{
            width: 14,
            height: 14,
            backgroundColor: color,
            display: 'inline-block',
            marginRight: 8,
            borderRadius: 2
          }}
        />
        <span style={{ fontSize: 13 }}>{text}</span>
      </div>
    );

    return (
      <div
        style={{
          display: 'flex',
          marginTop: 12,
          padding: '8px 12px',
          background: '#faf9f8',
          border: '1px solid #edebe9',
          borderRadius: 4
        }}
      >
        {legendItem('#107c10', 'Lowest bid')}
        {legendItem('#ffb900', 'Second best bid')}
        {legendItem('#a80000', 'Third best bids')}
      </div>
    );
  }

  private exportToExcel = (): void => {
    const items = this.state.listItems || [];

    if (!items.length) return;

    // Prepare data for Excel
    const excelData = items.map((item: any) => ({
      'Item Code': item.ItemCode,
      'Description': item.Description,
      'Qty': item.Qty,
      'UoM': item.UoM,
      'Comments': item.Comments,
      'Price': item.Price,
      'Vendor': item.Vendor,
      'Status': item.InitiatorStatus
    }));

    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(excelData);

    // Create workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Dashboard');

    // Export file
    XLSX.writeFile(workbook, 'QatarCementDashboard.xlsx');
  };



  public render(): React.ReactElement<IQatarCementDashboardProps> {

    return (

      <section className={styles.container}>
        <div className={styles.formpopup}>
          <div className={styles.formheader}>
            <div className={styles.formtitle}>{this.props.wpproperties.webpartTitle}</div>
          </div>
          <div className={styles.formbody}>


            <div style={{ marginBottom: 10, textAlign: 'right' }}>
              <button
                onClick={this.exportToExcel}
                style={{
                  padding: '6px 14px',
                  backgroundColor: '#107c10',
                  color: '#ffffff',
                  border: 'none',
                  borderRadius: 4,
                  cursor: 'pointer'
                }}
              >
                Export to Excel
              </button>
            </div>


            <DetailsList

              items={this.state.listItems || []}

              columns={this.getColumns()}

              layoutMode={DetailsListLayoutMode.justified}

              selectionMode={SelectionMode.none}

              groups={this.state.groups}
              groupProps={{
                showEmptyGroups: false
              }}

              onRenderRow={this.onRenderRow}   // ✅ ADD THIS LINE
            />
            {this.renderPriceLegend()}   {/* ✅ PRICE LEGEND */}

            <ModalOverlay

              isModalOpen={this.state.modalOverlay.isOpen}

              modalText={this.state.modalOverlay.Text}

            />
          </div>
        </div>
      </section>

    );

  }
}
