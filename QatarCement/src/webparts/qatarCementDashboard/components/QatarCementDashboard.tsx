import * as React from 'react';
import styles from './QatarCementDashboard.module.scss';
import type {
  IQatarCementDashboardProps,
  IQatarCementDashboardState
} from '../interfaces/IQatarCementDashboardProps';

import * as XLSX from 'xlsx';
import * as _ from 'lodash';

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

  /* =====================================================
     CORE LOGIC – PR → ItemCode → Items
     ===================================================== */
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

    const combinedItems: any[] = [];
    const groups: any[] = [];
    let startIndex = 0;

    for (const pr of prDetails) {
      const prItems = workflowItems.filter(
        (w: any) => String(w.prDetailsId) === String(pr.ID)
      );

      if (!prItems.length) continue;

      const itemGroups = _.groupBy(prItems, i =>
        (i.ItemCode || 'Unknown').trim()
      );

      const parentStart = startIndex;
      let parentCount = 0;
      const children: any[] = [];

      Object.keys(itemGroups).forEach(itemCode => {
        const items = itemGroups[itemCode];

        // ✅ ONLY Technically Accepted items
        const accepted = items.filter(
          i => i.InitiatorStatus === 'Technically Accepted'
        );

        // ✅ SKIP empty ItemCode groups
        if (!accepted.length) {
          return;
        }
        // Sort accepted by price
        accepted.sort((a, b) => Number(a.Price) - Number(b.Price));

        // Assign price rank
        accepted.forEach((item, idx) => {
          item.priceRank = idx;
        });

        // ❌ Do NOT include rejected at all
        const finalItems = accepted;

        children.push({
          key: `pr-${pr.ID}-item-${itemCode}`,
          name: itemCode,
          startIndex,
          count: finalItems.length,
          level: 1,
          isCollapsed: false
        });

        finalItems.forEach(item => combinedItems.push(item));

        startIndex += finalItems.length;
        parentCount += finalItems.length;
      });

      groups.push({
        key: `pr-${pr.ID}`,
        name: pr.PRNumber,
        startIndex: parentStart,
        count: parentCount,
        level: 0,
        isCollapsed: false,
        children
      });
    }

    this.setState({ listItems: combinedItems, groups });
  }

  /* =====================================================
     COLUMNS
     ===================================================== */
  private getColumns(): IColumn[] {
    return [
      { key: 'item', name: 'Item Code', fieldName: 'ItemCode', minWidth: 120 },
      { key: 'desc', name: 'Description', fieldName: 'Description', minWidth: 200 },
      { key: 'qty', name: 'Qty', fieldName: 'Qty', minWidth: 60 },
      { key: 'uom', name: 'UoM', fieldName: 'UoM', minWidth: 70 },
      { key: 'comments', name: 'Comments', fieldName: 'Comments', minWidth: 150 },
      {
        key: 'price',
        name: 'Price',
        fieldName: 'Price',
        minWidth: 100,
        onRender: (item: any) => {
          // ❌ Rejected → Always red
          if (item.InitiatorStatus === 'Technically Not Accepted') {
            return <span style={{ color: '#a80000', fontWeight: 600 }}>{item.Price}</span>;
          }

          // ✅ Accepted → Rank based
          let color = '#a80000';
          if (item.priceRank === 0) color = '#107c10';
          else if (item.priceRank === 1) color = '#ffb900';
          else if (item.isHighest) {
            color = '#a80000'; // 🔴 Highest ONLY
          }

          return <span style={{ color, fontWeight: 600 }}>{item.Price}</span>;
        }
      },
      { key: 'vendor', name: 'Vendor', fieldName: 'Vendor', minWidth: 160 },
      {
        key: 'status',
        name: 'Status',
        fieldName: 'InitiatorStatus',
        minWidth: 180,
        onRender: (item: any) => (
          <span
            style={{
              color:
                item.InitiatorStatus === 'Technically Not Accepted'
                  ? '#a80000'
                  : '#323130',
              fontWeight:
                item.InitiatorStatus === 'Technically Not Accepted' ? 600 : 400
            }}
          >
            {item.InitiatorStatus}
          </span>
        )
      }
    ];
  }


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
            borderRadius: 3
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
        {legendItem('#107c10', 'Lowest Accepted Price')}
        {legendItem('#ffb900', 'Second Best Accepted Price')}
        {legendItem('#a80000', 'Highest Price ')}
      </div>
    );
  }


  /* =====================================================
     ROW STYLING
     ===================================================== */
  private onRenderRow = (props?: IDetailsRowProps): JSX.Element | null => {
    if (!props) return null;

    const { itemIndex, group } = props;
    const relativeIndex =
      group && group.level === 1 ? itemIndex - group.startIndex : itemIndex;

    return (
      <DetailsRow
        {...props}
        styles={{
          root: {
            backgroundColor: relativeIndex % 2 === 0 ? '#ffffff' : '#f4f4f4',
            selectors: {
              ':hover': { backgroundColor: '#eaeaea' }
            }
          }
        }}
      />
    );
  };

  /* =====================================================
     EXPORT
     ===================================================== */
  private exportToExcel = (): void => {
    const data = (this.state.listItems || []).map(i => ({
      'Item Code': i.ItemCode,
      Description: i.Description,
      Qty: i.Qty,
      UoM: i.UoM,
      Comments: i.Comments,
      Price: i.Price,
      Vendor: i.Vendor,
      Status: i.InitiatorStatus
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Dashboard');
    XLSX.writeFile(wb, 'QatarCementDashboard.xlsx');
  };

  /* =====================================================
     RENDER
     ===================================================== */
  public render(): React.ReactElement<IQatarCementDashboardProps> {
    return (
      <section className={styles.container}>
        <div className={styles.formpopup}>
          <div className={styles.formheader}>
            <div className={styles.formtitle}>{this.props.wpproperties.webpartTitle}</div>
          </div>
          <div className={styles.formbody}>

            <div style={{ textAlign: 'right', marginBottom: 10 }}>
              <button
                onClick={this.exportToExcel}
                style={{
                  padding: '6px 14px',
                  backgroundColor: '#107c10',
                  color: '#fff',
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
              onRenderRow={this.onRenderRow}
            />
            {this.renderPriceLegend()}

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
