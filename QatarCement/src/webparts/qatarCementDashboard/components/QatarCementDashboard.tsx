import * as React from 'react';
import styles from './QatarCementDashboard.module.scss';
import type {
  IQatarCementDashboardProps,
  IQatarCementDashboardState
} from '../interfaces/IQatarCementDashboardProps';
import { QatarCementDashboardService } from '../services/QatarCementDashboardService';
import ModalOverlay from '../../../shared/controls/Overlay/Overlay';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn
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
   * CORE LOGIC – PR → Vendor → Items
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
      'ID,Vendor,ItemCode,Description,Qty,UoM,Comments,Price,PRDetailID',
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

      const vendorGroups = _.groupBy(prItems, (v: any) =>
        (v.Vendor || 'Unknown Vendor').trim()
      );

      groups.push({
        key: `pr-${pr.ID}`,
        name: pr.PRNumber,
        startIndex,
        count: prItems.length,
        level: 0,
        isCollapsed: true
      });

      Object.keys(vendorGroups).forEach(vendor => {
        const items = vendorGroups[vendor];

        // ✅ Sort by price (LOWEST first)
        const sortedItems = [...items].sort(
          (a, b) => Number(a.Price) - Number(b.Price)
        );

        groups.push({
          key: `pr-${pr.ID}-vendor-${vendor}`,
          name: vendor,
          startIndex,
          count: sortedItems.length,
          level: 1,
          isCollapsed: true
        });

        sortedItems.forEach((item, index) => {
          combinedItems.push({
            ...item,
            priceRank: index // 0 = lowest, 1 = second, 2+ = third
          });
          startIndex++;
        });
      });
    }

    this.setState({ listItems: combinedItems, groups });
  }

  /**
   * Columns
   */
  private getColumns(): IColumn[] {
    return [
      {
        key: 'itemcode',
        name: 'Item Code',
        fieldName: 'ItemCode',
        minWidth: 100
      },
      {
        key: 'description',
        name: 'Description',
        fieldName: 'Description',
        minWidth: 250
      },
      { key: 'qty', name: 'Qty', fieldName: 'Qty', minWidth: 60 },
      { key: 'uom', name: 'UoM', fieldName: 'UoM', minWidth: 60 },
      {
        key: 'comments',
        name: 'Comments',
        fieldName: 'Comments',
        minWidth: 200
      },
      {
        key: 'price',
        name: 'Price',
        fieldName: 'Price',
        minWidth: 100,
        onRender: (item: any) => {
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
      { key: 'vendor', name: 'Vendor', fieldName: 'Vendor', minWidth: 200 }
    ];
  }

  public render(): React.ReactElement<IQatarCementDashboardProps> {
    return (
      <section className={styles.container}>
        <DetailsList
          items={this.state.listItems || []}
          columns={this.getColumns()}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          groups={this.state.groups}
          groupProps={{
            showEmptyGroups: false,
            isAllGroupsCollapsed: true
          }}
        />

        <ModalOverlay
          isModalOpen={this.state.modalOverlay.isOpen}
          modalText={this.state.modalOverlay.Text}
        />
      </section>
    );
  }
}
