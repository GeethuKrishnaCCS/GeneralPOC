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
      'ID,Vendor,ItemCode,Description,Qty,UoM,Comments,Price,PRDetailID',
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
          isCollapsed: true
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
        isCollapsed: true,
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
          </div>
        </div>

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
