import * as React from 'react';
import styles from './QatarCementDashboard.module.scss';
import type { IQatarCementDashboardProps, IQatarCementDashboardState } from '../interfaces/IQatarCementDashboardProps';
import { QatarCementDashboardService } from '../services/QatarCementDashboardService';
import ModalOverlay from '../../../shared/controls/Overlay/Overlay';

export default class QatarCementDashboard extends React.Component<IQatarCementDashboardProps, IQatarCementDashboardState, {}> {
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
    this.setState({ modalOverlay: { isOpen: false, Text: '' } });
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
        <ModalOverlay
          isModalOpen={this.state.modalOverlay.isOpen}
          modalText={this.state.modalOverlay.Text}
        />
      </section>
    );
  }
}
