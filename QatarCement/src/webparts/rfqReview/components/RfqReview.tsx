import * as React from 'react';
import styles from './RfqReview.module.scss';
import type { IRfqReviewProps, IRfqReviewState } from '../interfaces/IRfqReviewProps';
import { RfqReviewService } from '../services/RfqReviewService';
import ModalOverlay from '../../../shared/controls/Overlay/Overlay';

export default class RfqReview extends React.Component<IRfqReviewProps, IRfqReviewState, {}> {
  private service: RfqReviewService;
  constructor(props: IRfqReviewProps) {
    super(props);
    this.state = {
      modalOverlay: {
        isOpen: false,
        Text: ''
      }
    };
    this.service = new RfqReviewService(this.props.context, this.props.context.pageContext.web.absoluteUrl);
  }
  public async componentDidMount(): Promise<void> {
    this.setState({ modalOverlay: { isOpen: true, Text: 'Loading...' } });
    const user = await this.service.getCurrentUser();
    console.log(user);
    this.setState({ modalOverlay: { isOpen: false, Text: '' } });
  }
  public render(): React.ReactElement<IRfqReviewProps> {
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
