import * as React from 'react';
import styles from './RfqReview.module.scss';
import type { IRfqReviewProps, IRfqReviewState } from '../interfaces/IRfqReviewProps';
import { RfqReviewService } from '../services/RfqReviewService';
import ModalOverlay from '../../../shared/controls/Overlay/Overlay';
import { TextField } from '@fluentui/react';

export default class RfqReview extends React.Component<IRfqReviewProps, IRfqReviewState, {}> {
  private service: RfqReviewService;
  constructor(props: IRfqReviewProps) {
    super(props);
    this.state = {
      modalOverlay: {
        isOpen: false,
        Text: ''
      },
      prNumber: '',
      department: '',
      priority: '',
      dueDate: '',
      prInitiator: '',
      businessJustification: ''
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
                <TextField label="Due Date" value={this.state.dueDate} readOnly styles={textFieldStyles} />
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
                <h3 >Supporting Documents</h3>
              </div>
            </div>

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
