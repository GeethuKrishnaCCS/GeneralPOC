import * as React from 'react';
import styles from './InitiatorDetailsTable.module.scss';
import { 
    DefaultButton, 
    Icon, 
    Link,
    Dialog,
    DialogType,
    DialogFooter
} from '@fluentui/react';

interface IInitiatorDetailsTableProps {
    itemDetails: any[];
    masterid: string;
    taskID: string | null;
    
    workflowDetailsAttachments: Record<number, Array<{ name: string; url: string }>>;
    vendorTermsAndConditions: string;
    vendorTechnicalSupport: string;
    vendorWarrantySupport: string;

}

const InitiatorDetailsTableForm: React.FC<IInitiatorDetailsTableProps> = ({
    itemDetails,

    workflowDetailsAttachments,
    vendorTermsAndConditions,
    vendorTechnicalSupport,
    vendorWarrantySupport,
   
}) => {

    // State for dialog
    const [isDialogOpen, setIsDialogOpen] = React.useState(false);

    // Dialog configuration
    const dialogContentProps = {
        type: DialogType.normal,
        title: 'Vendor Details',
        closeButtonAriaLabel: 'Close',
    };

    // Helper function to render attachments
    const renderAttachments = (workflowDetailsId: number) => {
        const attachments = workflowDetailsAttachments[workflowDetailsId] || [];

        if (attachments.length === 0) {
            return <span style={{ color: '#999' }}>No attachments</span>;
        }

        return (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                {attachments.map((attachment, idx) => (
                    <Link
                        key={idx}
                        href={attachment.url}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{
                            display: 'flex',
                            alignItems: 'center',
                            fontSize: '12px',
                            textDecoration: 'none'
                        }}
                    >
                        <Icon
                            iconName="Attach"
                            style={{
                                marginRight: '5px',
                                fontSize: '14px',
                                color: '#0078d4'
                            }}
                        />
                        <span style={{
                            color: '#0078d4',
                            textDecoration: 'underline',
                            wordBreak: 'break-word',
                            maxWidth: '150px'
                        }}>
                            {attachment.name}
                        </span>
                    </Link>
                ))}
            </div>
        );
    };

    return (
        <div className={styles.row}>
            <div className={styles.col12}>
                <div className={styles.doctable}>
                    <table className={styles.table}>
                        <thead>
                            <tr className={styles.tr}>
                                <th className={styles.th}>S.No</th>
                                <th className={styles.th}>Item Code</th>
                                <th className={styles.th}>Description</th>
                                <th className={styles.th}>Qty</th>
                                <th className={styles.th}>UOM</th>
                                <th className={styles.th}>Vendor</th>
                                <th className={styles.th}>Vendor Status</th>
                                <th className={styles.th}>Vendor Comments</th>
                                <th className={styles.th}>Vendor Action</th>
                                <th className={styles.th}>Attachments</th>
                               
                            </tr>
                        </thead>
                        <tbody>
                            {itemDetails.length > 0 ? (
                                itemDetails.map((item, index) => (
                                    <tr key={`${item.WorkflowDetailsId}-${index}`} className={styles.tr}>
                                        <td className={styles.th}>{index + 1}</td>
                                        <td className={styles.th}>{item.ItemCode}</td>
                                        <td className={styles.th}>{item.Description}</td>
                                        <td className={styles.th}>{item.Quantity}</td>
                                        <td className={styles.th}>{item.UOM}</td>
                                        <td className={styles.th}>{item.Vendor}</td>
                                        <td className={styles.th}>{item.Status || '-'}</td>
                                        <td className={styles.th}>
                                            <div style={{
                                                maxWidth: '200px',
                                                wordWrap: 'break-word',
                                                whiteSpace: 'pre-wrap'
                                            }}>
                                                {item.Comments || '-'}
                                            </div>
                                        </td>

                                        <td className={styles.th}>
                                            <Icon 
                                                iconName="CommentSolid" 
                                                style={{ 
                                                    marginRight: '5px',
                                                    cursor: 'pointer',
                                                    color: '#0078d4',
                                                    fontSize: '16px'
                                                }}
                                                onClick={() => setIsDialogOpen(true)}
                                                title="View Vendor Details"
                                            />
                                        </td>

                                        <td className={styles.th}>
                                            {item.WorkflowDetailsId
                                                ? renderAttachments(item.WorkflowDetailsId)
                                                : <span style={{ color: '#999' }}>-</span>
                                            }
                                        </td>
                                        
                                    </tr>
                                ))
                            ) : (
                                <tr>
                                    <td colSpan={13} style={{ textAlign: 'center', padding: '20px' }}>
                                        No items found
                                    </td>
                                </tr>
                            )}
                        </tbody>
                    </table>
                </div>

                {/* Vendor Details Dialog */}
                <Dialog
                    hidden={!isDialogOpen}
                    onDismiss={() => setIsDialogOpen(false)}
                    dialogContentProps={dialogContentProps}
                    minWidth={500}
                    maxWidth={700}
                >
                    <div style={{ padding: '10px 0' }}>
                        <div style={{ marginBottom: '15px' }}>
                            <strong style={{ display: 'block', marginBottom: '5px', color: '#323130' }}>
                                Terms and Conditions:
                            </strong>
                            <div style={{ 
                                padding: '10px', 
                                backgroundColor: '#f3f2f1', 
                                borderRadius: '4px',
                                whiteSpace: 'pre-wrap',
                                wordWrap: 'break-word'
                            }}>
                                {vendorTermsAndConditions || 'Not provided'}
                            </div>
                        </div>

                        <div style={{ marginBottom: '15px' }}>
                            <strong style={{ display: 'block', marginBottom: '5px', color: '#323130' }}>
                                Technical Support:
                            </strong>
                            <div style={{ 
                                padding: '10px', 
                                backgroundColor: '#f3f2f1', 
                                borderRadius: '4px',
                                whiteSpace: 'pre-wrap',
                                wordWrap: 'break-word'
                            }}>
                                {vendorTechnicalSupport || 'Not provided'}
                            </div>
                        </div>

                        <div style={{ marginBottom: '15px' }}>
                            <strong style={{ display: 'block', marginBottom: '5px', color: '#323130' }}>
                                Warranty Support:
                            </strong>
                            <div style={{ 
                                padding: '10px', 
                                backgroundColor: '#f3f2f1', 
                                borderRadius: '4px',
                                whiteSpace: 'pre-wrap',
                                wordWrap: 'break-word'
                            }}>
                                {vendorWarrantySupport || 'Not provided'}
                            </div>
                        </div>
                    </div>

                    <DialogFooter>
                        <DefaultButton onClick={() => setIsDialogOpen(false)} text="Close" />
                    </DialogFooter>
                </Dialog>
            </div>
        </div>
    );
};


export default InitiatorDetailsTableForm;