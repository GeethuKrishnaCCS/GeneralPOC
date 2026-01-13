import * as React from 'react';
import styles from './ProcurementManagerDetailsTable.module.scss';
import { Dropdown, IDropdownOption, PrimaryButton, DefaultButton, TextField, Icon, Link, Dialog, DialogFooter, DialogType } from '@fluentui/react';

interface IProcurementManagerDetailsTableProps {
    itemDetails: any[];
    masterid: string;
    taskID: string | null;
    procurementManagerResponses: Record<number, { status?: string; comments?: string }>;
    commonProcurementStatus: string;
    commonProcurementComments: string;
    workflowDetailsAttachments: Record<number, Array<{ name: string; url: string }>>;
    vendorTermsAndConditions: string;
    vendorTechnicalSupport: string;
    vendorWarrantySupport: string;
    onResponseChange: (itemId: number, field: 'status' | 'comments', value: string) => void;
    onCommonStatusChange: (value: string) => void;
    onCommonCommentsChange: (value: string) => void;
    onSubmitProcurementManager: () => void;
    onCancel: () => void;
}

const ProcurementManagerDetailsTable: React.FC<IProcurementManagerDetailsTableProps> = ({
    itemDetails,
    procurementManagerResponses,
    commonProcurementStatus,
    commonProcurementComments,
    workflowDetailsAttachments,
    vendorTermsAndConditions,
    vendorTechnicalSupport,
    vendorWarrantySupport,
    onResponseChange,
    onCommonStatusChange,
    onCommonCommentsChange,
    onSubmitProcurementManager,
    onCancel
}) => {

    // State for dialog
    const [isDialogOpen, setIsDialogOpen] = React.useState(false);

    const statusOptions: IDropdownOption[] = [
        { key: 'Approve', text: 'Approve' },
        { key: 'Reject', text: 'Reject' },
    ];

    // Dialog configuration
    const dialogContentProps = {
        type: DialogType.normal,
        // title: 'Vendor Details',
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
        <div>
            <div className={styles.row}>
                <div className={styles.col12}>
                    <div className={styles.doctable}>
                        <table className={styles.table}>
                            <thead>
                                <tr className={styles.tr}>
                                    <th className={styles.th}>S.No</th>
                                    <th className={styles.th}>Item Code</th>
                                    <th className={styles.th}>Description</th>
                                    <th className={styles.th}>Quantity</th>
                                    <th className={styles.th}>UOM</th>
                                    <th className={styles.th}>Vendor</th>
                                    <th className={styles.th}>Vendor Status</th>
                                    <th className={styles.th}>Vendor Comments</th>
                                    <th className={styles.th}>Vendor Action</th>
                                    <th className={styles.th}>Attachments</th>
                                    <th className={styles.th}>Initiator Status</th>
                                    <th className={styles.th}>Initiator Comments</th>
                                    <th className={styles.th}>Maintenance Manager Status</th>
                                    <th className={styles.th}>Maintenance Manager Comments</th>
                                    <th className={styles.th}>Procurement Manager Comments</th>
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
                                            <td className={styles.th}>{item.InitiatorStatus || '-'}</td>
                                            <td className={styles.th}>
                                                <div style={{
                                                    maxWidth: '200px',
                                                    wordWrap: 'break-word',
                                                    whiteSpace: 'pre-wrap'
                                                }}>
                                                    {item.InitiatorComments || '-'}
                                                </div>
                                            </td>
                                            <td className={styles.th}>{item.ManagerStatus || '-'}</td>
                                            <td className={styles.th}>
                                                <div style={{
                                                    maxWidth: '200px',
                                                    wordWrap: 'break-word',
                                                    whiteSpace: 'pre-wrap'
                                                }}>
                                                    {item.ManagerComments || '-'}
                                                </div>
                                            </td>
                                            <td className={styles.th}>
                                                <TextField
                                                    placeholder="Enter comments"
                                                    value={procurementManagerResponses[item.WorkflowDetailsId]?.comments || ''}
                                                    onChange={(e, newValue) => {
                                                        onResponseChange(item.WorkflowDetailsId, 'comments', newValue || '');
                                                    }}
                                                    multiline
                                                    rows={2}
                                                    styles={{ root: { minWidth: '200px' } }}
                                                />
                                            </td>
                                        </tr>
                                    ))
                                ) : (
                                    <tr>
                                        <td colSpan={15} style={{ textAlign: 'center', padding: '20px' }}>
                                            No items found
                                        </td>
                                    </tr>
                                )}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            {/* Common Procurement Manager Status and Comments Section */}
            <div className={styles.row} style={{ marginTop: '20px', marginBottom: '20px' }}>
                <div className={styles.col6}>
                    <Dropdown
                        label="Procurement Manager Status (Common for All Items)"
                        placeholder="Select Status"
                        options={statusOptions}
                        selectedKey={commonProcurementStatus}
                        onChange={(e, option) => {
                            if (option) {
                                onCommonStatusChange(option.key as string);
                            }
                        }}
                        required
                    />
                </div>
                <div className={styles.col6}>
                    <TextField
                        label="Procurement Manager Comments (Common)"
                        placeholder="Enter common comments for all items"
                        value={commonProcurementComments}
                        onChange={(e, newValue) => {
                            onCommonCommentsChange(newValue || '');
                        }}
                        multiline
                        rows={4}
                        autoAdjustHeight
                    />
                </div>
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

            <div className={styles.row}>
                <div className={styles.col12}>
                    <div className={styles.rgtalign}>
                        <PrimaryButton className={styles.btn} onClick={onSubmitProcurementManager}>
                            Submit Procurement Manager Review
                        </PrimaryButton>
                        <DefaultButton className={styles.btn} onClick={onCancel}>
                            Close
                        </DefaultButton>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default ProcurementManagerDetailsTable;