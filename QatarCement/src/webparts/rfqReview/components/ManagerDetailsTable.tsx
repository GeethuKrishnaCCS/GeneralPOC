import * as React from 'react';
import styles from './ManagerDetailsTable.module.scss';
import { Dropdown, IDropdownOption, PrimaryButton, DefaultButton, TextField } from '@fluentui/react';

interface IManagerDetailsTableProps {
    itemDetails: any[];
    masterid: string;
    taskID: string | null;
    managerResponses: Record<number, { status?: string; comments?: string }>;
    commonManagerStatus: string;
    commonManagerComments: string;
    onResponseChange: (itemId: number, field: 'status' | 'comments', value: string) => void;
    onCommonStatusChange: (value: string) => void;
    onCommonCommentsChange: (value: string) => void;
    onSubmitManager: () => void;
    onCancel: () => void;
}

const ManagerDetailsTable: React.FC<IManagerDetailsTableProps> = ({
    itemDetails,
    managerResponses,
    commonManagerStatus,
    commonManagerComments,
    onResponseChange,
    onCommonStatusChange,
    onCommonCommentsChange,
    onSubmitManager,
    onCancel
}) => {

    const statusOptions: IDropdownOption[] = [
        { key: 'Approved', text: 'Approved' },
        { key: 'Rejected', text: 'Rejected' },
    ];

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
                                    <th className={styles.th}>Vendor Price</th>
                                    <th className={styles.th}>Vendor Comments</th>
                                    <th className={styles.th}>Initiator Status</th>
                                    <th className={styles.th}>Initiator Comments</th>
                                    <th className={styles.th}>Maintenance Manager Comments</th>
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
                                            <td className={styles.th}>{item.Price || '-'}</td>
                                            <td className={styles.th}>{item.Comments || '-'}</td>
                                            <td className={styles.th}>{item.InitiatorStatus || '-'}</td>
                                            <td className={styles.th}>{item.InitiatorComments || '-'}</td>
                                            <td className={styles.th}>
                                                <TextField
                                                    placeholder="Enter comments"
                                                    value={managerResponses[item.WorkflowDetailsId]?.comments || ''}
                                                    onChange={(e, newValue) => {
                                                        onResponseChange(item.WorkflowDetailsId, 'comments', newValue || '');
                                                    }}
                                                    multiline
                                                    rows={2}
                                                />
                                            </td>
                                        </tr>
                                    ))
                                ) : (
                                    <tr>
                                        <td colSpan={12} style={{ textAlign: 'center' }}>
                                            No items found
                                        </td>
                                    </tr>
                                )}
                            </tbody>
                        </table>
                    </div>


                </div>
            </div>

            {/* Common Maintenance Manager Status and Comments */}
            <div className={styles.row} style={{ marginTop: '20px' }}>
                <div className={styles.col6}>
                    <Dropdown
                        label="Maintenance Manager Status (Common for all items)"
                        placeholder="Select Status"
                        options={statusOptions}
                        selectedKey={commonManagerStatus}
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
                        label="Common Comments (Applied to all items)"
                        placeholder="Enter common comments for all items"
                        value={commonManagerComments}
                        onChange={(e, newValue) => {
                            onCommonCommentsChange(newValue || '');
                        }}
                        multiline
                        rows={4}
                    />
                </div>
            </div>

            <div className={styles.row}>
                <div className={styles.col12}>
                    <div className={styles.rgtalign}>
                        <PrimaryButton className={styles.btn} onClick={onSubmitManager}>
                            Submit Maintenance Manager Review
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

export default ManagerDetailsTable;