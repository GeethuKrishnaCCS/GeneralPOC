import * as React from 'react';
import styles from './InitiatorDetailsTable.module.scss';


interface IInitiatorDetailsTableFormProps {
    itemDetails: any[];
    masterid: string;
    taskID: string | null;
    // initiatorResponses: Record<number, { status?: string; comments?: string }>;
    // onResponseChange: (itemId: number, field: 'status' | 'comments', value: string) => void;
    // onSubmitInitiator: () => void;
    // onCancel: () => void;
}

const InitiatorDetailsTableForm: React.FC<IInitiatorDetailsTableFormProps> = ({
    itemDetails,
    // initiatorResponses,
    // onResponseChange,
    // onSubmitInitiator,
    // onCancel
}) => {

    // const statusOptions: IDropdownOption[] = [
    //     { key: 'Technically Accepted', text: 'Technically Accepted' },
    //     { key: 'Technically Not Accepted', text: 'Technically Not Accepted' },
    // ];

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
                                {/* <th className={styles.th}>Technical Action</th>
                                <th className={styles.th}>Technical Comments</th> */}
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
                                        <td className={styles.th}>{item.Comments || '-'}</td>
                                        {/* <td className={styles.th}>
                                            <Dropdown
                                                placeholder="Select Status"
                                                options={statusOptions}
                                                selectedKey={initiatorResponses[item.WorkflowDetailsId]?.status || ''}
                                                onChange={(e, option) => {
                                                    if (option) {
                                                        onResponseChange(item.WorkflowDetailsId, 'status', option.key as string);
                                                    }
                                                }}
                                                styles={{ dropdown: { width: 150 } }}
                                            />
                                        </td>
                                        <td className={styles.th}>
                                            <TextField
                                                placeholder="Enter comments"
                                                value={initiatorResponses[item.WorkflowDetailsId]?.comments || ''}
                                                onChange={(e, newValue) => {
                                                    onResponseChange(item.WorkflowDetailsId, 'comments', newValue || '');
                                                }}
                                                multiline
                                                rows={2}
                                            />
                                        </td> */}
                                    </tr>
                                ))
                            ) : (
                                <tr>
                                    <td colSpan={10} style={{ textAlign: 'center' }}>
                                        No items found
                                    </td>
                                </tr>
                            )}
                        </tbody>
                    </table>
                </div>


                {/* <div className={styles.row}>
                    <div className={styles.col12}>
                        <div className={styles.rgtalign}>
                            <PrimaryButton className={styles.btn} onClick={onSubmitInitiator}>
                                Submit Technical Review
                            </PrimaryButton>
                            <DefaultButton className={styles.btn} onClick={onCancel}>                                Close
                            </DefaultButton>
                        </div>
                    </div>
                </div> */}
            </div>
        </div>
    );
};

export default InitiatorDetailsTableForm;