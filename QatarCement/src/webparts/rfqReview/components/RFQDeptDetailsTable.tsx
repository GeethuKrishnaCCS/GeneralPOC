import * as React from 'react';
import { Dropdown, TextField, TooltipHost } from '@fluentui/react';
import styles from './RfqReview.module.scss';

interface IItemDetails {
    ItemCode: string;
    Description: string;
    Quantity: string;
    UOM: string;
    vendors: string;
}

interface IRFQDeptDetailsTableProps {
    itemDetails: IItemDetails[];
    vendorOptions: { key: any; text: string }[];
    handleChange: (index: number, field: keyof IItemDetails, value: string) => void;
}

const RFQDeptDetailsTable: React.FC<IRFQDeptDetailsTableProps> = ({ itemDetails, vendorOptions, handleChange }) => {
    return (
        <div className={styles.row}>
            <div className={styles.col12}>
                {itemDetails.length > 0 && (
                    <div className={styles.doctable}>
                        <table className={styles.table}>
                            <thead>
                                <tr className={styles.tr}>
                                    <th className={styles.th}>Sl No</th>
                                    <th className={styles.th}>ItemCode</th>
                                    <th className={styles.th}>Description</th>
                                    <th className={styles.th}>Quantity</th>
                                    <th className={styles.th}>UOM</th>
                                    <th className={styles.th}>Vendors</th>
                                </tr>
                            </thead>
                            <tbody>
                                {itemDetails.map((item, key) => (
                                    <tr key={key} className={styles.tr}>
                                        <td className={styles.th}>{key + 1}</td>
                                        <td className={styles.th}>
                                            <TextField value={item.ItemCode} readOnly />
                                        </td>
                                        <td className={styles.th}>
                                            <TooltipHost content={item.Description}>
                                                <TextField value={item.Description} readOnly />
                                            </TooltipHost>
                                        </td>
                                        <td className={styles.th}>
                                            <TextField value={item.Quantity} readOnly />
                                        </td>
                                        <td className={styles.th}>
                                            <TextField value={item.UOM} readOnly />
                                        </td>
                                        <td className={styles.th}>
                                            <div className={styles.vendorCell}>
                                                <Dropdown
                                                    placeholder="Select Vendors"
                                                    multiSelect
                                                    options={vendorOptions}
                                                    selectedKeys={item.vendors ? item.vendors.split(',') : []}
                                                    onChange={(e, option) => {
                                                        let updated = [...(item.vendors ? item.vendors.split(',') : [])];

                                                        if (option?.selected) {
                                                            updated.push(option.key as string);
                                                        } else {
                                                            updated = updated.filter((v) => v !== option?.key);
                                                        }

                                                        handleChange(key, 'vendors', updated.join(','));
                                                    }}
                                                />
                                            </div>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}
            </div>
        </div>
    );
};

export default RFQDeptDetailsTable;