import * as React from 'react';
import { TextField, TooltipHost } from '@fluentui/react';
import styles from './RfqReview.module.scss';

interface IItemDetails {
    ItemCode: string;
    Description: string;
    Quantity: string;
    UOM: string;
    vendors: string;
}

interface IVendorDetailsTableProps {
    itemDetails: IItemDetails[];
}

const VendorDetailsTable: React.FC<IVendorDetailsTableProps> = ({ itemDetails }) => {
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
                                            <TextField value={item.vendors} readOnly />
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

export default VendorDetailsTable;