import * as React from 'react';
import { Dropdown, PrimaryButton, TextField, TooltipHost } from '@fluentui/react';
import styles from './RfqReview.module.scss';

interface IItemDetails {
    ItemCode: string;
    Description: string;
    Quantity: string;
    UOM: string;
    vendors: string;
    Id: number; 
}

interface IVendorResponse {
    status?: string;
    price?: string;
    comments?: string;
}

interface IVendorDetailsTableProps {
    itemDetails: IItemDetails[];
    currentUserEmail: string;
    vendorResponses: Record<number, IVendorResponse>;
    onResponseChange: (itemId: number, field: 'status' | 'price' | 'comments', value: string) => void; // ✓ Added 'price' field
    onSubmitVendor: () => void;
    onCancel: () => void;
}


const VendorDetailsTable: React.FC<IVendorDetailsTableProps> = ({ itemDetails, currentUserEmail, vendorResponses,
    onResponseChange, onSubmitVendor, onCancel }) => {
    const userEmailLower = currentUserEmail.toLowerCase().trim();

    const filteredItems = itemDetails.filter(item => {
        if (!item.vendors) return false;
        const vendorList = item.vendors.toLowerCase();
        const emails = vendorList.split(/[,;]/).map(e => e.trim());
        return emails.some(email => email === userEmailLower || email.includes(userEmailLower));
    });

    // ✓ Use item.Id instead of index for tracking
    const handleStatusChange = (event: React.FormEvent<HTMLDivElement>, option: any, itemId: number) => {
        onResponseChange(itemId, 'status', option.key);
    };

    const handlePriceChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string, itemId: number) => {
        onResponseChange(itemId, 'price', newValue || '');
    };

    const handleCommentsChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string, itemId: number) => {
        onResponseChange(itemId, 'comments', newValue || '');
    };

    // ✓ Calculate total price
    const calculateTotal = (): number => {
        return filteredItems.reduce((total, item) => {
            const price = parseFloat(vendorResponses[item.Id]?.price || '0');
            return total + (isNaN(price) ? 0 : price);
        }, 0);
    };

    const dropdownStatusOptions = [
        { key: 'Yes', text: 'Yes' },
        { key: 'No', text: 'No' },
    ];

    if (filteredItems.length === 0) {
        return (
            <div className={styles.row}>
                <div className={styles.col12}>
                    <p>No items are assigned to you in this RFQ.</p>
                </div>
            </div>
        );
    }


    return (
        <div className={styles.row}>
            <div className={styles.col12}>
                <div className={styles.doctable}>
                    <table className={styles.table}>
                        <thead>
                            <tr className={styles.tr}>
                                <th className={styles.th}>Sl No</th>
                                <th className={styles.th}>Item Code</th>
                                <th className={styles.th}>Description</th>
                                <th className={styles.th}>Quantity</th>
                                <th className={styles.th}>UOM</th>
                                {/* <th className={styles.th}>Assigned Vendors</th> */}
                                <th className={styles.th}>Status</th>
                                <th className={styles.th}>Price</th>
                                <th className={styles.th}>Comments</th>
                            </tr>
                        </thead>
                        <tbody>
                            {filteredItems.map((item, index) => (
                                <tr key={item.ItemCode + index} className={styles.tr}>
                                    <td className={styles.th}>{index + 1}</td>
                                    <td className={styles.th}>
                                        <TextField value={item.ItemCode || ''} readOnly />
                                    </td>
                                    <td className={styles.th}>
                                        <TooltipHost content={item.Description}>
                                            <TextField value={item.Description || ''} readOnly />
                                        </TooltipHost>
                                    </td>
                                    <td className={styles.th}>
                                        <TextField value={item.Quantity || ''} readOnly />
                                    </td>
                                    <td className={styles.th}>
                                        <TextField value={item.UOM || ''} readOnly />
                                    </td>
                                    {/* <td className={styles.th}>
                                        <TextField value={item.vendors || ''} readOnly />
                                    </td> */}
                                    <td className={styles.th}>
                                        <Dropdown
                                            selectedKey={vendorResponses[item.Id]?.status}
                                            onChange={(e, option) => handleStatusChange(e, option!, item.Id)}
                                            placeholder="Select an option"
                                            options={dropdownStatusOptions}
                                        />
                                    </td>
                                    <td className={styles.th}> {/* ✓ Added Price field */}
                                        <TextField
                                            placeholder='Enter price'
                                            value={vendorResponses[item.Id]?.price || ''}
                                            onChange={(e, newValue) => handlePriceChange(e, newValue!, item.Id)}
                                            type="number"
                                        />
                                    </td>
                                    <td className={styles.th}>
                                        <TextField
                                            placeholder='Comments'
                                            multiline
                                            rows={3}
                                            value={vendorResponses[item.Id]?.comments || ''}
                                            onChange={(e, newValue) => handleCommentsChange(e, newValue!, item.Id)}
                                        />
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                        <tfoot>
                            <tr className={styles.tr}>
                                <td className={styles.th} colSpan={6} style={{ textAlign: 'right', fontWeight: 'bold' }}>
                                    Total:
                                </td>
                                <td className={styles.th}>
                                    <TextField
                                        value={calculateTotal().toFixed(2)}
                                        readOnly
                                        styles={{
                                            field: {
                                                fontWeight: 'bold',
                                                backgroundColor: '#f3f3f3'
                                            }
                                        }}
                                    />
                                </td>
                                <td className={styles.th}></td>
                            </tr>
                        </tfoot>
                    </table>
                </div>



                <div className={styles.row}>
                    <div className={styles.col12}>
                        <div className={styles.rgtalign}>
                            <PrimaryButton className={styles.btn} onClick={onSubmitVendor}>Submit</PrimaryButton>
                            <PrimaryButton className={styles.btn} onClick={onCancel}>Close</PrimaryButton>
                        </div>
                    </div>
                </div>

            </div>
        </div>
    );
};

export default VendorDetailsTable;