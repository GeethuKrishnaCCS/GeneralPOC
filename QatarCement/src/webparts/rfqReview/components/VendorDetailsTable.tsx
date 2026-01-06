import * as React from 'react';
import { Dropdown, PrimaryButton, TextField, TooltipHost, Icon, Link, IconButton } from '@fluentui/react';
import styles from './RfqReview.module.scss';

interface IItemDetails {
    ItemCode: string;
    Description: string;
    Quantity: string;
    UOM: string;
    vendors: string;
    Id: number;
    WorkflowDetailsId?: number | null;
}

interface IVendorResponse {
    status?: string;
    price?: string;
    comments?: string;
    termsAndConditions?: string;
    technicalSupport?: string;
    warrantySupport?: string;
    attachments?: File[];
}

interface IAttachment {
    name: string;
    url: string;
    size?: number;
}

interface IVendorDetailsTableProps {
    itemDetails: IItemDetails[];
    currentUserEmail: string;
    vendorResponses: Record<number, IVendorResponse>;
    onResponseChange: (itemId: number, field: 'status' | 'price' | 'comments' | 'termsAndConditions' | 'technicalSupport' | 'warrantySupport', value: string) => void;
    onFileChange: (itemId: number, files: File[]) => void;
    onRemoveFile: (itemId: number, fileIndex: number) => void;
    onSubmitVendor: () => void;
    onCancel: () => void;
    attachments?: IAttachment[];
    isLoadingAttachments?: boolean;
}

const VendorDetailsTable: React.FC<IVendorDetailsTableProps> = ({ 
    itemDetails, 
    currentUserEmail, 
    vendorResponses,
    onResponseChange,
    onFileChange,
    onRemoveFile,
    onSubmitVendor, 
    onCancel,
    attachments = [],
    isLoadingAttachments = false
}) => {
    const userEmailLower = currentUserEmail.toLowerCase().trim();

    const filteredItems = itemDetails.filter(item => {
        if (!item.vendors) return false;
        const vendorList = item.vendors.toLowerCase();
        const emails = vendorList.split(/[,;]/).map(e => e.trim());
        return emails.some(email => email === userEmailLower || email.includes(userEmailLower));
    });

    const handleStatusChange = (event: React.FormEvent<HTMLDivElement>, option: any, itemId: number) => {
        onResponseChange(itemId, 'status', option.key);
    };

    const handlePriceChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string, itemId: number) => {
        onResponseChange(itemId, 'price', newValue || '');
    };

    const handleCommentsChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string, itemId: number) => {
        onResponseChange(itemId, 'comments', newValue || '');
    };

    const handleTextFieldChange = (field: 'termsAndConditions' | 'technicalSupport' | 'warrantySupport', newValue: string) => {
        onResponseChange(0, field, newValue || '');
    };

    const handleFileSelect = (itemId: number, event: React.ChangeEvent<HTMLInputElement>) => {
        const files = event.target.files;
        if (files && files.length > 0) {
            const fileArray = Array.from(files);
            const existingFiles = vendorResponses[itemId]?.attachments || [];
            onFileChange(itemId, [...existingFiles, ...fileArray]);
        }
        // Reset input to allow selecting the same file again
        event.target.value = '';
    };

    const calculateTotal = (): number => {
        return filteredItems.reduce((total, item) => {
            const price = parseFloat(vendorResponses[item.Id]?.price || '0');
            return total + (isNaN(price) ? 0 : price);
        }, 0);
    };

    const dropdownStatusOptions = [
        { key: 'Available', text: 'Available' },
        { key: 'Not Available', text: 'Not Available' },
    ];

    const formatFileSize = (bytes: number | undefined): string => {
        if (!bytes || bytes === 0) return '';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return '(' + Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i] + ')';
    };

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
        <>
            {/* Attachments Section */}
            {attachments && attachments.length > 0 && (
                <div className={styles.row}>
                    <div className={styles.col12}>
                        <div style={{ 
                            marginBottom: '20px', 
                            border: '1px solid #ddd', 
                            borderRadius: '4px',
                            padding: '15px',
                            backgroundColor: '#f9f9f9'
                        }}>
                            <div style={{ 
                                fontWeight: 600, 
                                marginBottom: '10px',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '8px',
                                fontSize: '16px'
                            }}>
                                <Icon iconName="Attach" />
                                <span>RFQ Attachments ({attachments.length})</span>
                            </div>
                            
                            {isLoadingAttachments ? (
                                <div style={{ padding: '10px', textAlign: 'center' }}>
                                    <span>Loading attachments...</span>
                                </div>
                            ) : (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                    {attachments.map((file, index) => (
                                        <div 
                                            key={index} 
                                            style={{ 
                                                display: 'flex', 
                                                alignItems: 'center',
                                                padding: '10px',
                                                backgroundColor: '#fff',
                                                borderRadius: '3px',
                                                border: '1px solid #e0e0e0'
                                            }}
                                        >
                                            <Icon 
                                                iconName="Page" 
                                                style={{ 
                                                    color: '#0078d4', 
                                                    marginRight: '10px',
                                                    fontSize: '20px'
                                                }} 
                                            />
                                            <Link 
                                                href={file.url} 
                                                target="_blank"
                                                style={{ 
                                                    flex: 1,
                                                    textDecoration: 'none'
                                                }}
                                            >
                                                {file.name} {formatFileSize(file.size)}
                                            </Link>
                                        </div>
                                    ))}
                                </div>
                            )}
                        </div>
                    </div>
                </div>
            )}

            {/* Three Multiline Text Fields */}
            <div className={styles.row}>
                <div className={styles.col12}>
                    <TextField
                        label="Terms and Conditions"
                        multiline
                        rows={4}
                        placeholder="Enter terms and conditions..."
                        value={vendorResponses[0]?.termsAndConditions || ''}
                        onChange={(e, newValue) => handleTextFieldChange('termsAndConditions', newValue!)}
                        styles={{ root: { marginBottom: '15px' } }}
                    />
                </div>
            </div>

            <div className={styles.row}>
                <div className={styles.col12}>
                    <TextField
                        label="Technical Support"
                        multiline
                        rows={4}
                        placeholder="Enter technical support details..."
                        value={vendorResponses[0]?.technicalSupport || ''}
                        onChange={(e, newValue) => handleTextFieldChange('technicalSupport', newValue!)}
                        styles={{ root: { marginBottom: '15px' } }}
                    />
                </div>
            </div>

            <div className={styles.row}>
                <div className={styles.col12}>
                    <TextField
                        label="Warranty Support"
                        multiline
                        rows={4}
                        placeholder="Enter warranty support details..."
                        value={vendorResponses[0]?.warrantySupport || ''}
                        onChange={(e, newValue) => handleTextFieldChange('warrantySupport', newValue!)}
                        styles={{ root: { marginBottom: '20px' } }}
                    />
                </div>
            </div>

            {/* Items Table */}
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
                                    <th className={styles.th}>Status</th>
                                    <th className={styles.th}>Price</th>
                                    <th className={styles.th}>Comments</th>
                                    <th className={styles.th}>Attachments</th>
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
                                        <td className={styles.th}>
                                            <Dropdown
                                                selectedKey={vendorResponses[item.Id]?.status}
                                                onChange={(e, option) => handleStatusChange(e, option!, item.Id)}
                                                placeholder="Select an option"
                                                options={dropdownStatusOptions}
                                            />
                                        </td>
                                        <td className={styles.th}>
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
                                        <td className={styles.th}>
                                            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                                <label 
                                                    htmlFor={`file-upload-${item.Id}`}
                                                    style={{
                                                        display: 'inline-block',
                                                        padding: '6px 12px',
                                                        backgroundColor: '#0078d4',
                                                        color: 'white',
                                                        borderRadius: '2px',
                                                        cursor: 'pointer',
                                                        textAlign: 'center',
                                                        fontSize: '14px'
                                                    }}
                                                >
                                                    <Icon iconName="Attach" style={{ marginRight: '5px' }} />
                                                    Choose Files
                                                </label>
                                                <input
                                                    id={`file-upload-${item.Id}`}
                                                    type="file"
                                                    multiple
                                                    onChange={(e) => handleFileSelect(item.Id, e)}
                                                    style={{ display: 'none' }}
                                                />
                                                
                                                {vendorResponses[item.Id]?.attachments && vendorResponses[item.Id]!.attachments!.length > 0 && (
                                                    <div style={{ marginTop: '8px' }}>
                                                        {vendorResponses[item.Id]!.attachments!.map((file, fileIndex) => (
                                                            <div 
                                                                key={fileIndex}
                                                                style={{
                                                                    display: 'flex',
                                                                    alignItems: 'center',
                                                                    justifyContent: 'space-between',
                                                                    padding: '4px 8px',
                                                                    backgroundColor: '#f3f2f1',
                                                                    borderRadius: '2px',
                                                                    marginBottom: '4px',
                                                                    fontSize: '12px'
                                                                }}
                                                            >
                                                                <span style={{ 
                                                                    overflow: 'hidden', 
                                                                    textOverflow: 'ellipsis',
                                                                    whiteSpace: 'nowrap',
                                                                    flex: 1
                                                                }}>
                                                                    <Icon iconName="Page" style={{ marginRight: '4px', fontSize: '12px' }} />
                                                                    {file.name}
                                                                </span>
                                                                <IconButton
                                                                    iconProps={{ iconName: 'Cancel' }}
                                                                    title="Remove file"
                                                                    ariaLabel="Remove file"
                                                                    onClick={() => onRemoveFile(item.Id, fileIndex)}
                                                                    styles={{
                                                                        root: {
                                                                            height: '20px',
                                                                            width: '20px'
                                                                        }
                                                                    }}
                                                                />
                                                            </div>
                                                        ))}
                                                    </div>
                                                )}
                                            </div>
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
                                    <td className={styles.th} colSpan={2}></td>
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
        </>
    );
};

export default VendorDetailsTable;