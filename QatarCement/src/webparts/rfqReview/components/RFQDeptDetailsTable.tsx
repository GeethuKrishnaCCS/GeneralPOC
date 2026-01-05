import * as React from 'react';
import { Dropdown, PrimaryButton, TextField, TooltipHost, Icon, IconButton } from '@fluentui/react';
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
    onSubmitRFQDept: () => void;
    onCancel: () => void;
    selectedFiles: File[];
    onFileSelect: (event: React.ChangeEvent<HTMLInputElement>) => void;
    onRemoveFile: (index: number) => void;
}

const RFQDeptDetailsTable: React.FC<IRFQDeptDetailsTableProps> = ({ 
    itemDetails, 
    vendorOptions, 
    handleChange, 
    onSubmitRFQDept, 
    onCancel,
    selectedFiles,
    onFileSelect,
    onRemoveFile
}) => {
    const fileInputRef = React.useRef<HTMLInputElement>(null);

    const handleFileButtonClick = () => {
        fileInputRef.current?.click();
    };

    const formatFileSize = (bytes: number): string => {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
    };

    return (
        <>
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

            {/* File Upload Section */}
            <div className={styles.row}>
                <div className={styles.col12}>
                    <div style={{ marginTop: '20px', marginBottom: '20px' }}>
                        <h4 style={{ marginBottom: '10px' }}>Attachments</h4>
                        
                        <input
                            ref={fileInputRef}
                            type="file"
                            multiple
                            onChange={onFileSelect}
                            style={{ display: 'none' }}
                            accept="*/*"
                        />
                        
                        <PrimaryButton
                            text="Select Files"
                            iconProps={{ iconName: 'Attach' }}
                            onClick={handleFileButtonClick}
                            style={{ marginBottom: '10px' }}
                        />

                        {selectedFiles.length > 0 && (
                            <div style={{ 
                                marginTop: '15px', 
                                border: '1px solid #ddd', 
                                borderRadius: '4px',
                                padding: '10px'
                            }}>
                                <div style={{ 
                                    fontWeight: 600, 
                                    marginBottom: '10px',
                                    display: 'flex',
                                    alignItems: 'center',
                                    gap: '8px'
                                }}>
                                    <Icon iconName="Attach" />
                                    <span>Selected Files ({selectedFiles.length})</span>
                                </div>
                                
                                {selectedFiles.map((file, index) => (
                                    <div 
                                        key={index} 
                                        style={{ 
                                            display: 'flex', 
                                            justifyContent: 'space-between',
                                            alignItems: 'center',
                                            padding: '8px',
                                            backgroundColor: index % 2 === 0 ? '#f9f9f9' : '#fff',
                                            borderRadius: '3px',
                                            marginBottom: '5px'
                                        }}
                                    >
                                        <div style={{ 
                                            display: 'flex', 
                                            alignItems: 'center',
                                            gap: '8px',
                                            flex: 1
                                        }}>
                                            <Icon iconName="Page" style={{ color: '#0078d4' }} />
                                            <span style={{ 
                                                overflow: 'hidden',
                                                textOverflow: 'ellipsis',
                                                whiteSpace: 'nowrap'
                                            }}>
                                                {file.name}
                                            </span>
                                            <span style={{ 
                                                color: '#666',
                                                fontSize: '12px',
                                                whiteSpace: 'nowrap'
                                            }}>
                                                ({formatFileSize(file.size)})
                                            </span>
                                        </div>
                                        <IconButton
                                            iconProps={{ iconName: 'Delete' }}
                                            title="Remove file"
                                            ariaLabel="Remove file"
                                            onClick={() => onRemoveFile(index)}
                                            styles={{
                                                root: { 
                                                    color: '#a4262c',
                                                    marginLeft: '10px'
                                                },
                                                rootHovered: { 
                                                    color: '#750b1c',
                                                    backgroundColor: '#fef0f1'
                                                }
                                            }}
                                        />
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                </div>
            </div>

            <div className={styles.row}>
                <div className={styles.col12}>
                    <div className={styles.rgtalign}>
                        <PrimaryButton className={styles.btn} onClick={onSubmitRFQDept}>
                            Submit
                        </PrimaryButton>
                        <PrimaryButton className={styles.btn} onClick={onCancel}>
                            Close
                        </PrimaryButton>
                    </div>
                </div>
            </div>
        </>
    );
};

export default RFQDeptDetailsTable;