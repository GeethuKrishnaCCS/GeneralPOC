import { BaseService } from "../../../shared/services/BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../shared/PnP/pnpjsConfig";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/items";
import "@pnp/sp/lists";

export class RfqReviewService extends BaseService {
  private spfi: SPFI;

  constructor(context: WebPartContext, siteUrl: string) {
    super(context, siteUrl);
    this.spfi = getSP(context);
  }

  public getCurrentUser() {
    return this.spfi.web.currentUser();
  }

  public async isUserInGroup(groupName: string): Promise<boolean> {
    try {
      const groups = await this.spfi.web.currentUser.groups();
      return groups.some(g => g.Title.toLowerCase() === groupName.toLowerCase());
    } catch (error) {
      console.error("Error checking group membership:", error);
      return false;
    }
  }

  /**
   * Get attachments from document library filtered by PRDetailID
   * @param libraryName - Internal name of the document library (e.g., "Documents" for "Shared Documents")
   * @param prDetailId - The PRDetailID to filter by
   * @returns Array of attachment objects with name, url, and size
   */
  public async getAttachmentsByPRDetailID(
    libraryName: string,
    prDetailId: string
  ): Promise<Array<{ name: string; url: string; size: number }>> {
    try {
      // For "Shared Documents", the internal name is usually "Documents"
      // You can also try using the display name directly
      const internalLibraryName = libraryName === "Shared Documents" ? "Documents" : libraryName;
      // const internalLibraryName = libraryName === "Shared Documents" ? "Documents" : libraryName;

      console.log(`Fetching attachments from library: ${internalLibraryName} for PRDetailID: ${prDetailId}`);

      // Query the document library with filter
      const items = await this.spfi.web.lists
        .getByTitle(internalLibraryName)
        .items
        .select("File/Name", "File/ServerRelativeUrl", "File/Length", "PRDetailID", "Id")
        .expand("File")
        .filter(`PRDetailID eq '${prDetailId}'`)();

      console.log("Fetched attachments:", items);

      // Map to attachment objects
      const attachments = items.map(item => ({
        name: item.File.Name,
        url: `${window.location.origin}${item.File.ServerRelativeUrl}`,
        size: item.File.Length
      }));

      return attachments;

    } catch (error) {
      console.error("Error fetching attachments:", error);

      // Try alternative approach if first attempt fails
      console.log("Attempting alternative method to fetch attachments...");
      try {
        return await this.getAttachmentsByPRDetailIDAlternative(libraryName, prDetailId);
      } catch (altError) {
        console.error("Alternative method also failed:", altError);
        return [];
      }
    }
  }

  /**
   * Alternative method to get attachments using folder approach
   * @param libraryName - Name of the document library
   * @param prDetailId - The PRDetailID to filter by
   * @returns Array of attachment objects
   */
  private async getAttachmentsByPRDetailIDAlternative(
    libraryName: string,
    prDetailId: string
  ): Promise<Array<{ name: string; url: string; size: number }>> {
    try {
      const internalLibraryName = libraryName === "Shared Documents" ? "Documents" : libraryName;

      // Get all files from the library
      const files = await this.spfi.web.lists
        .getByTitle(internalLibraryName)
        .items
        .select("FileLeafRef", "FileRef", "File_x0020_Size", "PRDetailID")
        .filter(`PRDetailID eq '${prDetailId}'`)();

      console.log("Files fetched using alternative method:", files);

      const attachments = files.map(file => ({
        name: file.FileLeafRef,
        url: `${window.location.origin}${file.FileRef}`,
        size: file.File_x0020_Size || 0
      }));

      return attachments;

    } catch (error) {
      console.error("Alternative method error:", error);
      throw error;
    }
  }

  /**
   * Upload a file to SharePoint document library
   * @param libraryPath - Server relative path of the document library
   * @param fileName - Name for the uploaded file
   * @param file - File object to upload
   * @param metadata - Optional metadata to set on the file (e.g., { PRDetailID: '123' })
   * @returns Server relative URL of the uploaded file
   */
  public async uploadFile(
    libraryPath: string,
    fileName: string,
    file: File,
    metadata?: Record<string, any>
  ): Promise<string> {
    try {
      // Validate file size (optional - set your own limit, e.g., 10MB)
      const maxSize = 10 * 1024 * 1024; // 10MB
      if (file.size > maxSize) {
        throw new Error(`File ${fileName} exceeds maximum size of 10MB`);
      }

      // Read file as array buffer
      const arrayBuffer = await this.readFileAsArrayBuffer(file);

      // Upload file to document library
      const result = await this.spfi.web
        .getFolderByServerRelativePath(libraryPath)
        .files
        .addUsingPath(fileName, arrayBuffer, { Overwrite: true });

      console.log(`File uploaded successfully: ${fileName}`, result);

      // If metadata is provided, update the list item properties
      if (metadata && Object.keys(metadata).length > 0) {
        const serverRelativeUrl = `${libraryPath}/${fileName}`;
        const file = await this.spfi.web.getFileByServerRelativePath(serverRelativeUrl);
        const item = await file.getItem();
        await item.update(metadata);
        console.log(`Metadata updated for file: ${fileName}`, metadata);
      }

      // Construct the server relative URL
      const serverRelativeUrl = `${libraryPath}/${fileName}`;

      return serverRelativeUrl;

    } catch (error) {
      console.error(`Error uploading file ${fileName}:`, error);

      // Provide more specific error messages
      if (error.message?.includes("does not exist")) {
        throw new Error(`Document library '${libraryPath}' does not exist. Please check the library name.`);
      } else if (error.message?.includes("access denied") || error.message?.includes("Access denied")) {
        throw new Error(`Access denied. You don't have permission to upload files to '${libraryPath}'.`);
      }

      throw error;
    }
  }

  /**
   * Upload multiple files to SharePoint document library
   * @param libraryPath - Server relative path of the document library
   * @param files - Array of files with their names
   * @returns Array of server relative URLs
   */
  public async uploadMultipleFiles(
    libraryPath: string,
    files: { name: string; file: File }[]
  ): Promise<string[]> {
    const uploadPromises = files.map(({ name, file }) =>
      this.uploadFile(libraryPath, name, file)
    );

    try {
      const urls = await Promise.all(uploadPromises);
      return urls;
    } catch (error) {
      console.error("Error uploading multiple files:", error);
      throw error;
    }
  }

  /**
   * Helper method to read file as ArrayBuffer
   */
  private readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * Create a folder in document library if it doesn't exist
   */
  public async ensureFolder(
    libraryPath: string,
    folderName: string
  ): Promise<void> {
    try {
      await this.spfi.web
        .getFolderByServerRelativePath(libraryPath)
        .folders
        .addUsingPath(folderName);
      console.log(`Folder created: ${folderName}`);
    } catch (error) {
      // Folder might already exist, check if it's a different error
      if (!error.message.includes("already exists")) {
        console.error(`Error creating folder ${folderName}:`, error);
        throw error;
      }
    }
  }

  /**
   * Alternative upload method using getFileByServerRelativePath
   * This method retrieves the file info after upload to get the exact URL
   */
  public async uploadFileWithInfo(
    libraryPath: string,
    fileName: string,
    file: File
  ): Promise<{ url: string; fileInfo: any }> {
    try {
      // Validate file size
      const maxSize = 10 * 1024 * 1024; // 10MB
      if (file.size > maxSize) {
        throw new Error(`File ${fileName} exceeds maximum size of 10MB`);
      }

      // Read file as array buffer
      const arrayBuffer = await this.readFileAsArrayBuffer(file);

      // Upload file
      await this.spfi.web
        .getFolderByServerRelativePath(libraryPath)
        .files
        .addUsingPath(fileName, arrayBuffer, { Overwrite: true });

      // Get the file info after upload
      const fileServerRelativeUrl = `${libraryPath}/${fileName}`;
      const fileInfo = await this.spfi.web
        .getFileByServerRelativePath(fileServerRelativeUrl)();

      console.log(`File uploaded and retrieved: ${fileName}`, fileInfo);

      return {
        url: fileInfo.ServerRelativeUrl,
        fileInfo: fileInfo
      };

    } catch (error) {
      console.error(`Error uploading file ${fileName}:`, error);
      throw error;
    }
  }


  // Add these methods to your existing RfqReviewService class

  /**
   * Upload files and attach them to a WorkflowDetails list item
   * @param libraryPath - Server relative path of the document library
   * @param workflowDetailsId - ID of the WorkflowDetails item to attach files to
   * @param files - Array of files to upload
   * @param prDetailId - PRDetailID for file naming
   * @param itemId - Item ID for file naming
   * @returns Array of attachment info with URLs
   */
  public async uploadAndAttachToWorkflowDetails(
    libraryPath: string,
    workflowDetailsId: number,
    files: File[],
    prDetailId: string,
    itemId: number
  ): Promise<Array<{ name: string; url: string }>> {
    try {
      const uploadedFiles: Array<{ name: string; url: string }> = [];

      // Upload each file
      for (const file of files) {
        const fileName = `WF${workflowDetailsId}_Item${itemId}_${file.name}`;

        // Read file as array buffer
        const arrayBuffer = await this.readFileAsArrayBuffer(file);

        // Upload file to document library
        await this.spfi.web
          .getFolderByServerRelativePath(libraryPath)
          .files
          .addUsingPath(fileName, arrayBuffer, { Overwrite: true });

        // Construct full URL
        const fileUrl = `${window.location.origin}${libraryPath}/${fileName}`;

        uploadedFiles.push({
          name: fileName,
          url: fileUrl
        });

        console.log(`File uploaded: ${fileName}`);
      }

      // Update WorkflowDetails item with attachment URLs
      if (uploadedFiles.length > 0) {
        await this.updateWorkflowDetailsAttachments(workflowDetailsId, uploadedFiles);
      }

      return uploadedFiles;

    } catch (error) {
      console.error("Error uploading files to WorkflowDetails:", error);
      throw error;
    }
  }

  /**
   * Update WorkflowDetails list item with attachment information
   * @param workflowDetailsId - ID of the WorkflowDetails item
   * @param attachments - Array of attachment info
   */
  public async updateWorkflowDetailsAttachments(
    workflowDetailsId: number,
    attachments: Array<{ name: string; url: string }>
  ): Promise<void> {
    try {
      // Format attachments as JSON string or concatenated URLs
      const attachmentUrls = attachments.map(att => att.url).join(';');

      // Update the item with attachment information
      await this.spfi.web.lists
        .getByTitle("WorkflowDetails")
        .items
        .getById(workflowDetailsId)
        .update({
          Attachments: attachmentUrls // Adjust field name if different
        });

      console.log(`WorkflowDetails item ${workflowDetailsId} updated with attachments`);

    } catch (error) {
      console.error(`Error updating WorkflowDetails item ${workflowDetailsId}:`, error);
      throw error;
    }
  }

  /**
  * Add attachments using SharePoint's native attachment functionality
  * This attaches files directly to the WorkflowDetails list item
  * @param workflowDetailsId - ID of the WorkflowDetails item
  * @param files - Array of files to attach
  */
  public async addAttachmentsToWorkflowDetails(
    workflowDetailsId: number,
    files: File[]
  ): Promise<void> {
    try {
      console.log(`Starting attachment process for WorkflowDetails ID: ${workflowDetailsId}`);
      console.log(`Number of files to attach: ${files.length}`);

      const item = this.spfi.web.lists
        .getByTitle("WorkflowDetails")
        .items
        .getById(workflowDetailsId);

      // Upload each file as attachment
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        console.log(`Uploading file ${i + 1}/${files.length}: ${file.name} (${file.size} bytes)`);

        const arrayBuffer = await this.readFileAsArrayBuffer(file);

        // Add as native SharePoint attachment
        await item.attachmentFiles.add(file.name, arrayBuffer);

        console.log(`✅ Attachment added successfully: ${file.name}`);
      }

      console.log(`✅ All ${files.length} attachments added to WorkflowDetails item ${workflowDetailsId}`);

    } catch (error) {
      console.error(`❌ Error adding attachments to WorkflowDetails item ${workflowDetailsId}:`, error);

      // More specific error messages
      if (error.message?.includes("does not exist")) {
        throw new Error(`WorkflowDetails list or item ${workflowDetailsId} does not exist`);
      } else if (error.message?.includes("access denied")) {
        throw new Error(`Access denied. You don't have permission to add attachments to WorkflowDetails`);
      }

      throw error;
    }
  }

  /**
   * Get attachments from WorkflowDetails item
   * @param workflowDetailsId - ID of the WorkflowDetails item
   * @returns Array of attachment info
   */
  public async getWorkflowDetailsAttachments(
    workflowDetailsId: number
  ): Promise<Array<{ name: string; url: string }>> {
    try {
      const attachments = await this.spfi.web.lists
        .getByTitle("WorkflowDetails")
        .items
        .getById(workflowDetailsId)
        .attachmentFiles();

      return attachments.map(att => ({
        name: att.FileName,
        url: `${window.location.origin}${att.ServerRelativeUrl}`
      }));

    } catch (error) {
      console.error(`Error getting attachments for WorkflowDetails item ${workflowDetailsId}:`, error);
      return [];
    }
  }


}