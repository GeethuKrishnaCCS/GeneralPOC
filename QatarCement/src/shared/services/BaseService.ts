import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp/presets/all";
import { getSP } from "../PnP/pnpjsConfig";
import "@pnp/sp/attachments";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/lists"
import "@pnp/sp/fields";
import "@pnp/sp/folders";

export class BaseService {
    private sp: SPFI;
    constructor(context: WebPartContext, siteUrl: string) {
        this.sp = getSP(context);
    }

    public getCurrentUser() {
        return this.sp.web.currentUser();
    }

    public getCurrentUserGroups() {
        return this.sp.web.currentUser.groups();
    }
    public getListItems(url: string): Promise<any> {
        return this.sp.web.getList(url).items();
    }
    public async getPagedListItems(queryurl: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select('*').top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    public async getPagedFilterListItems(queryurl: string, filter: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select('*').filter(filter).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    public async getPagedSelectFilterExpandListItems(queryurl: string, select: string, filter: string, expand: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select(select).filter(filter).expand(expand).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    public getItemsSelect(queryurl: string, select: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)()
    }

    public getItemsSelectExpand(queryurl: string, select: string, expand: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .expand(expand)()
    }
    public getItemsById(queryurl: string, id: any): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .getById(id)()
    }
    public getItemsByIdSelect(queryurl: string, id: any, select: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .getById(id)
            .select(select)()
    }
    public getItemsFilter(queryurl: string, filter: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .filter(filter)()
    }
    public getItemsSelectFilter(queryurl: string, select: string, filter: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .filter(filter)()
    }
    public getItemsSelectExpandFilter(queryurl: string, select: string, expand: string, filter: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .expand(expand)
            .filter(filter)()
    }
    public async getPagedItemsSelectExpand(queryurl: string, select: string, expand: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select(select).expand(expand).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    public async getPagedItemsSelectExpandFilter(queryurl: string, select: string, expand: string, filter: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select(select).expand(expand).filter(filter).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    public async getPagedItemsFilter(queryurl: string, filter: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select('*').filter(filter).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    public getItemsByIdSelectExpand(queryurl: string, id: any, select: string, expand: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .getById(id)
            .select(select)
            .expand(expand)()
    }
    public createNewItem(url: string, data: any): Promise<any> {
        return this.sp.web.getList(url).items.add(data);
    }




    public updateItem(url: string, data: any, id: number): Promise<any> {
        return this.sp.web.getList(url).items.getById(id).update(data);
    }
    public DeleteItem(url: string, id: number): Promise<any> {
        return this.sp.web.getList(url).items.getById(id).delete();
    }
    public async createListItemInsideFolder(): Promise<any> {
        // Set the list URL and folder name
        const listUrl = "/sites/ProcurementDMS/Lists/RequestMaster";
        const folderName = "MARKETING";
        return this.sp.web.getList(listUrl).items.add({
            Title: "Item inside folder",
            FileSystemObjectType: 0,
            FileLeafRef: "Item inside folder",
            'FolderPath': { 'DecodedUrl': `${listUrl}/${folderName}` }
            // Add other fields as needed
        });
    }

    public getSelectExpand(queryurl: string, select: string, expand: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .expand(expand)()
    }
    public async getUser(userId: number): Promise<any> {
        return this.sp.web.getUserById(userId)();
    }

    public async getFileContent(fileUrl: string): Promise<ArrayBuffer> {
        return this.sp.web.getFileByServerRelativePath(fileUrl).getBuffer();
    }

    public getDocLibrary(url: string): Promise<any> {
        return this.sp.site.getDocumentLibraries(url);
    }

    public async uploadDocument(libraryName: string, Filename: any, filedata: any): Promise<any> {
        const response = await this.sp.web.getFolderByServerRelativePath(libraryName).files.addUsingPath(Filename, filedata, { Overwrite: true });
        return response
    }

    public getLibraryItem(queryurl: string, id: number): Promise<any> {
        return this.sp.web.getList(queryurl).items.getById(id).select("FileRef,FileLeafRef")()
    }
    public async getFolderContent(FolderName: string): Promise<ArrayBuffer> {
        return this.sp.web.getFolderByServerRelativePath(FolderName).files();
    }

    public getItemsSelectFilterExpand(queryurl: string, select: string, expand: string, filter: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .expand(expand)
            .filter(filter)()
    }


    // public async uploadFileandUpdateProperty(libraryname: string, filename: string, filedata: any, metadata: any): Promise<any> {
    //     const file = await this.sp.web.getFolderByServerRelativePath(libraryname)
    //         .files.addUsingPath(filename, filedata, { Overwrite: true });
    //     const item = await this.sp.web.getFileByServerRelativePath(file.ServerRelativeUrl).getItem();
    //     item.update(metadata);
    //     return file;
    // }


}