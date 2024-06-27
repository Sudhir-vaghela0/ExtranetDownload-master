/* eslint-disable @typescript-eslint/no-explicit-any */
export interface ISharePointService {
    getListItems: (listTitle: string, orderBy: string) => Promise<any>;
    getFilteredListItems: (listId: string, filterQuery: string, orderBy: string, selects: string[], expands: string[]) => Promise<any>;
    getFileBlob: (path: string) => Promise<Blob>;
}