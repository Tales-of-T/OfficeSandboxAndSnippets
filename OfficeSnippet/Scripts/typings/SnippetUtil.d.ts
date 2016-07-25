declare module CodeSnippets.Util {
    function cleanWorkBook(): OfficeExtension.IPromise<any>;
    function ensureSheetExists(name: string): OfficeExtension.IPromise<any>;
}
