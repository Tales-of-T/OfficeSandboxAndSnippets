declare module CodeSnippets {
    module Interface {
        interface ISnippet<T extends ICodeExternal | ICodeInternal> {
            name: string;
            category?: string;
            description?: string;
            showInAPIPlayground?: boolean;
            setup?: () => OfficeExtension.IPromise<any>;
            code: T;
            validator?: () => OfficeExtension.IPromise<any>;
            customExecutorIfAny?: (setup: () => OfficeExtension.IPromise<any>, code: T, validator: () => OfficeExtension.IPromise<any>) => OfficeExtension.IPromise<any>;
        }
        interface ICodeInternal {
            jsOrTs: string | (() => OfficeExtension.IPromise<any>);
            htmlIfAny?: string;
            cssIfAny?: string;
        }
        interface ICodeExternal {
            jsOrTs: string;
            compileJsCodeIfAny?: string;
            htmlIfAny?: string;
            cssIfAny?: string;
        }
    }
    type IInternalSnippet = Interface.ISnippet<Interface.ICodeInternal>;
}
declare module CodeSnippets {
    function getAllSnippets(): {
        [name: string]: Interface.ISnippet<Interface.ICodeExternal>;
    };
}
declare module CodeSnippets {
    function snippet_Range_GetCell(): IInternalSnippet;
}
