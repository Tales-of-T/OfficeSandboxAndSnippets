module CodeSnippets {
    export module Interface {
        export interface ISnippet<T extends ICodeExternal | ICodeInternal> {
            name: string,
            category?: string,
            description?: string /*if for API Tutorial */,
            showInAPIPlayground?: boolean /* for snippets that we want to show in API tutorial in addition to documentation; default is true if not specified */,

            setup?: () => OfficeExtension.IPromise<any> /* note: for purposes of testing, for it's generally best to either create a new sheet, or fully delete the existing one.  You will also want to delete any existing tables or charts, if you're referring to any by name.  Truthfully, it's probably best to just start with calling a helper method that will add one sheet and delete all the rest (call it something like "blankOutWorkbook") */,
            code: T,            
            validator?: () => OfficeExtension.IPromise<any>,
            customExecutorIfAny?: (
                    setup: () => OfficeExtension.IPromise<any>,
                    code: T,
                    validator: () => OfficeExtension.IPromise<any>
                ) => OfficeExtension.IPromise<any>, /* for special cases where instead of running the usual setup-code-validate syntax, need to do something special. */
        }

        export interface ICodeInternal {
            jsOrTs: string | (() => OfficeExtension.IPromise<any>), /* function (hence JS already) or JS as string or TS as string */
            htmlIfAny?: string,
            cssIfAny?: string,
        }

        export interface ICodeExternal {
            jsOrTs: string, /* with newlines and tabs; can be TS syntax */
            compileJsCodeIfAny?: string, /* used if "code" was TS and so needed to compile.  Otherwise just eval the code */
            htmlIfAny?: string,
            cssIfAny?: string,
        }
    }

    export type IInternalSnippet = Interface.ISnippet<Interface.ICodeInternal>
}
