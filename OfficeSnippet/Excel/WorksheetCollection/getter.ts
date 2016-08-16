module CodeSnippets {
    export function snippet_WorksheetCollection_Getter(): IInternalSnippet {
        return {
            name: "",
            category: "",
            showInAPIPlayground: true,
            
            setup: function () {
                // TODO: Documentation folk, please either fill this in (if needed), or remove!
                return Promise.resolve();
            },
            code: {
                jsOrTs: function () {
                    return Excel.run(function (ctx) {
                        var worksheets = ctx.workbook.worksheets;
                        worksheets.load('items');
                        return ctx.sync().then(function () {
                            for (var i = 0; i < worksheets.items.length; i++) {
                                console.log(worksheets.items[i].name);
                                console.log(worksheets.items[i].index);
                            }
                        });
                    }).catch(function (error) {
                        console.log("Error: " + error);
                        if (error instanceof OfficeExtension.Error) {
                            console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        }
                    });
                }
            },
            validator: function () {
                // TODO: Documentation folk, this test will FAIL until you fill in the appropriate validation!
                return Promise.reject(new Error("Validation not defined, test bound to fail!"));
            }
        }
    }
    
}