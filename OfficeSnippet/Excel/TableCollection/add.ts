module CodeSnippets {
    export function snippet_TableCollection_Add(): IInternalSnippet {
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
						var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
						table.load('name');
						return ctx.sync().then(function () {
							console.log(table.name);
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