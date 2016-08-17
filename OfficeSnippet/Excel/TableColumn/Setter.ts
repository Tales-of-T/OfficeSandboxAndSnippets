module CodeSnippets {
    export function snippet_TableColumn_Setter(): IInternalSnippet {
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
						var tableName = 'Table1';
						var tables = ctx.workbook.tables;
						var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
						var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(2);
						column.values = newValues;
						column.load('values');
						return ctx.sync().then(function () {
							console.log(column.values);
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