module CodeSnippets {
    export function snippet_Range_GetCell(): IInternalSnippet {
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
						var sheetName = "Sheet1";
						var rangeAddress = "F5:G7";
						var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
						var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
						var formulas = [[null, null], [null, null], [null, "=G6-G5"]];
						var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
						range.numberFormat = numberFormat;
						range.values = values;
						range.formulas = formulas;
						range.load('text');
						return ctx.sync().then(function () {
							console.log(range.text);
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