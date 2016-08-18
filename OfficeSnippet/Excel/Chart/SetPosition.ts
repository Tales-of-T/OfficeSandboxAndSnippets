module CodeSnippets {
    export function snippet_Chart_SetPosition(): IInternalSnippet {
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
						var sheetName = "Charts";
						var rangeSelection = "A1:B4";
						var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeSelection);
						var sourceData = sheetName + "!" + "A1:B4";
						var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", range, "auto");
						chart.width = 500;
						chart.height = 300;
						chart.setPosition("C2", null);
						return ctx.sync();
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