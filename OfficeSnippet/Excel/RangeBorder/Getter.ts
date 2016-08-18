module CodeSnippets {
    export function snippet_RangeBorder_Getter(): IInternalSnippet {
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
						var rangeAddress = "A1:F8";
						var worksheet = ctx.workbook.worksheets.getItem(sheetName);
						var range = worksheet.getRange(rangeAddress);
						var borders = range.format.borders;
						borders.load('items');
						return ctx.sync().then(function () {
							console.log(borders.count);
							for (var i = 0; i < borders.items.length; i++) {
								console.log(borders.items[i].sideIndex);
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