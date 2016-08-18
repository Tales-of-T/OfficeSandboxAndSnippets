module CodeSnippets {
    export function snippet_Range_BoundingRect(): IInternalSnippet {
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
						var rangeAddress = "D4:G6";
						var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
						var range = range.getBoundingRect("G4:H8");
						range.load('address');
						return ctx.sync().then(function () {
							console.log(range.address); // Prints Sheet1!D4:H8
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