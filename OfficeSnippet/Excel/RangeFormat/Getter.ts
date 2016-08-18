module CodeSnippets {
    export function snippet_RangeFormat_Getter(): IInternalSnippet {
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
						var rangeAddress = "F:G";
						var worksheet = ctx.workbook.worksheets.getItem(sheetName);
						var range = worksheet.getRange(rangeAddress);
						range.load(["format/*", "format/fill", "format/borders", "format/font"]);
						return ctx.sync().then(function () {
							console.log(range.format.wrapText);
							console.log(range.format.fill.color);
							console.log(range.format.font.name);
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