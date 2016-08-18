module CodeSnippets {
    export function snippet_ChartFont_Setter(): IInternalSnippet {
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
						var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
						title.format.font.name = "Calibri";
						title.format.font.size = 12;
						title.format.font.color = "#FF0000";
						title.format.font.italic = false;
						title.format.font.bold = true;
						title.format.font.underline = "None";
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