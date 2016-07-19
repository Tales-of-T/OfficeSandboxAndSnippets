var CodeSnippets;
(function (CodeSnippets) {
    function snippet_Range_GetCell() {
        return {
            name: "Get Cell",
            category: "Range",
            showInAPIPlayground: true,
            setup: function () {
                return CodeSnippets.Util.ensureSheetExists("Sheet1");
            },
            code: {
                jsOrTs: function () {
                    return Excel.run(function (ctx) {
                        var sheetName = "Sheet1";
                        var rangeAddress = "B3:D8";
                        var worksheet = ctx.workbook.worksheets.getItem(sheetName);
                        var range = worksheet.getRange(rangeAddress);
                        var cell = range.getCell(2, 1);
                        cell.load('address');
                        return ctx.sync().then(function () {
                            console.log(cell.address);
                        });
                    });
                }
            },
            validator: function () {
                return Excel.run(function (ctx) {
                    var sheetName = "Sheet1";
                    var rangeAddress = "B3:D8";
                    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
                    var range = worksheet.getRange(rangeAddress);
                    var cell = range.getCell(2, 1);
                    cell.load('address');
                    return ctx.sync().then(function () {
                        assert.equal(cell.address, "Sheet1!C5");
                    });
                });
            }
        };
    }
})(CodeSnippets || (CodeSnippets = {}));
//# sourceMappingURL=GetCell.js.map