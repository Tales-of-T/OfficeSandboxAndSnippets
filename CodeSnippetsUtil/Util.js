var CodeSnippets;
(function (CodeSnippets) {
    var Util;
    (function (Util) {
        function ensureCleanSheet(name) {
            return Excel.run(function (ctx) {
                var sheet = ctx.workbook.worksheets.getItemOrNull(name);
                return ctx.sync().then(function () {
                    if (!sheet.isNull) {
                        ctx.workbook.worksheets.getItem(name).delete();
                    }
                    ctx.workbook.worksheets.add(name);
                }).then(ctx.sync);
            });
        }
        Util.ensureCleanSheet = ensureCleanSheet;
        function cleanWorkBook() {
            return Excel.run(function (ctx) {
                for (var worksheet in ctx.workbook.worksheets.items) {
                    worksheet.delete();
                }
                ;
                return ctx.sync().then(function () {
                });
            });
        }
        Util.cleanWorkBook = cleanWorkBook;
        function ensureSheetExists(name) {
            return Excel.run(function (ctx) {
                var sheet = ctx.workbook.worksheets.getItemOrNull(name);
                return ctx.sync()
                    .then(function () {
                    if (sheet.isNull) {
                        ctx.workbook.worksheets.add(name);
                    }
                })
                    .then(ctx.sync);
            });
        }
        Util.ensureSheetExists = ensureSheetExists;
    })(Util = CodeSnippets.Util || (CodeSnippets.Util = {}));
})(CodeSnippets || (CodeSnippets = {}));
//# sourceMappingURL=Util.js.map