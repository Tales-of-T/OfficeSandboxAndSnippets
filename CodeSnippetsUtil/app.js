var CodeSnippets;
(function (CodeSnippets) {
    var Util;
    (function (Util) {
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
    })(Util = CodeSnippets.Util || (CodeSnippets.Util = {}));
})(CodeSnippets || (CodeSnippets = {}));
//# sourceMappingURL=app.js.map