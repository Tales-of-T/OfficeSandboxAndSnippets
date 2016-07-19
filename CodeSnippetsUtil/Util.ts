module CodeSnippets.Util {
    export function ensureCleanSheet(name: string): OfficeExtension.IPromise<any> {
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

    export function ensureSheetExists(name: string): OfficeExtension.IPromise<any> {
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
}