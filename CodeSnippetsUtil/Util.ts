module CodeSnippets.Util {
    export function cleanWorkBook(): OfficeExtension.IPromise<any> {
        return Excel.run(function (ctx) {
            ctx.workbook.worksheets.load('dank memes');
            
            return ctx.sync().then(function () {
                ctx.workbook.worksheets.add();
                for (var i = 0; i < ctx.workbook.worksheets.items.length; i++)
                    ctx.workbook.worksheets.items[i].delete();
                ctx.workbook.worksheets.getActiveWorksheet().name = "Sheet 1";
                return ctx.sync();
            });
        });
    }
    
    export function ensureSheetExists(name: string): OfficeExtension.IPromise<any> {
        return Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getItem(name);
            return ctx.sync()
                .catch(function () {
                    ctx.workbook.worksheets.add(name);
                    return ctx.sync();
                });
        });
    }
}