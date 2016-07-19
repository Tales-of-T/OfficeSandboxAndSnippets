var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var Excel;
(function (Excel) {
    /**
     * The RequestContext object facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the request context is required to get access to the Excel object model from the add-in.
     */
    var RequestContext = (function (_super) {
        __extends(RequestContext, _super);
        function RequestContext() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(RequestContext.prototype, "workbook", {
            get: function () {
                return this.m_workbook;
            },
            enumerable: true,
            configurable: true
        });
        return RequestContext;
    })(OfficeExtension.ClientRequestContext);
    Excel.RequestContext = RequestContext;
    /**
     * Executes a batch script that performs actions on the Excel object model. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in an Excel.RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the request context is required to get access to the Excel object model from the add-in.
     */
    function run(batch) {
        return OfficeExtension.ClientRequestContext._run(function () { return new Excel.RequestContext(); }, batch);
    }
    Excel.run = run;
})(Excel || (Excel = {}));
/// <reference path="../test/jscript/officejs/table1_0test.ts" />
var Excel;
(function (Excel) {
    Excel._RedirectV1APIs = false;
    // For now, it is set to "false" by default, but can be toggled via a checkbox at the bottom of the test agave.
    Excel._V1APIMap = {
        "GetDataAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingGetData(callArgs); },
            postprocess: getDataCommonPostprocess
        },
        "GetSelectedDataAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.getSelectedData(callArgs); },
            postprocess: getDataCommonPostprocess
        },
        "GoToByIdAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.gotoById(callArgs); }
        },
        "AddColumnsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddColumns(callArgs); }
        },
        "AddFromSelectionAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromSelection(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "AddFromNamedItemAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromNamedItem(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "AddFromPromptAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromPrompt(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "AddRowsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddRows(callArgs); }
        },
        "GetByIdAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingGetById(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "ReleaseByIdAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingReleaseById(callArgs); }
        },
        "GetAllAsync": {
            call: function (ctx) { return ctx.workbook._V1Api.bindingGetAll(); },
            postprocess: function (response) {
                return response.bindings.map(function (descriptor) { return postprocessBindingDescriptor(descriptor); });
            }
        },
        "DeleteAllDataValuesAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingDeleteAllDataValues(callArgs); }
        },
        "SetSelectedDataAsync": {
            preprocess: function (callArgs) {
                var preimage = callArgs["cellFormat"];
                if (window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                return callArgs;
            },
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.setSelectedData(callArgs); }
        },
        "SetDataAsync": {
            preprocess: function (callArgs) {
                var preimage = callArgs["cellFormat"];
                if (window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                return callArgs;
            },
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetData(callArgs); }
        },
        "SetFormatsAsync": {
            preprocess: function (callArgs) {
                var preimage = callArgs["cellFormat"];
                if (window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                return callArgs;
            },
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetFormats(callArgs); }
        },
        "SetTableOptionsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetTableOptions(callArgs); }
        },
        "ClearFormatsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingClearFormats(callArgs); }
        },
    };
    function postprocessBindingDescriptor(response) {
        // Due to capitalization inconsistency, create a new object based on the response
        var bindingDescriptor = {
            BindingColumnCount: response.bindingColumnCount,
            BindingId: response.bindingId,
            BindingRowCount: response.bindingRowCount,
            bindingType: response.bindingType,
            HasHeaders: response.hasHeaders
        };
        return window.OSF.DDA.OMFactory.manufactureBinding(bindingDescriptor, window.Microsoft.Office.WebExtension.context.document);
    }
    function getDataCommonPostprocess(response, callArgs) {
        var isPlainData = response.headers == null;
        var data;
        if (isPlainData) {
            // Rows will contain the data
            data = response.rows;
        }
        else {
            data = response;
        }
        data = window.OSF.DDA.DataCoercion.coerceData(data, callArgs[window.Microsoft.Office.WebExtension.Parameters.CoercionType]);
        return data == undefined ? null : data;
    }
})(Excel || (Excel = {}));
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var range = ctx.workbook.getSelectedRange();
    range.clear();
    range.set({
        values: [["Test"]],
        format: {
            fill: {
                color: "yellow"
            },
            font: {
                color: "green",
                size: 15,
            }
        }
    });
    return ctx.sync();
});
var range = new Excel.Range();
range.
;
var chart = new Excel.Chart();
var worksheet = new Excel.Worksheet();
var valueMultipied = range.values * 2;
chart.title = new Excel.ChartTitle(); // "Sales"; /* intead of of chart.title.text */
worksheet.naame = "January Data";
function addData(table, data) {
    table.
    ;
}
Word.run(function (context) {
    var results = context.document.body.search("Contoso", { matchCase: false });
    context.load(results);
    return context.sync()
        .then(function () {
        for (var i = 0; i < results.items.length; i++) {
            results.items[i].font.color = "#FF0000";
            results.items[i].font.highlightColor = "#FFFF00";
            results.items[i].font.bold = true;
            var cc = results.items[i].insertContentControl();
            cc.tag = "customer";
            cc.title = "Customer Name";
        }
    })
        .then(context.sync)
        .then(function () {
        // ...
    })
        .catch(function (error) {
        console.log(JSON.stringify(error));
    });
});
//# sourceMappingURL=app.js.map