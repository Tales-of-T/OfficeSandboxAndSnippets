declare module CodeSnippets {
    export function getAllSnippets(): { [name: string]: Interface.ISnippet<Interface.ICodeExternal> };
}

(function () {

    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            App.initialize();
            $('#pull-snippets-button').click(pullASnippet);
        });
    };

    function pullASnippet() {
            var snippets = CodeSnippets.getAllSnippets();
            for (var func in snippets) {
                var currentSnippet = CodeSnippets[func];
            }
    }

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    App.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    App.showNotification('Error:', result.error.message);
                }
            }
        );
    }
})();