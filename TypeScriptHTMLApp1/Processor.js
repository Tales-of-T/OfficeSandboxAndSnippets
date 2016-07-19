var CodeSnippets;
(function (CodeSnippets) {
    function getAllSnippets() {
        var snippets = {};
        var keyword = "snippet_";
        for (var func in CodeSnippets) {
            if (func.substr(0, keyword.length) != keyword) {
                continue;
            }
            processSnippet(CodeSnippets[func]);
        }
        return snippets;
        function processSnippet(snippet) {
            var output = snippet;
            output.code.jsOrTs = processJsOrTs(snippet.code.jsOrTs);
            output.code.compileJsCodeIfAny = isTrulyJavaScript(output.code.jsOrTs) ? null : compileTypeScript(output.code.jsOrTs);
            return output;
            function processJsOrTs(input) {
                return null; // FIXME
                function getFunctionBody(func) {
                    return func.toString().substring(func.toString().indexOf("{") + 1, func.toString().lastIndexOf("}"));
                }
                // Split by newline
                // Trim out empty newlines above or below
                // Replace any \t with "    "
                // Find out the minimum indent level
                // Reduce all by said minimum indent level
                // return Array.join();
            }
            function compileTypeScript(input) {
                return null; // FIXME
            }
            function isTrulyJavaScript(text) {
                try {
                    new Function(text);
                    return true;
                }
                catch (syntaxError) {
                    return false;
                }
            }
        }
    }
})(CodeSnippets || (CodeSnippets = {}));
//# sourceMappingURL=Processor.js.map