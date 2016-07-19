module CodeSnippets {
    function getAllSnippets(): { [name: string]: Interface.ISnippet<Interface.ICodeExternal>; } {
        var snippets: { [name: string]: Interface.ISnippet<Interface.ICodeExternal> } = {};
        var keyword = "snippet_";
        for (var func in CodeSnippets) {
            if (func.substr(0, keyword.length) != keyword) {
                continue;
            }

            processSnippet(CodeSnippets[func]);
        }

        return snippets;

        function processSnippet(snippet: IInternalSnippet): Interface.ISnippet<Interface.ICodeExternal> {
            var output: Interface.ISnippet<Interface.ICodeExternal> = <any>snippet;
            output.code.jsOrTs = processJsOrTs(snippet.code.jsOrTs);
            output.code.compileJsCodeIfAny = isTrulyJavaScript(output.code.jsOrTs) ? null : compileTypeScript(output.code.jsOrTs);
            return output;

            function processJsOrTs(input: string | (() => OfficeExtension.IPromise<any>)): string {
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

            function compileTypeScript(input: string): string {
                return null; // FIXME
            }

            function isTrulyJavaScript(text: string) {
                try {
                    new Function(text);
                    return true;
                } catch (syntaxError) {
                    return false;
                }
            }
        }
    }
}
