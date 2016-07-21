module CodeSnippets {
    export function getAllSnippets(): { [name: string]: Interface.ISnippet<Interface.ICodeExternal>; } {
        var snippets: { [name: string]: Interface.ISnippet<Interface.ICodeExternal> } = {};
        var keyword = "snippet_";
        for (var func in CodeSnippets) {
            if (func.substr(0, keyword.length) != keyword) {
                continue;
            }

            processSnippet((<any>CodeSnippets)[func]);
        }

        return snippets;

        function processSnippet(snippet: IInternalSnippet): Interface.ISnippet<Interface.ICodeExternal> {
            var output: Interface.ISnippet<Interface.ICodeExternal> = <any>snippet;
            output.code.jsOrTs = processJsOrTs(snippet.code.jsOrTs);
            output.code.compileJsCodeIfAny = isTrulyJavaScript(output.code.jsOrTs) ? null : compileTypeScript(output.code.jsOrTs);
            return output;

            function processJsOrTs(input: string | (() => OfficeExtension.IPromise<any>)): string {

                function getFunctionBody(func: string | (() => OfficeExtension.IPromise<any>)): string {
                    return func.toString().substring(func.toString().indexOf("{") + 1, func.toString().lastIndexOf("}"));
                }
                // Split by new line, and remove empty new lines
                var inputStrings = getFunctionBody(input).match(/.+/g);
                var minIndex = Infinity;

                inputStrings.forEach(function (element, indx) {
                    element.replace("\t", "    "); // replace tabs
                    var currentMin = element.search(/\S/); // find the first non-whitespace character
                    if (currentMin == -1) { // if line is empty
                        inputStrings.splice(indx, 1); // get rid of it
                    } else if (currentMin < minIndex) { // else, if the minimum whitespace count is smaller than any previous row
                        minIndex = currentMin; // set to the current min val
                    }
                });

                // reduce all by minimum indent level
                inputStrings.forEach(function (element) {
                    element.slice(minIndex);
                });
                // return Array.join();
                return inputStrings.join(); 
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
