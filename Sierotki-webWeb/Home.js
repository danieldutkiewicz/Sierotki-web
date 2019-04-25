'use strict';

(function () {

    Office.onReady(function () {
        $(document).ready(function () {
            if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {

                $('#fixButton').click(fixLines);

                //$('#supportedVersion').html('This code is using Word 2016 or later.');
            }
            else {
                // Just letting you know that this code will not work with your version of Word.
                $('#supportedVersion').html('This code requires Word 2016 or later.');
            }
        });
    });

    function fixLines() {
        Word.run(function (context) {

            var characters = [
                "w ", "i ", "u ", "o ", "a ", "z ",
                "W ", "I ", "U ", "O ", "A ", "Z "
            ];

            var charactersReplacement = [
                "w&nbsp;", "i&nbsp;", "u&nbsp;", "o&nbsp;", "a&nbsp;", "z&nbsp;",
                "W&nbsp;", "I&nbsp;", "U&nbsp;", "O&nbsp;", "A&nbsp;", "Z&nbsp;"
            ];

            var body = context.document.body;

            //$('#wSumm').html(characters.length);

            for (var i = 0; i < characters.length; i++) {

                var searchResult = body.search(characters[i], { matchCase: true, matchPrefix: true });
                context.load(searchResult, 'text');

                return context.sync().then(function () {

                    for (var k = 0; k < searchResult.items.length; k++) {
                        searchResult.items[k].insertHtml(charactersReplacement[i], "Replace");
                    }
                });
            }
            
            return context.sync();
        });
    }

})();