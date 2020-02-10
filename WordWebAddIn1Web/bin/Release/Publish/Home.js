
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");

                $('#highlight-button').click(displaySelectedText);
                return;
            }

            $("#template-description").text("This sample highlights the longest word in the text you have selected in the document.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the longest word.");

            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#nameBtn').click(highlightFromArray);
            $('#postalCodeBtn').click(highlightPostalCodes);
            $('#streetNameBtn').click(highlightStreetName);
        });
    };

    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "Hej Gabriel! Välkommen till Supergatan eller Ytvägen. Det finns två postnummer August känner till: 603 21 och 70300.",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();

        })
            .catch(errorHandler);
    }

    function hightlightLongestWord() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });
                    console.log(longestWord);

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });
                    console.log(searchResults);

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }

    function highlightFromArray() {

        let searchList = ["Gabriel", "August"];

        searchList.forEach(function (searchWord) {
            return Word.run(function (context) {

                // Tell word for search for a word
                let searchResult = context.document.body.search(searchWord);

                // Load the properties for the result
                context.load(searchResult);

                // Execute the batch
                return context.sync()
                    .then(function () {

                        // Loop through the results
                        searchResult.items.forEach(function (result) {
                            result.font.highlightColor = 'green';
                            result.insertText("_____", Word.InsertLocation.replace);
                        });
                    });
            })
                .catch(errorHandler);
        });
    }

    function highlightPostalCodes() {

        var searchList = [];

        Word.run(function (context) {
            var body = context.document.body;

            context.load(body, 'text');
            return context.sync()
                .then(function () {
                    var words = body.text.split(/\W+/);

                    $.each(words, function (index, word) {
                        // Check if number and if length is 5
                        if (isNaN(word) == false && word.length == 5) {
                            searchList.push(word);
                        }
                        //Check if number and if length is 3 and string after is number and length is 2
                        else if (isNaN(word) == false && word.length == 3 && isNaN(words[index + 1]) == false && words[index + 1].length == 2) {
                            searchList.push(word + " " + words[index + 1]);
                        }
                    })
                })
                .then(function () {
                    searchList.forEach(function (searchWord) {
                        return Word.run(function (context) {

                            // Tell word for search for a word
                            let searchResult = context.document.body.search(searchWord);

                            // Load the properties for the result
                            context.load(searchResult);

                            // Execute the batch
                            return context.sync()
                                .then(function () {

                                    // Loop through the results
                                    searchResult.items.forEach(function (result) {
                                        result.font.highlightColor = 'red';
                                        result.insertText("_____", Word.InsertLocation.replace);
                                    });
                                });
                        })
                            .catch(errorHandler);
                    });
                });
        })
        .catch(errorHandler);
    }

    function highlightStreetName() {
        let searchList = ["Väg", "Gata"];

        searchList.forEach(function (searchWord) {
            return Word.run(function (context) {

                // Tell word for search for a word
                let searchResult = context.document.body.search(searchWord, { matchWholeWord: false, matchCase: false });

                // Load the properties for the result
                context.load(searchResult);

                // Execute the batch
                return context.sync()
                    .then(function () {

                        // Loop through the results
                        searchResult.items.forEach(function (result) {
                            result.font.highlightColor = 'yellow';
                            result.insertText("_____", Word.InsertLocation.replace);
                        });
                    });
            })
                .catch(errorHandler);
        });
    }

    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
