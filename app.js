'use strict';

(function () {
    // API_URL = 'http://nlp-ryze.cs.nthu.edu.tw:1214/translate/'
    // API_URL_1 = 'http://nlp-ryze.cs.nthu.edu.tw:1215/translate/'
    // HEADERS = {'Content-Type': 'application/json; charset=UTF-8'}

    // The initialize function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Use this to check whether the API is supported in the Word client.
            if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                // Do something that is only available via the new APIs
                $('#emerson').click(insertEmersonQuoteAtSelection);
                $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                $('#proverb').click(insertChineseProverbAtTheEnd);
                $('#run').click(goToMainPage);
                $('#supportedVersion').html('This code is using Word 2016 or greater.');
            }
            else {
                // Just letting you know that this code will not work with your version of Word.
                $('#supportedVersion').html('This code requires Word 2016 or greater.');
            }
        });
    };

    function insertEmersonQuoteAtSelection() {
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Queue a command to get the current selection.
            // Create a proxy range object for the selection.
            var range = thisDocument.getSelection();

            // Queue a command to replace the selected text.
            range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from Ralph Waldo Emerson.');
            });
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }
    function goToMainPage(){
       
        
        Word.run(function (context) {
            document.getElementById("show").textContent = context.document.body.text;
          //  var thisDocument = context.document.getSelection();
           
            // gec_it_post("You are an girl.");

            // Create a proxy object for the document.
           
            
           //  document.getElementById("show").textContent = thisDocument.textContent;


            // var paragraphs = context.document.getSelection().paragraphs;
            // paragraphs.load();
            // return context.sync().then(function () {
                   
            //     document.getElementById("show").textContent = paragraphs;
            //     // paragraphs.items[0].insertText(' New sentence in the paragraph.',
            //     //                             Word.InsertLocation.end);
            // }).then(context.sync);



            // Queue a command to get the current selection.
            // Create a proxy range object for the selection.
            var range = thisDocument.getSelection();

           // get_it_post(thisDocument);

            // // Queue a command to replace the selected text.
            // range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from Ralph Waldo Emerson.');
            });
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });

    }

    function gec_it_post(query){

        document.getElementById("show").textContent = "result:" + query;
        $.ajax({
            type: "POST",
            url: API_URL,
            data: JSON.stringify({text: query}),
            headers: HEADERS,
            dataType: 'json',
            success: function (data) {
                console.log("success")
                document.getElementById("show").textContent = "success"; 
                console.info(data);
               // document.getElementById("show").textContent = data.result;
            }, 
            error: function(XMLHttpRequest, textStatus, errorThrown) { 
                console.log("Status: " + textStatus); 
                console.log("Error: " + errorThrown);
              document.getElementById("show").textContent = "error "; 
            } 
          })
	}

    function insertChekhovQuoteAtTheBeginning() {
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a command to insert text at the start of the document body.
            body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from Anton Chekhov.');
            });
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }

    function insertChineseProverbAtTheEnd() {
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a command to insert text at the end of the document body.
            body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from a Chinese proverb.');
            });
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }
})();