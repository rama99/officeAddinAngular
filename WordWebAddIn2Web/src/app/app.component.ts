/// <reference path="../../node_modules/@types/office-js/index.d.ts" />



import { Component, OnInit } from '@angular/core';
declare var fabric: any; // Magic

@Component({
  moduleId: module.id,
  selector: 'my-app',
  templateUrl: 'app.component.html',
})
export class AppComponent implements OnInit
{
    name = 'Angular';
    messageBanner: any;    
    data: string;
    allSections: any;

    ngOnInit() {
       
      //  Office.initialize = function (reason) {
            // Initialize the FabricUI notification mechanism and hide it

            var element = document.querySelector('.ms-MessageBanner');
            this.messageBanner = new fabric.MessageBanner(element);

            this.data = "data from component";

            this.messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', 1.1)) { 
                return;
            }

            //$("#template-description").text("This sample highlights the longest word in the text you have selected in the document.");
            //$('#button-text').text("Highlight!");
            // $('#button-desc').text("Highlights the longest word."); 

            // this.loadSampleData();

            // Add a click event handler for the highlight button.
            //$('#highlight-button').click(
            //hightlightLongestWord);            
       // }

        this.loadSampleData();
    }



    hightlightLongestWord() {

        Word.run(function (context) {

           
        // Queue a command to get the current selection and then
        // create a proxy range object with the results.
        var range = context.document.getSelection();

        // variable for keeping the search results for the longest word.
        var searchResults:any;

        // Queue a command to load the range selection result.
        context.load(range, 'text');

        // Synchronize the document state by executing the queued commands
        // and return a promise to indicate task completion.
        return context.sync()
            .then(function () {

                // Get the longest word from the selection.
                var words = range.text.split(/\s+/);
                var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                // Queue a search command.
                searchResults = context.document.body.search(longestWord, { matchCase: true, matchWholeWord: true });

                // Queue a commmand to load the font property of the results.
                context.load(searchResults, 'font');

            })
            .then(context.sync)
            .then(function () {
                // Queue a command to highlight the search results.
                searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                searchResults.items[0].font.bold = true;
            })
            .then(context.sync)
    })
        .catch(this.errorHandler);
} 



    loadSampleData() {

    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;
       
        

        // Queue a commmand to clear the contents of the body.
        body.clear();
        // Queue a command to insert text into the end of the Word document body.
        body.insertText("This is a sample text inserted in the document here -- Angular2",
            Word.InsertLocation.end);

        body.insertParagraph("insert para here", "End");

        this.allSections = context.document.body.paragraphs;

        context.load(this.allSections, 'text, style');


        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {

            // Queue a command to get the last paragraph and create a 
            // proxy paragraph object.
            var paragraph = this.allSections.items[this.allSections.items.length - 1];

            // Queue a command to select the paragraph. The Word UI will 
            // move to the selected paragraph.
            paragraph.select();

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Selected the last paragraph.');
            });
        }); 


        // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
        //return context.sync();
    })
        .catch(this.errorHandler);
    }

    displaySelectedText() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
        function (result:any) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                this.showNotification('The selected text is:', '"' + result.value + '"');
            } else {
                this.showNotification('Error:', result.error.message);
            }
        });
    }

     errorHandler(error:any) {
    // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
    this.showNotification("Error:", error);
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    }

    showNotification(header:any, content:any) {
    //$("#notificationHeader").text(header);
    //$("#notificationBody").text(content);
    this.messageBanner.showBanner();
    this.messageBanner.toggleExpansion();
}




}
