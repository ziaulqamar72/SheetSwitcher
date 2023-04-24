// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

function Add(first, second) {


    var values = [
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
    ];

    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {
        // Create a proxy object for the active sheet
        var sheet = ctx.workbook.worksheets.getActiveWorksheet();
        // Queue a command to write the sample data to the worksheet
        sheet.getRange("B3:D5").values = values;

        // Run the queued-up commands, and return a promise to indicate task completion
        return ctx.sync();
    })
   
  return first + second;
}
CustomFunctions.associate("ADD", Add);


function ShowTaskPane(inputValue) {
    let action = inputValue.toUpperCase();
    if (action == "SHOW") {
        Office.addin.showAsTaskpane();
    }
    else if (action == "HIDE") {
        Office.addin.hide();
    } else {
        inputValue = "";
    }
    var taskpane = action;
    return taskpane;
}
CustomFunctions.associate("ShowTaskPane", ShowTaskPane);




Office.actions.associate('SHOWTASKPANE', function () {
    return Office.addin.showAsTaskpane()
        .then(function () {
            return;
        })
        .catch(function (error) {
            return error.code;
        });
});


Office.actions.associate('HIDETASKPANE', function () {
    return Office.addin.hide()
        .then(function () {
            return;
        })
        .catch(function (error) {
            return error.code;
        });
});






//Office.initialize = function (reason) {
//    // The Office runtime is now initialized
//    // Set up event listeners here

//    // Set up event listener for keydown event on the document
//    if (Office.context.document) {
//        Office.context.document.addHandlerAsync("documentSelectionChanged", onKeyDown);
//    }
//}

//function onKeyDown(eventArgs) {
//    // Handle the keydown event here
//    var key = eventArgs.originalEvent.code;
//    console.log("Key pressed: " + key);
//}
