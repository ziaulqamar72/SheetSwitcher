Office.onReady(function () {
    //var doc = Office.context.document;
    //doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function (eventArgs) {
    //    // do something when the selection changes
    //});
    // Add event listener on keypress
    document.addEventListener('keypress', (event) => {
        var name = event.key;
        var code = event.code;
        console.log(event.ctrlKey);
        // Alert the key name and key code on keydown
        console.log(`Key pressed ${name} \r\n Key code value: ${code}`);


        if (code === 'Backquote') {
            Office.addin.hide();
        };








    }, false);

})