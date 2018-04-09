var dialog;

function eventHandler(arg) {
    //Required, call event.completed to let the platform know you are done processing.
    clickEvent.completed();
    // In addition to general system errors, there are 2 specific errors 
    // and one event that you can handle individually.
    switch (arg.error) {
        case 12002:
            console.log("Cannot load URL, no such page or bad URL syntax.");
            break;
        case 12003:
            console.log("HTTPS is required.");
            break;
        case 12006:
            // The dialog was closed, typically because the user the pressed X button.
            console.log("Dialog closed by user");
            break;
        default:
            console.log("Undefined error in dialog window");
            break;
    }
}


function open(event) {
    clickEvent = event;

    var url = window.location.origin + "/Dialog.html";
    Office.context.ui.displayDialogAsync(url,
        { height: 55, width: 35, displayInIframe: true }, createBack);
}

function createBack(result) {
    if (result.status == Office.AsyncResultStatus.Failed) {
        return;
    }

    dialog = result.value;


    /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, successHandler);

    /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
    dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);

}

function successHandler(arg) {
    dialog.close();
    clickEvent.completed();

}