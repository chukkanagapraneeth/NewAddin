Office.initialize = function () {
    $(document).ready(function () {
        $('#itsd').click(createITSDMail);
    });
};


function statusUpdate(icon, text) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        icon: icon,
        message: text,
        persistent: false
    });
}


function defaultStatus(event) {
    statusUpdate("icon16", "Hello World!");
}


function createITSDMail() {
    var item = Office.context.mailbox.item;
    var bodyText = 
    '<div style="background-color:#D9D9D9" ;"border:solid";"border-color:black"><p style = "background-color:#D9D9D9"; "font-family:Calibri" >' +
    '<b>INSTRUCTIONS:</b> <br /> <pre style="font-family:Calibri">- Please answer all relevant questions below</pre> <pre style="font-family:Calibri">- Do not remove any of the prepopulated text.</pre>' +
    '<pre style = "font-family:Calibri" > - The more details you provide the fewer interactions will be needed to resolve your issue/request.</pre >' +
    '<div><p><i><span style="font-size:10.0pt">Please answer the below questions for assistance.</span></i></p><p style="background:#BFBFBF">1. SAP ID:</p> <p style="font-family:Calibri"; "font-size:xx-small">&nbsp;</p> <p style="background:#BFBFBF">2. Phone number:</p>' +
    '<p style="font-family:Calibri" ;"font-size:xx-small">&nbsp;</p><p style="background:#BFBFBF">3. Issue:</p>' ;

    if (Office.context.mailbox.displayNewMessageForm) {

        Office.context.mailbox.displayNewMessageForm(
            {
                toRecipients: ["it.servicedesk@ericsson.com"],
                htmlBody: bodyText
            });
    }

}


function onDataSet(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
    }
}




