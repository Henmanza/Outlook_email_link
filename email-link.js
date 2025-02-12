function getSelectedEmail() {
    Office.context.mailbox.item.getAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            let item = asyncResult.value;
            let emailId = item.itemId;
            
            // Construct the Outlook Web link
            let outlookWebUrl = "https://outlook.office.com/mail/id/" + encodeURIComponent(emailId);
            
            // Copy link to clipboard
            navigator.clipboard.writeText(outlookWebUrl).then(function () {
                Office.context.ui.displayDialogAsync(
                    'Copied to clipboard! Now share the link.',
                    { width: 20, height: 10 }
                );
            }, function (err) {
                console.error('Failed to copy: ', err);
            });
        } else {
            console.error('Failed to get email:', asyncResult.error);
        }
    });
}

// Register the function as a button action
Office.actions.associate("copyEmailLink", getSelectedEmail);
