/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

// Add start-up logic code here, if any.
Office.onReady();

function onNewMessageComposeHandler(event) {
    const item = Office.context.mailbox.item;
    const signatureIcon = "iVBORw0KGgoAAAANSUhEUgAAACcAAAAnCAMAAAC7faEHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAzUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKMFRskAAAAQdFJOUwAQIDBAUGBwgI+fr7/P3+8jGoKKAAAACXBIWXMAAA7DAAAOwwHHb6hkAAABT0lEQVQ4T7XT2ZalIAwF0DAJhMH+/6+tJOQqot6X6joPiouNBo3w9/Hd6+hrYnUt6vhLcjEAJevVW0zJxABSlcunhERpjY+UKoNN5+ZgDGu2onNz0OngjP2FM1VdyBW1LtvGeYrBLs7U5I1PTXZt+zifcS3Icw2GcS3vxRY3Vn/iqx31hUyTnV515kdTfbaNhZLI30AceqDiIo4tyKEmJpKdP5M4um+nUwfDWxAXdzqMNKQ14jLdL5ntXzxcRF440mhS6yu882Kxa30RZcUIjTCJg7lscsR4VsMjfX9Q0Vuv/Wd3YosD1J4LuSRtaL7bzXGN1wx2cytUdncDuhA3fu6HPTiCvpQUIjZ3sCcHVbvLtbNTHlysx2w9/s27m9gEb+7CTri6hR1wcTf2gVf3wBRe3CMbcHYvTODkXhnD0+178K/pZ9+n/C1ru/2HAPwAo7YM1X4+tLMAAAAASUVORK5CYII=";

    // Get the sender's account information.
    item.from.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            console.log(result.error.message);
            event.completed();
            return;
        }

        // Create a signature based on the sender's information.
        const name = result.value.displayName;
        const options = { asyncContext: name, isInline: true };
        item.addFileAttachmentFromBase64Async(signatureIcon, "signatureIcon.png", options, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.log(result.error.message);
                event.completed();
                return;
            }

            // Add the created signature to the message.
            const signature = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style type="text/css">
        .ImprintUniqueID {
            margin: 0cm 0cm 0pt;
        }
        .ImprintUniqueIDTable {
            margin: 0cm 0cm 0pt;
            border-collapse: collapse;
            width: 100%;
        }
        .Section1 {
            page: Section1;
        }
    </style>
</head>
<body>
    <table class="ImprintUniqueIDTable" cellspacing="0" cellpadding="0" border="0">
        <tbody>
            <tr>
                <td><font style="font-family: Calibri; font-size: 14pt; color: #001D56; font-weight: bold;">Fullname</font></td>
            </tr>
            <tr>
                <td style="padding-bottom: 5px; padding-top: 5px;">
                    <table class="ImprintUniqueIDTable" cellspacing="0" cellpadding="0" border="0" style="width: auto;">
                        <tbody>
                            <tr>
                                <td style="padding-bottom: 5px; padding-top: 5px; padding-left: 10px; padding-right: 10px; background-color: #6bdad4;">
                                    <font style="font-family: Calibri; font-size: 11pt; color: #001D56; font-weight: bold;">Title</font>
                                    <font style="font-family: Calibri; font-size: 11pt; color: #001D56; font-weight: normal;"> - Department </font>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="border-top: #001d56 1px solid; padding-top: 5px;">
                    <table class="ImprintUniqueIDTable" cellspacing="0" cellpadding="0" border="0" style="width: auto;">
                        <tbody>
                            <tr>
                                <td style="padding-right: 10px;" align="center">
                                    <img width="19" height="15" style="border: 0px solid;" src="https://www.prudentialbank.com.gh/images/icons/Tel Icon.png" alt="Telephone">
                                </td>
                                <td>
                                    <font style="font-family: Calibri; font-size: 11pt; color: #000001;">telephone Ext: 1076</font>
                                </td>
                            </tr>
                            <tr>
                                <td style="padding-top: 5px; padding-right: 10px;" align="center">
                                    <a href="#" target="_blank">
                                        <img width="8" height="15" style="border: 0px solid;" src="https://www.prudentialbank.com.gh/images/icons/Phone Icon.png" alt="Mobile">
                                    </a>
                                </td>
                                <td style="padding-top: 5px;">
                                    <font style="font-family: Calibri; font-size: 11pt; color: #000001;">phone number</font>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="padding-top: 14px;">
                    <font style="font-family: Calibri; font-size: 11pt;">Prudential Bank Ltd., 8 John Harmond Street, Ring Road Central</font><br>
                    <font style="font-family: Calibri; font-size: 11pt;">Accra - Ghana</font>
                </td>
            </tr>
            <tr>
                <td style="padding-top: 5px;">
                    <table class="ImprintUniqueIDTable" cellspacing="0" cellpadding="0" border="0" style="width: auto;">
                        <tbody>
                            <tr>
                                <td style="padding-right: 5px;">
                                    <a href="https://www.facebook.com/prudentialbankgh" target="_blank">
                                        <img width="25" height="25" style="border: 0px solid;" src="https://www.prudentialbank.com.gh/images/email/EmailIcons/facebook.png" alt="Facebook">
                                    </a>
                                </td>
                                <td style="padding-right: 5px;">
                                    <a href="https://www.instagram.com/prudentialbankgh/" target="_blank">
                                        <img width="25" height="25" style="border: 0px solid;" src="https://www.prudentialbank.com.gh/images/email/EmailIcons/instagram.png" alt="Instagram">
                                    </a>
                                </td>
                                <td style="padding-right: 5px;">
                                    <a href="https://twitter.com/pblghana" target="_blank">
                                        <img width="25" height="25" style="border: 0px solid;" src="https://www.prudentialbank.com.gh/images/email/EmailIcons/twitter.png" alt="Twitter">
                                    </a>
                                </td>
                                <td>
                                    <a href="https://www.linkedin.com/company/prudential-bank-gh" target="_blank">
                                        <img width="25" height="25" style="border: 0px solid;" src="https://www.prudentialbank.com.gh/images/email/EmailIcons/linkedin.png" alt="LinkedIn">
                                    </a>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </td>
            </tr>
        </tbody>
    </table>
    <br>
    <a href="https://www.prudentialbank.com.gh/" target="_blank">
        <img width="100%" height="10%" style="border: 0px solid;" src="https://www.prudentialbank.com.gh/images/EBanners/randombanners/2.jpg" alt="Background Image">
    </a>
</body>
</html>` + result.asyncContext;
            item.body.setSignatureAsync(signature, { coercionType: Office.CoercionType.Html }, (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.log(result.error.message);
                    event.completed();
                    return;
                }

                // Show a notification when the signature is added to the message.
                // Important: Only the InformationalMessage type is supported in Outlook mobile at this time.
                const notification = {
                    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                    message: "Company signature added.",
                    icon: "none",
                    persistent: false                        
                };
                item.notificationMessages.addAsync("signature_notification", notification, (result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.log(result.error.message);
                        event.completed();
                        return;
                    }

                    event.completed();
                });
            });
        });
    });
}
