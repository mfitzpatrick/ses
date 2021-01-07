/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */
Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("quick-reply-form").onsubmit = submitMsg;
        run();
    }
});

/*
 * When the document is initialised, run this function to register a handler that will execute
 * when the read-email-context changes.
 */
Office.initialize = function(reason) {
    $(document).ready(function() {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function(eventArgs) {
            run();
        });
    });
}

/*
 * Main handler function for this task pane. Fetch all messages related to the currently-selected
 * message and display them in the pane.
 */
export async function run() {
    // Get a reference to the current message
    var item = Office.context.mailbox.item;
    if (null == item) {
        console.log("No mailbox item");
    } else {
        updateSenderContext(item.from);
        getRelatedMsgs(item.from.emailAddress);
    }
}

/*
 * Update the sender-information context fields at the top of the pane.
 */
function updateSenderContext(from) {
    var currentAddr = document.querySelector("#contact-email");
    if (from.emailAddress != currentAddr.textContent) {
        //clear the list
        document.querySelector("#app-body").innerHTML = '';
    }
    document.querySelector("#contact-name").textContent = from.displayName;
    currentAddr.textContent = from.emailAddress;
}

/*
 * Wrap 'request' SOAP XML in a SOAP envelope for sending via EWS.
 */
function wrapRequest(request) {
    var wrapped =
        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '  <soap:Header><t:RequestServerVersion Version="Exchange2013" /></soap:Header>' +
        '  <soap:Body>' +
             request +
        '  </soap:Body>' +
        '</soap:Envelope>';
    return wrapped;
}

/*
 * Get a recent history of messages related to the specified contact.
 * This function uses an EWS request to search for all messages containing the emailAddr specified.
 * For all messages found, it extracts the ID and key values, then performs a second EWS request
 * to retrieve the message bodies.
 */
function getRelatedMsgs(emailAddr) {
    var request = wrapRequest(
        '<m:FindItem>' +
        '  <m:ItemShape>' +
        '    <t:BaseShape>IdOnly</t:BaseShape>' +
        '    <t:AdditionalProperties>' +
        '      <t:FieldURI FieldURI="item:IsFromMe" />' +
        '    </t:AdditionalProperties>' +
        '  </m:ItemShape>' +
        '  <m:ParentFolderIds><t:DistinguishedFolderId Id="inbox" /></m:ParentFolderIds>' +
        '  <m:QueryString>' + emailAddr + '</m:QueryString>' +
        '</m:FindItem>'
    );

    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
        var response = $.parseXML(result.value);
        var msgs = response.getElementsByTagName("t:Message");
        var idTuples = [];
        for (let item of msgs) {
            var id = item.getElementsByTagName("t:ItemId")[0];
            var msgID = id.getAttribute("Id");
            var changeKey = id.getAttribute("ChangeKey");
            idTuples.push([msgID, changeKey]);
        }
        getMsgBodies(idTuples);
    });
}

/*
 * Retrieve full message information (including body text) for the list of message IDs passed
 * in to this function.
 * This function uses an EWS request to search for all messages with the specified ID and key.
 * The returned message contents are then parsed to extract information like the sender's email
 * address, timestamp, and contents. For each message received, it will try to add a new chat entry
 * to the page.
 *
 * The idTuples object should be an array of [msgID, changeKey] objects.
 */
function getMsgBodies(idTuples) {
    var contactAddr = document.querySelector("#contact-email").textContent;
    var idList = ""
    for (let item of idTuples) {
        idList += '<t:ItemId Id="' + item[0] + '" ChangeKey="' + item[1] + '" />';
    }
    var request = wrapRequest(
        '<m:GetItem>' +
        '  <m:ItemShape>' +
        '    <t:BaseShape>Default</t:BaseShape>' +
        '    <t:BodyType>Text</t:BodyType>' +
        '  </m:ItemShape>' +
        '  <m:ItemIds>' +
             idList +
        '  </m:ItemIds>' +
        '</m:GetItem>'
    );

    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
        var response = $.parseXML(result.value);
        var msgs = response.getElementsByTagName("t:Message");
        var idTuples = [];
        for (let item of msgs) {
            //re-retrieve useful information like ID, key, send-timestmp, and 'from' email address
            var id = item.getElementsByTagName("t:ItemId")[0];
            var msgID = id.getAttribute("Id");
            var changeKey = id.getAttribute("ChangeKey");
            var fromMailbox = item.getElementsByTagName("t:From")[0];
            var fromAddr = fromMailbox.getElementsByTagName("t:EmailAddress")[0].textContent;
            var isFromMe = (fromAddr.localeCompare(contactAddr) == 0) ? false : true;
            var ts = item.getElementsByTagName("t:DateTimeSent")[0].textContent;
            //read the body contents. Make sure it's not too large.
            var body = item.getElementsByTagName("t:Body")[0].textContent;
            if (body.length > 1000) {
                body = body.substring(0, 1000);
            }
            //add email item to the chat list if required
            addChatEntry(msgID, changeKey, body, ts, isFromMe);
        }
    });
}

/*
 * Add chat item's contents to the sidebar. Clone the template node, fill it in, and insert it
 * in the correct timestamp order in the parent node.
 * This function will first check if a chat item with the same ID and key already exists. If one
 * is found, it will abort. If one is not found, we will create a new chat entry item and insert
 * it in the correct location in the list (using timestamp-ordering).
 */
function addChatEntry(id, changeKey, body, ts, isFromMe) {
    //chat-entry builder helper function
    function buildChatEntry() {
        var template = document.querySelector("#chat-template");
        var clone = template.content.cloneNode(true);
        var entry = clone.querySelector("[name=chat-entry-holder]");
        if (isFromMe) {
            entry.classList.add("chat-entry-tx");
        } else {
            entry.classList.add("chat-entry-rx");
        }
        clone.querySelector("[name=chat-content]").textContent = body;
        clone.querySelector("[name=chat-ts]").textContent = ts;
        clone.querySelector("[name=chat-id]").textContent = id;
        clone.querySelector("[name=chat-changekey]").textContent = changeKey;
        return clone;
    }

    //Check for duplicate items. Set a flag to indicate if a duplicate is found, and set a refItem
    //if we find a location to add this in the list of children.
    var refItem = null;
    var isDuplicate = false; //another entry with the same ID is found
    var entryDate = new Date(ts);
    var existingEntries = document.getElementsByName("chat-entry-holder");
    for (var i = 0; i < existingEntries.length; i++) {
        if (id.localeCompare(existingEntries[i].querySelector("[name=chat-id]").textContent) == 0 &&
            changeKey.localeCompare(existingEntries[i].querySelector("[name=chat-changekey]").textContent) == 0) {
                isDuplicate = true;
                break;
        }
        if (refItem == null &&
                entryDate < new Date(existingEntries[i].querySelector("[name=chat-ts]").textContent)) {
            //The new chat entry item should go before this one
            refItem = existingEntries[i];
        }
    }

    //Add the new chat entry item if no duplicates were found
    if (!isDuplicate) {
        var clone = buildChatEntry();
        var chatView = document.querySelector("#app-body");
        if (refItem == null) {
            //there are currently no children, so just append
            chatView.appendChild(clone);
        } else {
            chatView.insertBefore(clone, refItem);
        }
        chatView.scrollTop = chatView.scrollHeight;
    }
}

/*
 * Send an email message using the EWS feature. This will create a new email object, configure the
 * recipient as the current contact's information, and actually send it. When this function is
 * complete it will call the configured callback so the UI can be updated.
 */
function sendMail(body, cb) {
    var contactAddr = document.querySelector("#contact-email").textContent;
    var request = wrapRequest(
        '<m:CreateItem MessageDisposition="SendAndSaveCopy">' +
        ' <m:SavedItemFolderId>' +
        '  <t:DistinguishedFolderId Id="sentitems" />' +
        ' </m:SavedItemFolderId>' +
        ' <m:Items>' +
        '  <t:Message>' +
        '   <t:Subject>SMS</t:Subject>' +
        '   <t:Body>' + body + '</t:Body>' +
        '   <t:ToRecipients>' +
        '    <t:Mailbox><t:EmailAddress>' + contactAddr + '</t:EmailAddress></t:Mailbox>' +
        '   </t:ToRecipients>' +
        '  </t:Message>' +
        ' </m:Items>' +
        '</m:CreateItem>'
    );

    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
        console.log("SentMail:", result);
        var response = $.parseXML(result.value);
        var respCode = response.getElementsByTagName("m:ResponseCode");
        if (respCode.textContent.localeCompare("NoError") == 0) {
            cb(true);
        } else {
            cb(false);
        }
    });
}

/*
 * This is the event-handler function which runs when the quick-sending form is submitted. It gets
 * the form data, packages it, and calls the sendMail function to actually send it. When the send
 * is complete, it uses a callback to clear the text field and amend the scroll.
 */
function submitMsg(theForm) {
    var msgtext = document.querySelector("#composebox").value.trim();
    if (msgtext.len == 0) {
        console.log("No message has been composed in the text area");
    } else {
        sendMail(msgtext, function(didSendSuccessfully) {
            if (didSendSuccessfully) {
                document.querySelector("#composebox").value = "";
            }
        });
        //scroll to bottom of window
        var chatView = document.querySelector("#app-body");
        chatView.scrollTop = chatView.scrollHeight;
    }
    return false;
}

