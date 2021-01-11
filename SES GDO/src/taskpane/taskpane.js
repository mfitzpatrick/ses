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
        document.getElementById("expand-options").onclick = toggleMsgOptions;
        document.getElementById("msg-template").onchange = setTemplate;
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
        //check if the 'from' field is our user's information
        var iam = Office.context.mailbox.userProfile.emailAddress;
        if (iam.localeCompare(item.from.emailAddress) == 0) {
            //This is a message sent from us. If there is only 1 recipient, set the sender context
            //to that recipient. If not, don't get related messages.
            var toList = Office.context.mailbox.item.to;
            if (toList.length == 1) {
                updateSenderContext(toList[0]);
                getRelatedMsgs(toList[0].emailAddress);
            } else {
                //clear contents
                document.querySelector("#app-body").innerHTML = '';
                document.querySelector("#contact-name").textContent = "Too Many Recipients";
                document.querySelector("#contact-email").textContent = "";
            }
        } else {
            updateSenderContext(item.from);
            getRelatedMsgs(item.from.emailAddress);
        }
    }
}

/*
 * Update the sender-information context fields at the top of the pane.
 */
function updateSenderContext(contact) {
    var currentAddr = document.querySelector("#contact-email");
    if (contact.emailAddress != currentAddr.textContent) {
        //clear the list
        document.querySelector("#app-body").innerHTML = '';
    }
    document.querySelector("#contact-name").textContent = contact.displayName;
    currentAddr.textContent = contact.emailAddress;
}

/*
 * Handle the error that occurred.
 * This will log information about the error to the console.
 * This will forward error information to a section where it will be monitored.
 */
function ehandle(context, req, resp) {
    console.log(`Error occurred in ${context} with req:`, req, "\nand resp:", resp);
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
 *
 * There are multiple folders that need to be searched (inbox and sent-items). This will create 2
 * separate EWS messages to search these folders.
 */
function getRelatedMsgs(emailAddr) {
    var folderlist = ["inbox", "sentitems"];
    for (let folder of folderlist) {
        var request = wrapRequest(
            '<m:FindItem>' +
            '  <m:ItemShape>' +
            '    <t:BaseShape>IdOnly</t:BaseShape>' +
            '  </m:ItemShape>' +
            '  <m:ParentFolderIds>' +
            '    <t:DistinguishedFolderId Id="' + folder + '" />' +
            '  </m:ParentFolderIds>' +
            '  <m:QueryString>Participants:' + emailAddr + '</m:QueryString>' +
            '</m:FindItem>'
        );

        Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
            var response = $.parseXML(result.value);
            if (response == null) {
                ehandle("getRelatedMsgs result object", request, result);
                return undefined;
            }
            var msgs = response.getElementsByTagName("t:Message");
            var idTuples = [];
            for (let item of msgs) {
                var id = item.getElementsByTagName("t:ItemId")[0];
                var msgID = id.getAttribute("Id");
                var changeKey = id.getAttribute("ChangeKey");
                idTuples.push([msgID, changeKey]);
            }
            if (idTuples.length > 0) {
                getMsgBodies(idTuples);
            }
        });
    }
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
    var myAddr = Office.context.mailbox.userProfile.emailAddress;
    var idList = ""
    for (let item of idTuples) {
        idList += '<t:ItemId Id="' + item[0] + '" ChangeKey="' + item[1] + '" />';
    }
    var request = wrapRequest(
        '<m:GetItem>' +
        '  <m:ItemShape>' +
        '    <t:BaseShape>Default</t:BaseShape>' +
        '    <t:BodyType>Text</t:BodyType>' +
        '    <t:AdditionalProperties>' +
        '      <t:FieldURI FieldURI="item:IsFromMe" />' +
        '    </t:AdditionalProperties>' +
        '  </m:ItemShape>' +
        '  <m:ItemIds>' +
             idList +
        '  </m:ItemIds>' +
        '</m:GetItem>'
    );

    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
        var response = $.parseXML(result.value);
        if (response == null) {
            ehandle("getMsgBodies result object", request, result);
            return undefined;
        }
        var msgs = response.getElementsByTagName("t:Message");
        var idTuples = [];
        for (let item of msgs) {
            //re-retrieve useful information like ID, key, send-timestmp, and 'from' email address
            var id = item.getElementsByTagName("t:ItemId")[0];
            var msgID = id.getAttribute("Id");
            var changeKey = id.getAttribute("ChangeKey");
            var fromMailbox = item.getElementsByTagName("t:From")[0];
            var fromAddr = fromMailbox.getElementsByTagName("t:EmailAddress")[0].textContent;
            var isFromMe = (fromAddr.localeCompare(myAddr) == 0);
            var ts = item.getElementsByTagName("t:DateTimeSent")[0].textContent;
            //read the body contents. Make sure it's not too large.
            var body = item.getElementsByTagName("t:Body")[0].textContent;
            //truncate the legal disclaimer added by some mail services
            body = truncateBody(body, "This correspondence is for the named");
            body = truncateBody(body, "-#-");
            //search for forwarded-message identifiers
            body = truncateBody(body, "---------- Forwarded message");
            body = truncateBody(body, "----- Original Message -----");
            if (body.length > 1000) {
                //cut off messages that are significantly long
                body = body.substring(0, 1000);
            }
            body = body.trim();
            if (body.length > 0) {
                //add email item to the chat list if required
                addChatEntry(msgID, changeKey, body, ts, isFromMe);
            }
        }
    });
}

/*
 * For a specified string, body, search for a substring containing 'identifier'. If this substring
 * is found, cut the string off at this point and return it.
 */
function truncateBody(body, identifier) {
    var idx = body.indexOf(identifier);
    if (idx > 0) {
        //truncate the string from the 'identifier' index
        body = body.substring(0, idx);
    }
    return body;
}

/*
 * Add chat item's contents to the sidebar. Clone the template node, fill it in, and insert it
 * in the correct timestamp order in the parent node.
 * This function will first check if a chat item with the same ID and key already exists. If one
 * is found, it will abort. If one is not found, we will create a new chat entry item and insert
 * it in the correct location in the list (using timestamp-ordering).
 */
function addChatEntry(id, changeKey, body, ts, isFromMe) {
    var entryDate = new Date(ts);
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
        clone.querySelector("[name=chat-ts]").textContent = entryDate;
        clone.querySelector("[name=chat-id]").textContent = id;
        clone.querySelector("[name=chat-changekey]").textContent = changeKey;
        return clone;
    }

    //Check for duplicate items. Set a flag to indicate if a duplicate is found, and set a refItem
    //if we find a location to add this in the list of children.
    var refItem = null;
    var toDelete = null;
    var isDuplicate = false; //another entry with the same ID is found
    var existingEntries = document.getElementsByName("chat-entry-holder");
    for (var i = 0; i < existingEntries.length; i++) {
        //check for empty ID
        if (existingEntries[i].querySelector("[name=chat-id]").textContent.length == 0) {
            //this is a fake item, remove it
            toDelete = existingEntries[i].querySelector("[name=chat-id]").parentElement;
        }
        //check for duplicate ID
        if (id.localeCompare(existingEntries[i].querySelector("[name=chat-id]").textContent) == 0 &&
            changeKey.localeCompare(existingEntries[i].querySelector("[name=chat-changekey]").textContent) == 0) {
                isDuplicate = true;
                break;
        }
        //find closest chat entry (via date)
        if (refItem == null &&
                entryDate < new Date(existingEntries[i].querySelector("[name=chat-ts]").textContent)) {
            //The new chat entry item should go before this one
            refItem = existingEntries[i];
        }
    }

    //delete null-ID item in the hope that it has just been filled by a most-recent get
    if (toDelete != null) {
        toDelete.remove();
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
 * Retrieves all email addresses from a group named COL in the current user's contacts section.
 */
function getCOL(cb) {
    var request = wrapRequest(
        '<m:FindPeople>' +
        ' <m:IndexedPageItemView BasePoint="Beginning" MaxEntriesReturned="500" Offset="0" />' +
        ' <m:ParentFolderId>' +
        '  <t:DistinguishedFolderId Id="contacts" />' +
        ' </m:ParentFolderId>' +
        '</m:FindPeople>'
    );

    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
        var col = [];
        console.log("Retrieved COL:", result);
        var response = $.parseXML(result.value);
        var respCode = response.getElementsByTagName("ResponseCode")[0];
        if (respCode.textContent.localeCompare("NoError") != 0) {
            ehandle("getCOL had an error", request, result);
            return undefined;
        }
        var contacts = response.getElementsByTagName("Persona");
        for (let item of contacts) {
            //find all instances of EmailAddress (even nested instances)
            var emails = item.getElementsByTagName("EmailAddress");
            console.log("COL emails:", emails);
            for (let email of emails) {
                //Check if the text node contains this domain name
                if (email.textContent.endsWith("bccdisastermanage.sms.optus.com.au")) {
                    console.log("Relevant email:", email.textContent);
                    col.push(email.textContent);
                }
            }
        }
        if (cb != null) {
            cb(col);
        }
    });
}

/*
 * Send an email message using the EWS feature. This will create a new email object, configure the
 * recipient as the current contact's information, and actually send it. When this function is
 * complete it will call the configured callback so the UI can be updated.
 */
function sendMail(contactList, body, cb) {
    var cList = ""
    for (let item of contactList) {
        cList += '<t:Mailbox><t:EmailAddress>' + item + '</t:EmailAddress></t:Mailbox>';
    }
    var request = wrapRequest(
        '<m:CreateItem MessageDisposition="SendAndSaveCopy">' +
        ' <m:SavedItemFolderId>' +
        '  <t:DistinguishedFolderId Id="sentitems" />' +
        ' </m:SavedItemFolderId>' +
        ' <m:Items>' +
        '  <t:Message>' +
        '   <t:Subject>GDO SMS</t:Subject>' +
        '   <t:Body BodyType="Text">' + body + '\n-#-</t:Body>' +
        '   <t:ToRecipients>' + cList + '</t:ToRecipients>' +
        '  </t:Message>' +
        ' </m:Items>' +
        '</m:CreateItem>'
    );

    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
        console.log("SentMail:", result);
        var response = $.parseXML(result.value);
        var respCode = response.getElementsByTagName("m:ResponseCode")[0];
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
    function doCleanup(didSendSuccessfully) {
        if (didSendSuccessfully) {
            //update msg context list
            addChatEntry("", "", document.querySelector("#composebox").value, new Date(), true);
            //clean up UI
            document.querySelector("#composebox").value = "";
            document.querySelector("#msg-options").style.display = "none";
            document.querySelector("#msg-template").value = "none";
            document.querySelector("#msg-recipients").value = "contact";
        }
    }

    var msgtext = document.querySelector("#composebox").value.trim();
    if (msgtext.len == 0) {
        console.log("No message has been composed in the text area");
    } else if (document.querySelector("#msg-recipients").value.localeCompare("col") == 0) {
        getCOL(function(col) {
            sendMail(col, msgtext, doCleanup);
        });
    } else {
        var contactAddr = document.querySelector("#contact-email").textContent;
        sendMail([contactAddr], msgtext, doCleanup);
    }
    //scroll to bottom of window
    var chatView = document.querySelector("#app-body");
    chatView.scrollTop = chatView.scrollHeight;
    return false;
}

/*
 * Toggles the display-state of the msg-options div. This allows us to keep the custom message
 * options hidden under normal circumstances, and display them only if the user chooses to.
 */
function toggleMsgOptions() {
    var opt = document.querySelector("#msg-options");
    var isOpen = (opt.style.display.localeCompare("flex") == 0);
    if (isOpen) {
        opt.style.display = "none";
    } else {
        opt.style.display = "flex";
    }
}

/*
 * This event handler is called when the msg-template pick-list selected item is changed. The item
 * newly selected defines the message template that will be used.
 * The message template is set in the Compose Box before the function returns.
 */
function setTemplate() {
    var tmpl = '';
    var now = new Date();
    var datestr = `${now.getFullYear()}/${now.getMonth() + 1}/${now.getDate()}-` +
        `${now.getHours().toString().padStart(2, "0")}` +
        `:${now.getMinutes().toString().padStart(2, "0")}`;
    //Take an educated guess at a suitable reply-by time so we can prefill it.
    var inOneHr = new Date(now.getTime() + (1 * 60 * 60 * 1000));
    var replyByTimeStr = `${inOneHr.getHours().toString().padStart(2, "0")}` +
        `:${inOneHr.getMinutes().toString().padStart(2, "0")}`;
    var iam = Office.context.mailbox.userProfile.displayName;
    var templateOption = document.getElementById("msg-template");
    switch (templateOption.value) {
        case "storm":
            tmpl = `SES - LEANFORWARD - ${datestr} - DETAILS - Reply ASAP` +
                ` with Name, ID Number, and Qualifications and await further instructions.` +
                ` Thanks GDO ${iam.split(" ")[0]}`;
            break;
        case "lean":
            tmpl = `SES - LEANFORWARD - ${datestr} - DETAILS - Reply by ${replyByTimeStr}` +
                ` with Name, ID Number, and Qualifications and await further instructions.` +
                ` Thanks GDO ${iam.split(" ")[0]}`;
            break;
        case "standup":
            tmpl = `SES - STANDUP - ${datestr} - Please proceed to East HQ by ${replyByTimeStr}.` +
                ` You will be in team NUM. Please acknowledge this message.` +
                ` Thanks GDO ${iam.split(" ")[0]}`;
            break;
        case "standdown":
            tmpl = `SES - STANDDOWN - ${datestr} - Thankyou for taking the time to reply.` +
                ` Thanks GDO ${iam.split(" ")[0]}`;
            break;
        case "sitrep":
            tmpl = `SES - SITREP - ${datestr} - Eastern teams deployed as-per TAMS.` +
                ` Thanks GDO ${iam.split(" ")[0]}`;
            break;
        default:
            tmpl = '';
            break;
    }
    document.querySelector("#composebox").value = tmpl;
}

