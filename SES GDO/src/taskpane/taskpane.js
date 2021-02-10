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
 * Helper function to get the valid REST item ID for a given item. This performs the ID conversion
 * process if necessary before returning the ID. It can be used in REST API calls.
 */
function getItemRestId() {
    console.log("Hostname is:", Office.context.mailbox.diagnostics.hostName);
    if (Office.context.mailbox.diagnostics.hostName === 'OutlookAndroid') {
        // itemId is already REST-formatted.
        return Office.context.mailbox.item.itemId;
    } else {
        // Convert to an item ID for API v2.0.
        return Office.context.mailbox.convertToRestId(
            Office.context.mailbox.item.itemId,
            Office.MailboxEnums.RestVersion.v2_0);
    }
}

/*
 * Get a recent history of messages related to the specified contact.
 * For each message received, add a new chat entry to the UI via the addChatEntry() function call.
 * This uses the REST API to perform a search over the whole mailbox, and returns the messages
 * found to include the given email address.
 */
function getRelatedMsgs(emailAddr) {
    var myAddr = Office.context.mailbox.userProfile.emailAddress;
    Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
        var uriEmailAddr = encodeURIComponent(emailAddr); //url-encode email address for query string
        if (result.status === "succeeded") {
            $.ajax({
                url: Office.context.mailbox.restUrl + `/v2.0/me/messages?$search="participants:${uriEmailAddr}"&$top=30`,
                dataType: 'json',
                headers: {'Authorization': 'Bearer ' + result.value},
            }).done(function(item) {
                console.log("search ajax returned:", item);
                for (let msg of item.value) {
                    //retrieve useful information like ID, key, send-timestmp, and 'from' email address
                    var isFromMe = (msg.Sender.EmailAddress.Address.localeCompare(myAddr) == 0);
                    //read the body contents. Make sure it's not too large.
                    var body = msg.Body.Content;
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
                        addChatEntry(msg.Id, msg.ChangeKey, body, msg.CreatedDateTime, isFromMe);
                    }
                }
            }).fail(function(error) {
                console.log("ajax failed:", error);
            });
        } else {
            // Handle the error.
            console.log("error happened:", result);
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
    var searchDomain = "bccdisastermanage.sms.optus.com.au";
    Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
        var uriEmailAddr = encodeURIComponent(emailAddr); //url-encode email address for query string
        if (result.status === "succeeded") {
            $.ajax({
                url: Office.context.mailbox.restUrl + `/v2.0/me/people?$search="${searchDomain}"&$top=300`,
                dataType: 'json',
                headers: {'Authorization': 'Bearer ' + result.value},
            }).done(function(item) {
                console.log("people search ajax returned:", item);
                for (let email of item.value) {
                    //Check if the text node contains this domain name
                    if (email.textContent.endsWith(searchDomain)) {
                        console.log("Relevant email:", email.textContent);
                        col.push(email.textContent);
                    }
                }
                if (cb != null) {
                    cb(col);
                }
            });
        }
    });
}

/*
 * Send an email message. This will create a new email object, configure the
 * recipient list from the list passed in, and actually send it. When this function is
 * complete it will call the configured callback so the UI can be updated.
 */
function sendMail(contactList, body, cb) {
    var msg = {
        Message: {
            Subject: "GDO SMS",
            Body: {
                ContentType: "Text",
                Content: body + "\n-#-",
            },
            ToRecipients: [],
        }
    };
    for (let item of contactList) {
        msg.Message.ToRecipients.push({"EmailAddress": {"Address": item}});
    }
    Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
        if (result.status === "succeeded") {
            console.log("sending msg: ", JSON.stringify(msg));
            //create a draft
            $.ajax({
                url: Office.context.mailbox.restUrl + "/v2.0/me/sendMail",
                dataType: 'json',
                contentType: 'application/json',
                headers: { 'Authorization': 'Bearer ' + result.value },
                type: "POST",
                data: JSON.stringify(msg)
            }).done(function(item) {
                console.log("create mail ajax returned:", item);
            }).fail(function(error) {
                if (error.status != 202) {
                    console.log("ajax failed:", error);
                } //else it is 202, which is a success
            });
        } else {
            // Handle the error.
            console.log("error happened:", result);
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

