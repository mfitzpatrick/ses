/*
 * This file uses an authenticated Google API connection with a mailbox to retrieve emails
 * based on specified filters.
 * In short:
 *  - Retrieve and cache a list of all message threads with recent activity
 *  - Retrieve and cache messages sent to and from a particular recipient
 *  - Retrieve and cache a list of group messages sent
 */

// Client ID and API key from the Developer Console
var CLIENT_ID = '';
var API_KEY = '';
var EMAIL_DOMAIN = '';
var FORWARDER_EMAIL = '';

// Array of API discovery doc URLs for APIs used by the quickstart
var DISCOVERY_DOCS = ["https://www.googleapis.com/discovery/v1/apis/gmail/v1/rest"];

// Authorisation scopes required by the Google API; multiple scopes can be included, separated by spaces.
var SCOPES = 'https://www.googleapis.com/auth/gmail.modify' +
    ' https://www.googleapis.com/auth/gmail.compose' +
    ' https://www.googleapis.com/auth/gmail.send';

var signoutButton = document.getElementById('logout_btn');

// Array of message threads from the SMS inbox. This is used to indicate which messages are most
// recent.
var sms_inbox_threads = [];
var sms_label_id = "not-set";

/*
 * Asynchronously load a JSON document from URI. Call the callback with the JSON payload if the load
 * is successful.
 */
function loadJSON(uri, callback) {
    var xobj = new XMLHttpRequest();
    xobj.overrideMimeType("application/json");
    xobj.open("GET", uri, true);
    xobj.onreadystatechange = function() {
        if (xobj.readyState == 4 && xobj.status == "200") {
            callback(xobj.responseText);
        }
    };
    xobj.send(null);
}

/**
 *  On load, called to load the auth2 library and API client library.
 */
function handleClientLoad() {
    loadJSON("/config.json", function(data) {
        var config_data = JSON.parse(data);
        API_KEY = config_data.api_key;
        CLIENT_ID = config_data.web.client_id;
        EMAIL_DOMAIN = config_data.email_domain;
        FORWARDER_EMAIL = config_data.forwarder_email;
        gapi.load('client:auth2', initClient);
    });
}

/**
 *  Initialises the API client library and sets up sign-in state
 *  listeners.
 */
function initClient() {
    gapi.client.init({
        apiKey: API_KEY,
        clientId: CLIENT_ID,
        discoveryDocs: DISCOVERY_DOCS,
        scope: SCOPES
    }).then(function () {
        // Listen for sign-in state changes.
        gapi.auth2.getAuthInstance().isSignedIn.listen(updateSigninStatus);

        // Handle the initial sign-in state.
        updateSigninStatus(gapi.auth2.getAuthInstance().isSignedIn.get());
        signoutButton.onclick = handleSignoutClick;
    }, function(error) {
        console.log(JSON.stringify(error, null, 2));
    });
}

/**
 *  Called when the signed in status changes, to update the UI
 *  appropriately. After a sign-in, the API is called.
 */
function updateSigninStatus(isSignedIn) {
    if (isSignedIn) {
        signoutButton.style.display = 'block';
        // We're now signed in; retrieve user information and start downloading messages
        whoAmI();
        getSMSLabelID();
    } else {
        signoutButton.style.display = 'none';
    }
}

/**
 *  Sign in the user upon button click.
 */
function handleAuthClick(event) {
    gapi.auth2.getAuthInstance().signIn();
}

/**
 *  Sign out the user upon button click.
 */
function handleSignoutClick(event) {
    gapi.auth2.getAuthInstance().signOut();
}

/*
 * Resolves the existential crisis of determining which user we're authenticated as.
 */
function whoAmI() {
    gapi.client.gmail.users.getProfile({
        'userId': 'me',
    }).then(function(response) {
        console.log("Who Am I?", response.result);
    });
}

/*
 * List labels for this account and find the label called "SMS". The label ID string is often
 * different from the label name string.
 */
function getSMSLabelID() {
    gapi.client.gmail.users.labels.list({
        'userId': 'me',
    }).then(function(response) {
        var labels = response.result.labels;

        if (labels && labels.length > 0) {
            for (i = 0; i < labels.length; i++) {
                var l = labels[i];
                if (l.name == "SMS") {
                    sms_label_id = l.id;
                    //now we can start off retrieving messages
                    getNewMessages();
                    break;
                }
            }
        } else {
            console.log("No labels found in this account");
        }
    });
}

/*
 * Refresh the stored list of message thread IDs.
 */
function getNewMessages() {
    gapi.client.gmail.users.threads.list({
        'userId': 'me',
        'labelIds': sms_label_id,
    }).then(function(response) {
        var threads = response.result.threads;

        if (threads && threads.length > 0) {
            for (i = 0; i < threads.length; i++) {
                var t = threads[i];
                if (!(t.id in sms_inbox_threads)) {
                    // Store an indicator of this thread in our threads array
                    sms_inbox_threads[t.id] = t;
                    // Also retrieve all messages from this thread for caching
                    getMessagesFromThread(t.id);
                }
            }
        } else {
            console.log("No threads found in SMS label");
        }
    });
}

/*
 * Retrieve all messages stored in a given thread.
 */
function getMessagesFromThread(threadID) {
    gapi.client.gmail.users.threads.get({
        'userId': 'me',
        'id': threadID,
        'format': 'FULL',
    }).then(function(response) {
        var msgs = response.result.messages;

        if (msgs && msgs.length > 0) {
            for (i = 0; i < msgs.length; i++) {
                var msg = msgs[i];
                // read message contents into cache
                getOneMessage(msg.id);
            }
        } else {
            console.log("No messages found for thread:", threadID);
        }
    });
}

/*
 * Get a single message's contents from gmail.
 */
function getOneMessage(messageID) {
    gapi.client.gmail.users.messages.get({
        'userId': 'me',
        'id': messageID,
        'format': 'FULL',
    }).then(function(response) {
        var msg = response.result;

        if (msg.payload.headers.length > 0) {
            //search the headers for a header containing the EMAIL_DOMAIN, and cache it in a local
            //variable
            var email = "";
            var date = "";
            var body = "";
            for (i = 0; i < msg.payload.headers.length; i++) {
                var hdr = msg.payload.headers[i];
                if (hdr.name.toUpperCase() == "TO" || hdr.name.toUpperCase() == "FROM") {
                    if (hdr.value.includes(EMAIL_DOMAIN)) {
                        email = hdr.value;
                        body = b64Decode(msg.payload.body.data).trim();
                    } else if (email.length == 0 && body.length == 0 && hdr.value.includes(FORWARDER_EMAIL)) {
                        //the response was sent via a forwarder mailbox. We need to parse out the
                        //forwarder headers to determine who sent this message
                        var parsed = parseForwardedMessage(msg.payload.body);
                        if (parsed != null) {
                            email = parsed.sender;
                            body = parsed.contents;
                        } else {
                            console.log("parseForwardedMessage failed");
                        }
                    }
                } else if (hdr.name.toUpperCase() == "DATE") {
                    date = hdr.value;
                }
            }
            //if we found all the headers we were looking for, add them to our message DB
            if (email.length > 0 && date.length > 0 && body.length > 0) {
                sesDB.addMsg(msg, email, date, body);
            }
        } else {
            console.log("Couldn't find message:", messageID);
        }
    });
}

/*
 * Parse the base64-encoded message body object and search for a 'forwarded message' header. From
 * this, extract the sender information, and also extract the message contents following the header.
 */
function parseForwardedMessage(msgBody) {
    var body = b64Decode(msgBody.data);
    var bodyLines = body.split("\n");
    var sender = '';
    var contents = '';
    var contentsStart = 0;
    for (i = 0; i < bodyLines.length; i++) {
        var line = bodyLines[i];
        var upperLine = bodyLines[i].toUpperCase();
        //check if a 'from' line appears with the correct email domain
        if (upperLine.includes("FROM") && line.includes(EMAIL_DOMAIN)) {
            sender = line.split(": ")[1];
        }
        if (contentsStart == 0) {
            //now check for the start of the body section. It will appear after the subject line
            if (upperLine.includes("SUBJECT")) {
                contentsStart = i + 1;
            }
        } else {
            if (upperLine.includes("ORIGINAL MESSAGE")) {
                //return the required object, we can ignore all following lines
                return {'sender': sender, 'contents': contents.trim()};
            }
            contents += line;
        }
    }
    //we searched to the end without finding anything, returning null indicates failure
    return null;
}

function sendMail(dst, body) {
    var msg = btoa(
        "Content-Type: text/plain; charset=\"UTF-8\"\n" +
        "Content-Transfer-Encoding: message/rfc2822\n" +
        "to: " + dst + "\n" +
        "from: \"SES Eastern Group Operations (SES)\" <seseasterngroupops@ses.qfes.qld.gov.au>\n" +
        "\n" +
        body
    );
    gapi.client.gmail.users.messages.send({
        'userId': 'me',
        'resource': {
            'raw': msg,
            'labelIds': [sms_label_id],
        },
    }).then(function(response) {
        console.log("sendMail response:", response);
    });
}

/*
 * Base64 Decoder functions. These are included because the standard browser base64 functions don't
 * work well enough for this. This decoder function replaces special characters with an equivalent
 * which is acceptable to the browser's implementation of atob().
 */
function b64Decode(val) {
    function unicodeBase64Decode(text) {
        return decodeURIComponent(Array.prototype.map.call(window.atob(text), function(c) {
            return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(''));
    }

    return val.split(/\r?\n/g).map(function(value) {
        return (unicodeBase64Decode(value.replace(/\s+/g, '').replace(/\-/g, '+').replace(/\_/g, '/')));
    }).join("\n");
}

