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
        console.log("Labels:", response.result);
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
        console.log("Threads:", threads)

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
        console.log("Thread:", threadID, "Data:", response.result)
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
        console.log("Message:", messageID, "Data:", response.result)
        var msg = response.result;

        if (msg.payload.headers.length > 0) {
            //search the headers for a header containing the EMAIL_DOMAIN, and cache it in a local
            //variable
            var email = "";
            var date = "";
            for (i = 0; i < msg.payload.headers.length; i++) {
                var hdr = msg.payload.headers[i];
                if (hdr.name.toUpperCase() == "TO" || hdr.name.toUpperCase() == "FROM") {
                    if (hdr.value.includes(EMAIL_DOMAIN)) {
                        email = hdr.value;
                    }
                } else if (hdr.name.toUpperCase() == "DATE") {
                    date = hdr.value;
                }
            }
            //if we found all the headers we were looking for, add them to our message DB
            if (email.length > 0 && date.length > 0) {
                sesDB.addMsg(msg, email, date);
            }
        } else {
            console.log("Couldn't find message:", messageID);
        }
    });
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

