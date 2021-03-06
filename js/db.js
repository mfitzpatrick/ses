/*
 * This is the application-level database interface for storing messages for easy retrieval.
 * It wraps to the promise-based indexedDB implementation in idb.js.
 */

var sesDB = (function() {
    'use strict';

    if (!('indexedDB' in window)) {
        console.log('This browser doesn\'t support IndexedDB');
        return;
    }

    // Callback called when a new message is added to an existing message chain
    var msgCallbackFilterPeer = '';
    var addMsgCallback = null;
    // Callback called when a new message is added to an new message chain
    var addMsgChainCallback = null;

    var sesDBpromise = idb.open('ses', 1, function(upgradeDB) {
        switch (upgradeDB.oldVersion) {
        case 0:
            console.log("Upgrading SES DB");
            if (!upgradeDB.objectStoreNames.contains('messages')) {
                var os = upgradeDB.createObjectStore('messages', {keyPath: 'id'});
                // the 'peer' index holds the remote peer's phone number for more easy querying
                os.createIndex('peer', 'peer', {unique: false});
            }
            //create the 'contacts' cache, which will cache the user's phone number and name combination
            //for more easy display on the relevant pages
            if (!upgradeDB.objectStoreNames.contains('contacts')) {
                var os = upgradeDB.createObjectStore('contacts', {keyPath: 'phone'});
            }
            break;
        }
    });

    function addMsg(msg, email, ts, body, is_sent) {
        var peer = email.split("@", 1)[0];
        //add this new message to our global database
        sesDBpromise.then(function(db) {
            var tx = db.transaction('messages', 'readwrite');
            var os = tx.objectStore('messages');
            var dbmsg = {
                id: msg.id,
                msg: msg,
                email: email,
                peer: peer,
                is_sent: is_sent,
                ts: ts,
                body: body,
            };
            os.add(dbmsg);
            return tx.complete;
        }).then(function() {
            //also add message to the relevant message chain
            if (addMsgCallback != null && msgCallbackFilterPeer == peer) {
                addMsgCallback(msg, ts, body, is_sent);
            } else if (addMsgChainCallback != null) {
                // call the callback to inform it of the new message
                addMsgChainCallback(msg, peer, ts, body);
            } else {
                console.log("no message callback");
            }
        }).catch(function(err) {
            console.log("Failure during addMsg:", err);
        });
    }

    function updateMsgChainCallback(f) {
        addMsgChainCallback = f;
    }

    function updateChatCallback(peer, f) {
        msgCallbackFilterPeer = peer;
        addMsgCallback = f;
    }

    /*
     * Retrieve all unique peer entries (unique as-per phone number). For each retrieved entry,
     * call the cb() function.
     */
    function getUniquePeers(cb) {
        sesDBpromise.then(function(db) {
            var tx = db.transaction('messages', 'readonly');
            var os = tx.objectStore('messages');
            var idx = os.index('peer');
            return idx.openCursor(null, 'nextunique');
        }).then(function handlePeer(cursor) {
            if (cursor) {
                cb(cursor._cursor.value);
            } else {
                //Retrieved all DB entries
                return;
            }
            return cursor.continue().then(handlePeer);
        }).then(function() {
            console.log("Peer retrieval complete");
        });
    }

    /*
     * Retrieve all messages associated with a given peer (specified by the phone number).
     * For each retrieved peer, call the cb() function.
     */
    function getMessagesForPeer(peer, cb) {
        sesDBpromise.then(function(db) {
            var tx = db.transaction('messages', 'readonly');
            var os = tx.objectStore('messages');
            var idx = os.index('peer');
            return idx.openCursor(IDBKeyRange.only(peer));
        }).then(function handleMsg(cursor) {
            if (cursor) {
                cb(cursor._cursor.value);
            } else {
                //Retrieved all DB entries
                return;
            }
            return cursor.continue().then(handleMsg);
        }).then(function() {
            console.log("Message retrieval complete");
        });
    }

    /*
     * Cache a user's name and phone number in the DB.
     */
    function addContact(phoneNumber, name) {
        sesDBpromise.then(function(db) {
            var tx = db.transaction('contacts', 'readwrite');
            var os = tx.objectStore('contacts');
            var dbmsg = {
                phone: phoneNumber,
                name: name,
            };
            os.add(dbmsg);
            return tx.complete;
        }).then(function() {
            console.log("Contact added:", name, phoneNumber);
        });
    }

    /*
     * Retrieve a user's name when supplied with only their phone number.
     * The name is returned via the provided callback.
     */
    function getNameForPhoneNumber(phoneNumber, cb) {
        sesDBpromise.then(function(db) {
            var tx = db.transaction('contacts', 'readonly');
            var os = tx.objectStore('contacts');
            return os.get(phoneNumber);
        }).then(function(result) {
            if (cb != null) {
                cb(result);
            } else {
                console.log("getNameForPhoneNumber retrieved:", result);
            }
        });
    }

    return {
        addMsg: (addMsg),
        updateMsgChainCallback: (updateMsgChainCallback),
        updateChatCallback: (updateChatCallback),
        getUniquePeers: (getUniquePeers),
        getMessagesForPeer: (getMessagesForPeer),
        addContact: (addContact),
        getNameForPhoneNumber: (getNameForPhoneNumber),
    };
})();

