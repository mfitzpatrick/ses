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
            console.log("Upgrading ses DB");
            if (!upgradeDB.objectStoreNames.contains('messages')) {
                console.log("Creating messages object store");
                var os = upgradeDB.createObjectStore('messages', {keyPath: 'id'});
                // the 'peer' index holds the remote peer's phone number for more easy querying
                os.createIndex('peer', 'peer', {unique: false});
            }
            break;
        }
    });

    function addMsg(msg, email, ts) {
        var peer = email.split("@", 1);
        //add this new message to our global database
        sesDBpromise.then(function(db) {
            var tx = db.transaction('messages', 'readwrite');
            var os = tx.objectStore('messages');
            var dbmsg = {
                id: msg.id,
                msg: msg,
                email: email,
                peer: peer,
                ts: ts,
            };
            os.add(dbmsg);
            return tx.complete;
        }).then(function() {
            //also add message to the relevant message chain
            if (addMsgCallback != null && msgCallbackFilterPeer == peer) {
                var is_sent = ("SENT" in msg.labelIds) ? true : false;
                addMsgCallback(msg, ts, is_sent);
            } else if (addMsgChainCallback != null) {
                // call the callback to inform it of the new message
                addMsgChainCallback(msg, peer, ts);
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

    function getUniquePeers(cb) {
        sesDBpromise.then(function(db) {
            var tx = db.transaction('messages', 'readonly');
            var os = tx.objectStore('messages');
            var idx = os.index('peer');
            idx.openCursor(null, 'nextunique').onsuccess = function(evt) {
                var cursor = evt.target.result;
                if (cursor) {
                    console.log("Read from DB:", cursor);
                    cb(cursor);
                    cursor.continue();
                } else {
                    console.log("Retrieved all DB entries");
                }
            };
        });
    }

    return {
        addMsg: (addMsg),
        updateMsgChainCallback: (updateMsgChainCallback),
        updateChatCallback: (updateChatCallback),
        getUniquePeers: (getUniquePeers),
    };
})();

