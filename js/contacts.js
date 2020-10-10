/*
 * Manage and interface with the contacts picker available in the Google Chrome API.
 */

var contacts = (function() {
    'use strict';
    var lookupWorker = null;
    var useContactsManager = false;

    if ('contacts' in navigator && 'ContactsManager' in window) {
        useContactsManager = true;
    } else {
        console.log("NB: The 'Contacts Picker API' is not available in this browser");
    }

    /*
     * Open the contacts picker API window and allow the user to select a contact (or list of
     * contacts) for us to use.
     */
    async function choose() {
        if (!useContactsManager) {
            console.log("Choose function is not yet implemented in browsers without contacts manager");
            return;
        }

        try {
            const contacts = await navigator.contacts.select(['name', 'tel'], {multiple: true});
            return contacts;
        } catch (ex) {
            console.log("Contacts picker errored:", ex);
        }
    }

    /*
     * Look up a user's name information based on their phone number. When the user is returned,
     * call the specified callback.
     * NB: this operates by an external callback, which must be set before this function can be used.
     * The external callback is required because JavaScript cannot look up the local contacts list
     * on a device directly at this time.
     */
    function lookup(phone, cb) {
        if (lookupWorker == null) {
            console.log("Contacts lookup() worker function has not been set");
            return null;
        }
        return lookupWorker(phone);
    }

    /*
     * Set the lookup backend worker function.
     */
    function setLookupWorker(cb) {
        lookupWorker = cb;
    }

    return {
        choose: (choose),
        lookup: (lookup),
        setLookupWorker: (setLookupWorker),
    };
})();

