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
    // document.getElementById("run").onclick = run;
    run();
  }
});

export async function run() {
    // Get a reference to the current message
    var item = Office.context.mailbox.item;
    getRelatedMsgs(item.from.emailAddress);
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
        console.log("EWS:", result);
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
 * The idTuples object should be an array of [msgID, changeKey] objects.
 */
function getMsgBodies(idTuples) {
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
        console.log("Message:", result);
        var response = $.parseXML(result.value);
        var msgs = response.getElementsByTagName("t:Message");
        var idTuples = [];
        for (let item of msgs) {
            var body = item.getElementsByTagName("t:Body")[0].textContent;
            if (body.length > 100) {
                body = body.substring(0, 100);
            }
            addChatEntry(body);
        }
    });
}

/*
 * Add chat item's contents to the sidebar. Clone the template node, fill it in, and append it
 * to the parent node.
 */
function addChatEntry(body) {
    var template = document.querySelector("#chat-template");
    var clone = template.content.cloneNode(true);
    clone.querySelector("[name=chat-entry]").textContent = body;
    document.querySelector("#app-body").appendChild(clone);
}

