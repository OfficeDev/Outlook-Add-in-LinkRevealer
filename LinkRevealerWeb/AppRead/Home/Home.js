
/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            retrieveMessageBody();
        });
    };

    // Displays the "Subject" and "From" fields, based on the current mail item
    function retrieveMessageBody() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        var from;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            sendRequest();
        }
    }

    // Send an EWS request for the message's body. 
    function sendRequest() {
        var mailbox = Office.context.mailbox;
        var request = getBodyRequest(mailbox.item.itemId);
        var envelope = getSoapEnvelope(request);

        mailbox.makeEwsRequestAsync(envelope, findLinksInMessageBody);
    };

    function getBodyRequest(id) {
        // Return a GetItem EWS operation request for the subject of the specified item.  
        var result =

     '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
     '      <ItemShape>' +
     '        <t:BaseShape>IdOnly</t:BaseShape>' +
     '        <t:AdditionalProperties>' +
     '            <t:FieldURI FieldURI="item:Body"/>' +
     '        </t:AdditionalProperties>' +
     '      </ItemShape>' +
     '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
     '    </GetItem>';

        return result;
    };

    function getSoapEnvelope(request) {
        // Wrap an Exchange Web Services request in a SOAP envelope. 
        var result =

        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '  <soap:Header>' +
        '  <t:RequestServerVersion Version="Exchange2013"/>' +
        '  </soap:Header>' +
        '  <soap:Body>' +

        request +

        '  </soap:Body>' +
        '</soap:Envelope>';

        return result;
    };
    

    // Function called when the EWS request is complete. 
    function findLinksInMessageBody(asyncResult) {
        var response = asyncResult.value;
        var context = asyncResult.context;

        // Process the returned response here. 
        var responseSpan = document.getElementById("response");
        var xmlDoc = $.parseXML(response);
        var $xml = $(xmlDoc);


        var htmlParser = new DOMParser().parseFromString($xml.text(), "text/html");
        var links = htmlParser.getElementsByTagName("a");
        var phishyLinkCount = 0;
        var normalLinkCount = 0;
        $.each(
               links,
                function (i, v) {

                    // Add the phishy icon to any URLs that look suspect
                    var vInnerText = v.innerText.toLowerCase().trim();
                    var hrefText = v.href.toLowerCase().trim();
                    var phishyIcon = "";
                    var linkIsPhishy = ((vInnerText.search("http") == 0) && vInnerText != hrefText);
                    
                    if (linkIsPhishy) {
                        phishyLinkCount++;
                        $("#links-table").append("<div class='ms-Table-row ms-font-xs ms-bgColor-redDark ms-font-color-white'>" +
                                                "<span class='ms-Table-cell phishy-link'>" + vInnerText + "</span>" +
                                                "<span class='ms-Table-cell phishy-link'>" + hrefText + "</span>" +
                                                "</div>");
                    }
                    else {
                        normalLinkCount++;
                        $("#links-table").append("<div class='ms-Table-row ms-font-xs'>" +
                                               "<span class='ms-Table-cell normal-link'>" + vInnerText + "</span>" +
                                               "<span class='ms-Table-cell normal-link'>" + hrefText + "</span>" +
                                               "</div>");
                    }

                }
            );

        $('#result').append("Number of links found in this email: " + (normalLinkCount + phishyLinkCount) + " Number of phishy links (red): " + phishyLinkCount);
    };

})();

// *********************************************************
//
// Outlook-Add-in-LinkRevealer, https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************