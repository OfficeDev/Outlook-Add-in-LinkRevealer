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
        $.each(
               links,
                function (i, v) {

                    // Add the phishy icon to any URLs that look suspect
                    var vInnerText = v.innerText.toLowerCase().trim();
                    var hrefText = v.href.toLowerCase().trim();
                    var phishyIcon = "";
                    var linkIsPhishy = false;
                    
                    if (vInnerText.search("http") == 0)
                    {
                        if (vInnerText != hrefText)
                        {
                            linkIsPhishy = true;
                            phishyLinkCount++;
                        }
                    }

                    if (linkIsPhishy) {
                        //Display the link text
                        $("#links-table").append("<tr><th>" + vInnerText + "</th><th>" + hrefText  + "</tr>");
                    }
                    else {
                        // Display the URL behind the link text
                        $("#links-table").append("<tr><th></th><th>" + vInnerText + "</th><th>" + hrefText + "</tr>");
                    }
                }
            );
    };

})();