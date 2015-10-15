
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
            var hostName = Office.context.mailbox.diagnostics.hostName;
            if (hostName == 'OutlookWebApp') {
                Office.context.mailbox.item.body.getAsync({ coercionType: "html" }, processHtmlBody);
            }
            else {
                Office.context.mailbox.item.body.getAsync("html", processHtmlBody);
            }
        });
    };


    function processHtmlBody(asyncResult) {
        var htmlParser = new DOMParser().parseFromString(asyncResult.value, "text/html");
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
                        $("#links-table").append("<div class='ms-Table-row ms-font-xs ms-font-color-white'>" +
                                               "<span class='ms-Table-cell normal-link'>" + vInnerText + "</span>" +
                                               "<span class='ms-Table-cell normal-link'>" + hrefText + "</span>" +
                                               "</div>");
                    }

                }
            );

        $('#result').append("Number of links found in this email: " + (normalLinkCount + phishyLinkCount) + " Number of phishy links (red): " + phishyLinkCount);
    }

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