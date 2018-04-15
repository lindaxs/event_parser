/*
  * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
  * See LICENSE in the project root for license information.
*/

'use strict';

  // The initialize function must be run each time a new page is loaded
Office.initialize = function (reason) {
  $(document).ready(function () {
    loadItemProps(Office.context.mailbox.item);
  });
}

function loadItemProps(item) {
  // Get the table body element
  var tbody = $('.prop-table');
  var body = item.body;
  
  body.getAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded){
        tbody.append(makeTableRow("XXX: ", "BAD"));
    } else {
      tbody.append(makeTableRow("XXX: ", asyncResult.value.trim()));
    }
  });

  // Add a row to the table for each message property
  // tbody.append(makeTableRow("Id", item.itemId));
  tbody.append(makeTableRow("Event:", item.subject));
  // tbody.append(makeTableRow("Message Id", item.internetMessageId));
  tbody.append(makeTableRow("With", item.from.displayName));
  tbody.append(makeTableRow("Date: "));
  tbody.append(makeTableRow("Time: "));
  tbody.append(makeTableRow("Location: "));
  
}

function makeTableRow(name, value) {
  return $("<tr><td><strong>" + name + 
    "</strong></td><td class=\"prop-val\"><code>" +
    value + "</code></td></tr>");
}

function loadEntities() {
  // getSelectedRegExMatches is in preview, so need to test for it
  if (Office.context.mailbox.item.getSelectedRegExMatches !== undefined) {
    var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
    if (selectedMatches) {
      // Note that the use of selectedMatches.OrderNumber, where
      // OrderNumber corresponds to the RegExName attribute of the Rule element
      // in the manifest
      $("#selected-match").text(JSON.stringify(selectedMatches.OrderNumber, null, 2));
    } else {
      $("#selected-match").text("Selected matches was null");
    }
  } else {
    $("#selected-match").text("Method not supported on your client");
  }

  // Get all matches
  var allMatches = Office.context.mailbox.item.getRegExMatches();
  if (allMatches) {
    // Note that the use of selectedMatches.OrderNumber, where
    // OrderNumber corresponds to the RegExName attribute of the Rule element
    // in the manifest
    $("#all-matches").text(JSON.stringify(allMatches.OrderNumber, null, 2));
  } else {
    $("#all-matches").text("All matches was null");
  }
}

function getBody(){
  var _item = Office.context.mailbox.item;
  var body = _item.body;
  
  
}
