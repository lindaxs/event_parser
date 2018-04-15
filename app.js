/*
  * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
  * See LICENSE in the project root for license information.
*/

'use strict';

(function () {
  
  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      loadItemProps(Office.context.mailbox.item);
      getBody();
    });
  };
  
  function loadItemProps(item) {
    // Get the table body element
    var tbody = $('.prop-table');
    
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
  
function getBody(){
    var _item = Office.context.mailbox.item;
    var body = _item.body;
    
    body.getAsync(Office.CoercionType.Text, function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded){
        //Handle Error
        } else {
        showDataDialog('Body', asyncResult.value.trim());
      }
    });
  }
  
})();