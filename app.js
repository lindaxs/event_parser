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
      console.log("Error, failed to fetch!")
    } else {
      var string = get_details(asyncResult.value.trim());
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

function get_details(content) {
  content = content.replace(/[^a-z:0-9]/gmi, " ").replace(/\s+/g, " ").replace(/ +(?= )/g,'').toLowerCase();
  //add pm or some shit
  content = content.split(" ");
  var i;
  for(i = 0; i < content.length; i++){
    if(hasNumber(content[i])){
      if(!hasPM(content[i])){
        if(!hasPM(content[i+1])){
          content[i] = content[i] + 'pm';
        }
      }
    }
  }
  
  buildings = ['thayer', 'maclean', 'alpha chi', 'azd', 'gym', 'armana', 'baker', 'bartlett', 'berry', 'bildner', 'blunt', 'brace', 'burke', 'butterfield', 'byrne', 'carpenter', 'carson', 'foco', 'collis', 'xd', 'ekt', 'gile', 'goldstein', 'haldeman', 'hood', 'kd', 'kde', 'kkg', 'kemeny', 'mass', 'morton', 'murdough', 'new hamp', 'rauner', 'richardson', 'rocky', 'silsby', 'sd', 'sanborn', 'sudikoff', 'steele', 'thompson', 'topliff', 'zimmerman', 'lalacs', 'beta', 'paganucci lounge', 'sigma delt', 'fahey']
  two_word_buildings = ['edgerton house', 'dartmouth hall', 'hanover inn', 'north fairbanks', 'one wheelock', 'st thomas']
  
  function get_time(content){
    var k;
    for(k = 0; k < content.length; k++){
      var word = content[k];
      if(word == "midnight"){
        return "11:45";
      }
      if(hasNumber(word)){
        if(hasPM(word) || word.indexOf("-") != 0){
          return word;
        }
        if(word.indexOf(":") != 0){
          if(word.length < 8){
            return word
          }
        }
      }
    }
    
    return null
  }
  function get_date(content){
    for(var i = 0; i < content.length; i++){
      var word = content[i];
      var key_dates = ["tonight", "tomorrow", "today", 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
      for(var k = 0; k < key_dates.length; k++){
        if(word.indexOf(key_dates[k]) != -1){
          return word;
        }
      }
    }
    
    for(var i = 0; i < content.length; i++){
      var word = content[i];
      months = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'septemeber', 'october', 'november', 'december'];
      for(var k = 0; k < months.length; k++){
        if(word.indexOf(months[k]) != -1)
        return month + content[i+1];
      }
    }
    
    return "today";
  }
  
  function get_location(content){
    for(var i = 0; i < content.length; i++){
      word = content[i]
      for(var k = 0; k < buildings.length; k ++){
        if(word.indexOf(buildings[k]) != -1){
          return word;
        }
      }     
    }
    var full_string = content.join(" ");
    
    for(var i = 0; i < two_word_buildings.length; i++){
      if(full_string.indexOf(two_word_buildings[i]) != -1){
        return two_word_buildings[i];
      }
    }
    
    return null;
  }
  var time = get_time(content);
  var date = get_date(content);
  var locale = get_location(content);
  
  console.log(time + ";" + date + ";" + locale);
  
  if(time == null || date == null || locale == null){
    console.log(content);
  }
  return(time + ";" + date + ";" + locale);
}
  function hasNumber(myString) {
    return /\d/.test(myString);
  }
  function hasPM(myString) {
    return /pm/.test(myString);
  }
