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
  var string = "";
  
  body.getAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded){
      console.log("Error, failed to fetch!")
    } else {
      string = get_details(asyncResult.value.trim());
      var details_array = string.split(";");
      tbody.append(makeTableRow("Time: ", details_array[0]));
      tbody.append(makeTableRow("Date: ", details_array[1]));
      tbody.append(makeTableRow("Location: ", details_array[2]));
    }
  });

  // Add a row to the table for each message property
  // tbody.append(makeTableRow("Id", item.itemId));
  tbody.append(makeTableRow("Event:", item.subject));
  // tbody.append(makeTableRow("Message Id", item.internetMessageId));
  tbody.append(makeTableRow("With", item.from.displayName));
  // tbody.append(makeTableRow("Time: ", string));
  console.log(string);
  
}

function makeTableRow(name, value) {
  return $("<tr><td><strong>" + name + 
    "</strong></td><td class=\"prop-val\"><code>" +
    value + "</code></td></tr>");
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
          if(!(hasNumber(content[i][0]) && hasNumber(content[i][1]) && hasNumber(content[i][2]))){
            content[i] = content[i] + 'pm';
          }
        }
      }
    }
  }
  
  
  var buildings = ['thayer', 'maclean', 'alpha chi', 'axid', 'gym', 'armana', 'baker', 'bartlett', 'berry', 'bildner', 'blunt', 'brace', 'burke', 'butterfield', 'byrne', 'carpenter', 'carson', 'foco', 'collis', 'xd', 'ekt', 'gile', 'goldstein', 'haldeman', 'hood', 'kd', 'kde', 'kkg', 'kemeny', 'mass', 'morton', 'murdough', 'new hamp', 'rauner', 'richardson', 'rocky', 'silsby', 'sd', 'sanborn', 'sudikoff', 'steele', 'thompson', 'topliff', 'zimmerman', 'lalacs', 'beta', 'paganucci lounge', 'sigma delt', 'fahey']
  var two_word_buildings = ['dirt cowboy', 'edgerton house', 'dartmouth hall', 'hanover inn', 'north fairbanks', 'one wheelock', 'st thomas']
  
  function get_time(content){
    var k;
    var returned_word;
    for(k = 0; k < content.length; k++){
      var word = content[k];
      console.log(word);
      if(word == "midnight"){
        return "11:45";
      }
      if(hasNumber(word)){
        if(hasPM(word) && word.indexOf("-") != 0){
           returned_word = word;
        }
        if(word.indexOf(":") != 0){
          if(word.length < 18){
            return word;
          }
        }
        
      }
    }
    return returned_word
    
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
      var word = content[i]
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
