/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
import axios from 'axios'; 
  

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    debugger;
    const item = Office.context.mailbox.item;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("item-subject").innerHTML = " " + item.subject;
    document.getElementById("exampleFormControlTextarea1").innerHTML =  item.body+Office.context.mailbox.item?.body.getAsync(Office.CoercionType.Html, (asyncResult) => {
      console.log(asyncResult.value);
});

Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {   if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {  
  
  } else 
 {   
   document.getElementById("exampleFormControlTextarea1").innerHTML =asyncResult.value;   }});




  }
  abc();
});

function abc()
{
 
  axios.get('https://core-dev.markit-systems.com/~tanusree/Sites/Trinity/dataImportApi/getTitleDropdown') 
  .then(response => { 
      const responseData = response.data; // Access the response data 
      // Process the response data here 
      const countriesDropDown = document.getElementById("countriesDropDown");
      debugger;
      for (let i = 0; i <responseData.DATA.length; i++) {
      
        let option = document.createElement("option");
        option.setAttribute('value', responseData.DATA[i].title_id);
      
        let optionText = document.createTextNode(responseData.DATA[i].title);
        option.appendChild(optionText);
      
        countriesDropDown.appendChild(option);
      }
  }) 
  .catch(error => { 
      // Handle any errors 
  });

}


function datasave(firstname,lastname,body,attachment)
{

  
  axios.post('https://core-dev.markit-systems.com/~tanusree/Sites/Trinity/dataImportApi/insertDataOutLookTable', {
    "firstName": firstname,
    "lastName": lastname,
    "emailBody": body,
    "emailattachment": attachment
  })
  .then(function (response) {
    console.log(response);
    debugger;
    document.getElementById("reponseid").innerHTML = response.data.DATA+response.status;
  })
  .catch(function (error) {
    console.log(error);
  });
}





export async function run() {
  /**
   * Insert your Outlook code here
   * 
   */

  // Get a reference to the current message
const item = Office.context.mailbox.item;

// Write message property value to the task pane


datasave("Aniruddho","kodali",item.subject,item.body);
}
