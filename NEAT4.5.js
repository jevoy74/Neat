/************************************************************************************************
 *    Name:       NEAT4.5C.js 
 *    Version:    4.5C
 *    Author:     Jason Evoy - jevoy@avaya.com
 *    Date:       Feb 15th, 2012
 *    Notes:	  This is a joke script for Collen.  2 Pop up windows.  
 *    Min Req:    Windows 2000, 2003, XP, Internet Explorer 5.5 
 ************************************************************************************************/

PopupContinue();
PopupDone();


function PopupContinue()
{
   var InfoIcon = 64; var OkButton = 0; var Timeout = 15;
   var Title = "Nortel Enterprise Audit Tool Ver 4.5C                       - Jason Evoy";
   var MessageLine1 = "Creating log file...........................\n\nPlease wait until finished message appears before opening file\nThe audit usually takes less than 1 minute to complete";  
   var WshShell = WScript.CreateObject("WScript.Shell");
   var ButtonCode = WshShell.Popup(MessageLine1, Timeout, Title, OkButton + InfoIcon);
}

function PopupDone()
  {
    var OkButton = 0; var Timeout = 60; var InfoPic = 16;
    var Title = "CRITICAL SYSTEM ERROR";
    var shell = WScript.CreateObject("WScript.Shell");
    var DialogBox =  shell.Popup("This program has performed an illegal operation.\n\nThe NEAT tool is not designed to be a coffee pot that you can just throw around and break.\n\nPlease tell Colleen Long about this problem. ", Timeout, Title, OkButton + InfoPic);
    if (DialogBox == OkButton)
       {
        WScript.Quit();
       }
  }  
