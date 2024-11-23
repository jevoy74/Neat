/**********************************************************************************************************
 *    Name:       SCCSAuditv2.5.js 
 *    Version:    2.5
 *    Author:     Jason Evoy - jevoy@nortel.com
 *    Date:       May 4, 2005
 *    Notes:	  This is a script that polls the registry & WMI for PC data & 
 *                outputs it to a file.  This script is for auditing purposes
 *                only, and does not make any changes to the PC. 
 *                Designed for SCCS, SECC, SWC, & OTM, but will work on any PC running W2K or W2K3
 *                Email bugs/comments to jevoy@nortel.com                 
 **********************************************************************************************************/
var TotalMessage = "";
PopupContinue();
MakeFile();
PopupDone();

function PopupContinue()
{
   var InfoIcon        = 64;
   var OkButton	       = 0;
   var Timeout         = 4;
   var Title = "Nortel Enterprise Audit Tool Ver 2.5                       - Jason Evoy";
   var MessageLine1 = "Creating log file...........................\n\nPlease wait until finished message appears before opening file\n";  
   var WshShell = WScript.CreateObject("WScript.Shell");
   var ButtonCode = WshShell.Popup(MessageLine1, Timeout, Title, OkButton + InfoIcon);

}

function PopupDone()
  {
    var OkButton = 0; var Timeout = 4; var InfoPic = 64;
    var Title = "Nortel Enterprise Audit Tool Ver 2.5                     - Jason Evoy";
    var shell = WScript.CreateObject("WScript.Shell");
    var DialogBox =  shell.Popup("Completed.  Log file saved as: " + GetComputerName() + ".txt", Timeout, Title, OkButton + InfoPic);
    if (DialogBox == OkButton)
       {
        WScript.Quit();
       }
  }  

function GetComputerName()
{
    	var TotalMessage = "";
	try
    { 
        var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\SYSTEM\\ControlSet001\\Control\\ComputerName\\ActiveComputerName\\ComputerName";
        TotalMessage = shell.RegRead( regkey1 ) ;
	return( TotalMessage );
    }
    catch( e )
    {
        return( "" );
    }
}

function MakeFile()
   {
   var forReading = 1, forWriting = 2, forAppending = 8, cr = "\r\n";
   var shell = new ActiveXObject( "Scripting.FileSystemObject" );
   shell.CreateTextFile( GetComputerName()+ ".txt" );
   var os = shell.GetFile( GetComputerName() + ".txt" );
   os = os.OpenAsTextStream( forAppending, 0 );
   os.write("+------------------------------------------------------------------------------+" + cr);
   os.write("¦Nortel Enterprise Audit Tool Results - ");
   os.write(Date());
   os.write(cr);
   os.write("+------------------------------------------------------------------------------+" + cr +cr);   
   os.write(GetICCMINFO());
   os.write(GetSWCINFO());
   os.write(GetOTMInfo());
   os.write(GetSUDATETIME());	
   os.write(GetNBNMINFO());
   os.write(GetAccessLink());
   os.write(GetM1Info());
   os.write(GetOS());
   os.write(cr);
   os.write(chkcpu());
   os.write(chkmemory());
   os.write(chkpagefile());
   os.write(ShowNetwork());
   os.write(GetDAYLIGHTSAVINGSTIME());
   os.write(GetMEDIASENSING());
   os.write(GetIPFORWARDING());
   os.write(GetAPIPA());
   os.write(GetNIC());
   os.write(ShowDriveList());
   os.write(cr);
   os.write(GetPATHLENGTH());  
   os.write(GetPATH());
   os.write(cr);
   os.write("+------------------------------------------------------------------------------+" + cr);
   os.write("¦Running Processes                                                             ¦" + cr);
   os.write("+------------------------------------------------------------------------------+" + cr + cr);
   os.write(GetPROCESS() +cr );
   os.write("+------------------------------------------------------------------------------+" + cr);
   os.write("¦Services                                                                      ¦" + cr);
   os.write("+------------------------------------------------------------------------------+" + cr + cr);
   os.write(chkservices());
   os.write("+------------------------------------------------------------------------------+" + cr);
   os.write("¦Installed Software (Through Windows Installer)                                ¦" + cr);
   os.write("+------------------------------------------------------------------------------+" + cr + cr);
   os.write(GetInstalledPrograms() +cr);	
   os.write("+------------------------------------------------------------------------------+" + cr);
   os.write("¦Registry Run Keys                                                             ¦" + cr);
   os.write("+------------------------------------------------------------------------------+" + cr + cr);
   TotalMessage = "";
   os.write(GetRunKey());			
   os.Close();
   }

function GetNBNMINFO()
{
    try
    { 
        var cr = "\r\n";
	var TotalMessage = "";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Nortel\\Setup\\NINAMESERVER\\CLan";
	var regkey2 = "HKLM\\Software\\Nortel\\Setup\\NINAMESERVER\\ELan";
        var regkey3 = "HKLM\\Software\\Nortel\\Setup\\NINAMESERVER\\NSAddr";
	var regkey5 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MultiCastAddr";
	var regkey4 = "HKLM\\Software\\Nortel\\Setup\\NINAMESERVER\\SiteName";
	var regkey6 = "HKLM\\System\\CurrentControlSet\\Services\\NBNM_Service\\ImagePath";
        TotalMessage += "Site Name: ";
	TotalMessage += (shell.RegRead( regkey4 ) + cr );
	TotalMessage += "CLAN IP: ";
	TotalMessage += (shell.RegRead( regkey1 ) + cr );
	TotalMessage += "NBNM IP: ";
	TotalMessage += (shell.RegRead( regkey3 ) + " Path:" + shell.RegRead( regkey6 ) + cr );	
	TotalMessage += "ELAN IP: ";
	TotalMessage += (shell.RegRead( regkey2 ) + cr );
	TotalMessage += "Multicast IP: ";
	TotalMessage += (shell.RegRead( regkey5 ) + cr );
	return( TotalMessage);
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetICCMINFO()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Symposium Info:" + cr + "Keycode: ";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Nortel\\Setup\\ICCMKeycode";
 	var regkey2 = "HKLM\\Software\\Nortel\\Setup\\ICCMSerial";
  	var regkey3 = "HKLM\\Software\\Nortel\\Setup\\Platform Version";
  	var regkey4 = "HKLM\\Software\\Nortel\\SCCS\\SU Record\\SUID";
	var regkey4a = "HKLM\\Software\\Nortel\\SCCS\\SU Record\\Time";
	var regkey4b = "HKLM\\Software\\Nortel\\SCCS\\SU Record\\Date";
	var regkey5 = "HKLM\\Software\\Nortel\\Setup\\UserCompany";
	var regkey6 = "HKLM\\Software\\Nortel\\Setup\\UserName";
        TotalMessage += (shell.RegRead( regkey1 ) + cr );
	TotalMessage += "Serial: ";
	TotalMessage += (shell.RegRead( regkey2 ) + cr );
	TotalMessage += "Version: ";
	TotalMessage += (shell.RegRead( regkey3 ) + cr );
	TotalMessage += "Company Name: ";
	TotalMessage += (shell.RegRead( regkey5 ) + cr );
	TotalMessage += "Customer Name: ";
	TotalMessage += (shell.RegRead( regkey6 ) + cr );
	TotalMessage += "Service Update: ";
	TotalMessage += (shell.RegRead( regkey4 ) + cr );
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetSWCINFO()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Web Client Info:" + cr + "Keycode: ";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Nortel\\WClient\\KeyCode";
 	var regkey2 = "HKLM\\Software\\Nortel\\WClient\\diskserial";
  	var regkey3 = "HKLM\\Software\\Nortel\\WClient\\Version";
  	var regkey4 = "HKLM\\Software\\Nortel\\RTD\\IP Send";
	var regkey5 = "HKLM\\Software\\Nortel\\RTD\\IP Receive";
	var regkey6 = "HKLM\\Software\\Nortel\\EmergencyHelp\\IP Send";
        TotalMessage += (shell.RegRead( regkey1 ) + cr );
	TotalMessage += "Serial Number: ";
	TotalMessage += (shell.RegRead( regkey2 ) + cr );
	TotalMessage += "Version: ";
	TotalMessage += (shell.RegRead( regkey3 ) + cr );
	TotalMessage += "RTD IP Send: ";
	TotalMessage += (shell.RegRead( regkey4 ) + cr );
	TotalMessage += "RTD IP Receive: ";
	TotalMessage += (shell.RegRead( regkey5 ) + cr );
	TotalMessage += "Emergency Help: ";
	TotalMessage += (shell.RegRead( regkey6 ) + cr );
	return( TotalMessage + cr );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetSUDATETIME()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Service update installed on:";
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey1 = "HKLM\\Software\\Nortel\\SCCS\\SU Record\\Time";
	var regkey2 = "HKLM\\Software\\Nortel\\SCCS\\SU Record\\Date";
	TotalMessage += (shell.RegRead( regkey2 )); 
	TotalMessage += " at " + (shell.RegRead( regkey1 ) + cr );
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}
function GetM1Info()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Meridian 1 / CSE1k Info: "+cr;
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Nortel\\Setup\\Switch\\Serial";
	var regkey2 = "HKLM\\Software\\Nortel\\Setup\\Switch\\Name";
	var regkey3 = "HKLM\\Software\\Nortel\\Setup\\Switch\\IP";
 	var regkey4 = "HKLM\\Software\\Nortel\\Setup\\Switch\\Customer";
	TotalMessage += ("Serial: " + shell.RegRead( regkey1 ) +cr);
	TotalMessage += ("IP: " + shell.RegRead( regkey3 ) + cr);
	TotalMessage += ("IP HostName: " + shell.RegRead( regkey2 ) + cr);
	TotalMessage += ("Customer: " + shell.RegRead( regkey4 ) + cr);
	return( TotalMessage + cr );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetIPFORWARDING()
{
    try
    {
        var IPFORWARD, TotalMessage = "IP Forwarding is ", cr = "\r\n";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\IPEnableRouter";
        IPFORWARD = shell.RegRead( regkey );
	if (IPFORWARD == 0)
	   {
		TotalMessage += "Disabled";
           }
	if (IPFORWARD == 1)
	   {   		
		TotalMessage += "Enabled";
	   }	
        return( TotalMessage + cr );
    }
    catch( e )
    {
        return( "IP Forwarding: " + e.description + cr );
    }
}	

function GetAPIPA()
{
    try
    {
        var TotalMessage = "Automatic Private IP Addressing is ", cr = "\r\n";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\IPAutoconfigurationEnabled";       
        var APIPA = shell.RegRead( regkey );
	if (APIPA == 0) TotalMessage += "Disabled";
        else TotalMessage += "Enabled";
		
        return( TotalMessage + cr );
    }
    catch( e )
    {
        return( TotalMessage + "Enabled" + cr);
    }
}	

function GetOS()
{
    try
    {
        var cr = "\r\n"; var TotalMessage = "Operating System Info:\r\n"; 
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProductName";
 	var regkey2 = "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\CSDVersion";
 	var regkey3 = "HKLM\\SYSTEM\\ControlSet001\\Control\\ComputerName\\ActiveComputerName\\ComputerName";
	var regkey4 = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Hostname";
	var regkey5 = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Domain";
        TotalMessage += ("OS: " + shell.RegRead( regkey1 ));
	TotalMessage += (" " + shell.RegRead( regkey2 ) + cr);
	TotalMessage += ("Computer Name: " + shell.RegRead( regkey3 ) + cr);
	TotalMessage += ("Hostname: " + shell.RegRead( regkey4 ) + cr);
	TotalMessage += ("Domain: " + shell.RegRead( regkey5 ) + cr);
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetDAYLIGHTSAVINGSTIME()
{
    try
    {
        var DST, cr = "\r\n";
	var TotalMessage = "Daylight Savings Time is ";
        var shell = new ActiveXObject( "WScript.Shell" );
        var regkey = "HKLM\\SYSTEM\\ControlSet001\\Control\\TimeZoneInformation\\DisableAutoDaylightTimeSet";
        DST = shell.RegRead( regkey );
	if (DST == 1)
	{
	   TotalMessage += "Disabled";	
	   return( TotalMessage + cr );
        }
    }
    catch( e )
    {
	TotalMessage += "Enabled"; 
        return( TotalMessage + cr );
    } 
}

function GetMEDIASENSING()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Media Sensing feature is ";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\DisableDHCPMediaSense";
        var DHCP = shell.RegRead( regkey );
	if (DHCP == 1)
	{
	   TotalMessage += "Disabled";	
	}
    }
    catch( e )
    {
	TotalMessage += "Enabled"; 
    }
return(TotalMessage + cr); 
}

function ShowDriveList() 
{ 
   var cr = "\r\n";
   var MB, DiskDrive, TypeOfDrive; 
   var OneMeg = "1048576";
   var TotalMessage = "Drive Information:" + cr;
   var fs = new ActiveXObject("Scripting.FileSystemObject"); 
   var i = new Enumerator(fs.Drives); 

   for (; !i.atEnd(); i.moveNext()) 
   { 
     DiskDrive = i.item(); 
     TotalMessage = TotalMessage + DiskDrive.DriveLetter; 
     TotalMessage += " - "; 
     if (DiskDrive.DriveType == 0 ) TypeOfDrive = "Unknown Drive\t";
     if (DiskDrive.DriveType == 1 ) TypeOfDrive = "Floppy Disk\t";
     if (DiskDrive.DriveType == 2 ) TypeOfDrive = "Fixed Disk\t";
     if (DiskDrive.DriveType == 3 ) TypeOfDrive = "Network Drive\t";
     if (DiskDrive.DriveType == 4 ) TypeOfDrive = "CD-ROM\t";
     if (DiskDrive.DriveType == 5 ) TypeOfDrive = "RAM Disk\t";
     TotalMessage += TypeOfDrive;
     
     if (DiskDrive.IsReady) 
     	{
      	MB = (DiskDrive.TotalSize / OneMeg);
      	MB = (MB * 10);
      	MB = Math.round(MB);
      	MB = (MB / 10);	
      	MB += " MB - ";
      	TotalMessage += MB;
      	TotalMessage = TotalMessage + DiskDrive.FileSystem;
      	TotalMessage += cr;
        }
     else (TotalMessage += cr);
  }
  return (TotalMessage ); 
}

function ShowNetwork()
{
   var cr = "\r\n";
   var shell = WScript.CreateObject("WScript.Network");
   var TotalMessage = "";
   TotalMessage += ("Logged in as: " + shell.UserName + cr);
   TotalMessage += ("User Domain: " + shell.UserDomain + cr);
   return ( TotalMessage );
}

function GetPATH()
{
   var shell = new ActiveXObject( "WScript.Shell" );
   var WshEnv = shell.Environment("Process");
   var cr = "\r\n";
   var TotalMessage = cr;
   TotalMessage += ("Path: ");
   TotalMessage += (WshEnv("PATH") + cr);
   return (TotalMessage + cr);
}

function GetPATHLENGTH()
{
   var shell = new ActiveXObject( "WScript.Shell" );
   var WshEnv = shell.Environment("Process");
   var TotalMessage = "Path is ";
   var cr = "\r\n";
   Pathlength = WshEnv("PATH");
   return ( TotalMessage + Pathlength.length + " characters " + cr);
}

function chkservices()
{
   var Locator = new ActiveXObject ("WbemScripting.SWbemLocator");
   var Service = Locator.connectserver();
   var serviceerr = "0", TotalMessage = "", cr = "\r\n";
   
   try
      {
      var services = Service.Get ("Win32_Service");
      var iservice = new Enumerator (services.Instances_());
      }
   catch (error)
      {
      Return ("Could not get Services info because: " + error.description);
      }

   for (;!iservice.atEnd();iservice.moveNext())
   {
      var serviceInstance = iservice.item();
      TotalMessage += (serviceInstance.DisplayName + " - " + 
      serviceInstance.StartMode + " & " + serviceInstance.State + cr );
   }
return (TotalMessage + cr);
}

function GetPROCESS()
{
   var Locator = new ActiveXObject ("WbemScripting.SWbemLocator");
   var process = Locator.connectserver();
   var TotalMessage = "";
   var cr = "\r\n"; var tab = "\t";
   var OkButton = 0; var Timeout = 0; var WarnPic = 48;
   var shell = WScript.CreateObject("WScript.Shell");
 
   try
   {
	var processes = process.Get ("Win32_Process");
        var iprocess = new Enumerator (processes.Instances_());
   }
   catch ( e )
   {
        return ("Processes: " + e.description);
   }

 for (;!iprocess.atEnd();iprocess.moveNext())
   {
	var PRO = iprocess.item();
	  if (PRO.Name != ("System Idle Process" && "wscript.exe"))
             {
	     TotalMessage += PRO.Name + "      " + tab;
	     TotalMessage += "PID: " + PRO.ProcessID + "      " + tab;       
 	     TotalMessage += "Handles: " + PRO.HandleCount + tab;
	     TotalMessage += "Threads: " + PRO.ThreadCount + tab;
 	     TotalMessage += cr;
	     }
   }
return (TotalMessage);
}

function GetAccessLink()
{
    try
    {
        var TCPIP, TotalMessage = "Access Link is using ", cr = "\r\n";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\SOFTWARE\\Nortel\\LinkHandler\\nbalh\\Device";
	var regkey2 = "HKLM\\SOFTWARE\\Nortel\\LinkHandler\\nbalh\\IPAddress";
	var regkey3 = "HKLM\\SOFTWARE\\Nortel\\LinkHandler\\nbalh\\Linkspeed";
	var regkey4 = "HKLM\\SOFTWARE\\Nortel\\LinkHandler\\nbalh\\UseTCP";
        TCPIP = shell.RegRead( regkey4 );
	if (TCPIP == 0)
	   {
		TotalMessage += shell.RegRead( regkey1 )
		TotalMessage += " " + shell.RegRead( regkey3 ) + " baud" + cr;
           }
	if (TCPIP == 1)
	   {   		
		TotalMessage += shell.RegRead( regkey2 )
		TotalMessage += " port " + shell.RegRead( regkey1 ) + cr;
	   }	
        return( TotalMessage + cr );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}	

function GetInstalledPrograms()
{
	sRegTypes = new Array(
   	   "                              ",    // 0
   	   "REG_SZ                        ",    // 1
  	   "REG_EXPAND_SZ                 ",    // 2
  	   "REG_BINARY                    ",    // 3
 	   "REG_DWORD                     ",    // 4
 	   "REG_DWORD_BIG_ENDIAN          ",    // 5
 	   "REG_LINK                      ",    // 6
 	   "REG_MULTI_SZ                  ",    // 7
 	   "REG_RESOURCE_LIST             ",    // 8
 	   "REG_FULL_RESOURCE_DESCRIPTOR  ",    // 9
 	   "REG_RESOURCE_REQUIREMENTS_LIST",    // 10
 	   "REG_QWORD                    ");    // 11
	

	HKLM = 0x80000002;
	HKCR = 0x80000000;
	HKCU = 0x80000001;
	HKUS = 0x80000003;
	HKCC = 0x80000005;
	
	RegistryPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"

	Locator = new ActiveXObject("WbemScripting.SWbemLocator");
	oSvc = Locator.ConnectServer("127.0.0.1", "root\\default", "", "");
	oReg = oSvc.Get("StdRegProv");
	oMethod = oReg.Methods_.Item("EnumKey");
	oInParam = oMethod.InParameters.SpawnInstance_();
	oInParam.hDefKey = HKLM;
	oInParam.sSubKeyName = RegistryPath;
	oOutParam = oReg.ExecMethod_(oMethod.Name, oInParam);
	aNames = oOutParam.sNames.toArray();

	for (x = 0; x <= aNames.length -1; x++)
	{
	   var KeyName = aNames[x]
	   GetSoftwareKey (KeyName);
   	} 
return (TotalMessage);
}

function GetRunKey()
{
	sRegTypes = new Array(
   	   "                              ",    // 0
   	   "REG_SZ                        ",    // 1
  	   "REG_EXPAND_SZ                 ",    // 2
  	   "REG_BINARY                    ",    // 3
 	   "REG_DWORD                     ",    // 4
 	   "REG_DWORD_BIG_ENDIAN          ",    // 5
 	   "REG_LINK                      ",    // 6
 	   "REG_MULTI_SZ                  ",    // 7
 	   "REG_RESOURCE_LIST             ",    // 8
 	   "REG_FULL_RESOURCE_DESCRIPTOR  ",    // 9
 	   "REG_RESOURCE_REQUIREMENTS_LIST",    // 10
 	   "REG_QWORD                    ");    // 11
	
	HKLM = 0x80000002;
	HKCR = 0x80000000;
	HKCU = 0x80000001;
	HKUS = 0x80000003;
	HKCC = 0x80000005;
	
	RegistryPath1 = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\"

	Locator = new ActiveXObject("WbemScripting.SWbemLocator");
	oSvc = Locator.ConnectServer("127.0.0.1", "root\\default", "", "");
	oReg = oSvc.Get("StdRegProv");
	oMethod = oReg.Methods_.Item("EnumKey");
	oInParam = oMethod.InParameters.SpawnInstance_();
	oInParam.hDefKey = HKLM;
	oInParam.sSubKeyName = RegistryPath1;
	oOutParam = oReg.ExecMethod_(oMethod.Name, oInParam);
	aNames = oOutParam.sNames.toArray();

	for (x = 0; x <= aNames.length -1; x++)
	{
	   var KeyName1 = aNames[x]
           if (KeyName1 == "Run" || KeyName1 == "RunOnce" || KeyName1 == "RunOnceEx") { GetRunSoftwareKey (KeyName1); }
   	} 
return (TotalMessage);
}

function GetSoftwareKey (KeyName)
{

try
   {
   oMethod = oReg.Methods_.Item("EnumValues");
   oInParam = oMethod.InParameters.SpawnInstance_();
   oInParam.hDefKey = HKLM;
   oInParam.sSubKeyName = RegistryPath + "\\" + KeyName;
   oOutParam = oReg.ExecMethod_(oMethod.Name, oInParam);
   var newreg = RegistryPath + "\\" + KeyName
   bNames = oOutParam.sNames.toArray();
   aTypes = oOutParam.Types.toArray();
   
   for (y = 0; y <= bNames.length -1; y++)
   {
      if (aTypes[y] == "1")
         {
         subkeyval = bNames[y];
            if (subkeyval == "DisplayName" || subkeyval == "QuietDisplayName")
            {
   	    oMethod = oReg.Methods_.Item("GetStringValue");
            oInParam = oMethod.InParameters.SpawnInstance_();
            oInParam.hDefKey = HKLM;
            oInParam.sSubKeyName = newreg;
            oInParam.sValueName = subkeyval;
            oOutParam = oReg.ExecMethod_(oMethod.Name, oInParam);
            cNames = oOutParam.sValue
            TotalMessage += cNames;
            TotalMessage += "\r\n";
            }
            else
            {
            }
         }
      else
      {
      }
   }
}
   catch (e)
   {
   } 
}

function GetRunSoftwareKey (KeyName1)
{

try
   {
   oMethod = oReg.Methods_.Item("EnumValues");
   oInParam = oMethod.InParameters.SpawnInstance_();
   oInParam.hDefKey = HKLM;
   oInParam.sSubKeyName = RegistryPath1 + "\\" + KeyName1;
   oOutParam = oReg.ExecMethod_(oMethod.Name, oInParam);
   var newreg = RegistryPath1 + "\\" + KeyName1
   bNames = oOutParam.sNames.toArray();
   aTypes = oOutParam.Types.toArray();
   
   for (y = 0; y <= bNames.length -1; y++)
   {
      if (aTypes[y] == "1")
         {
         subkeyval = bNames[y];
            if (subkeyval != "(Default)")
            {
   	    oMethod = oReg.Methods_.Item("GetStringValue");
            oInParam = oMethod.InParameters.SpawnInstance_();
            oInParam.hDefKey = HKLM;
            oInParam.sSubKeyName = newreg;
            oInParam.sValueName = subkeyval;
            oOutParam = oReg.ExecMethod_(oMethod.Name, oInParam);
            cNames = oOutParam.sValue
	    TotalMessage += subkeyval;
	    TotalMessage += "\t\t\t";	
            TotalMessage += cNames;
            TotalMessage += "\r\n";
            }
            else
            {
            }
         }
      else
      {
      }
   }
}
   catch (e)
   {
   } 
}

function chkmemory()
{
   var Locator = new ActiveXObject ("WbemScripting.SWbemLocator");
   var Memory = Locator.connectserver();
   var memoryerr = "0", TotalMessage = "", cr = "\r\n";
   
   try
      {
      var memorys = Memory.Get ("Win32_OperatingSystem");
      var imemory = new Enumerator (memorys.Instances_());
      }
   catch (error)
      {
      Return ("Could not get memory info because: " + error.description);
      }

   for (;!imemory.atEnd();imemory.moveNext())
   {
      var memoryInstance = imemory.item();
      TotalMessage 
      += ( cr + "Total Physical Memory: " + memoryInstance.TotalVisibleMemorySize + cr	      
      + "Free Physical Memory: " + memoryInstance.FreePhysicalMemory +cr 
      + "Total Virtual Memory Size:" + memoryInstance.TotalVirtualMemorySize + cr
      + "Free Virtual Memory: " + memoryInstance.FreeVirtualMemory + cr);
   }
return (TotalMessage);
}

function chkpagefile()
{
   var Locator = new ActiveXObject ("WbemScripting.SWbemLocator");
   var Pagefile = Locator.connectserver();
   var pagefileerr = "0", TotalPageMessage = "", cr = "\r\n";
   
   try
      {
      var pagefiles = Pagefile.Get ("Win32_PageFile");
      var ipagefile = new Enumerator (pagefiles.Instances_());
      }
   catch (error)
      {
      Return ("Could not get pagefile info because: " + error.description);
      }

   for (;!ipagefile.atEnd();ipagefile.moveNext())
   {
      var pagefileInstance = ipagefile.item();
      TotalPageMessage += ("Initial paging file size: " + pagefileInstance.InitialSize + " MB" + cr
      + "Maximum paging file size: " + pagefileInstance.MaximumSize + " MB" + cr
      + "Path: " + pagefileInstance.Name + cr); 

   }
return (TotalPageMessage + cr);
}

function chkcpu()
{
   var Locator = new ActiveXObject ("WbemScripting.SWbemLocator");
   var CPU = Locator.connectserver();
   var cpuerr = "0", TotalMessage = "", cr = "\r\n";
   var StoreFamily = 0;
   var CPUFamily = "";
   var counter = 1;
   try
      {
      var cpus = CPU.Get ("Win32_Processor");
      var icpu = new Enumerator (cpus.Instances_());
      }
   catch (error)
      {
      Return ("Could not get CPU info because: " + error.description);
      }

   for (;!icpu.atEnd();icpu.moveNext())
   {
   var cpuInstance = icpu.item();
   TotalMessage += ( cpuInstance.Role + " " + counter )
   StoreFamily = cpuInstance.Family; 
   if (StoreFamily == 3 ) CPUFamily = " 8086 ";
   if (StoreFamily == 4 ) CPUFamily = " 80286 ";
   if (StoreFamily == 5 ) CPUFamily = " 80386 ";
   if (StoreFamily == 6 ) CPUFamily = " 80486 ";
   if (StoreFamily == 7 ) CPUFamily = " 8087 ";
   if (StoreFamily == 8 ) CPUFamily = " 80287 ";
   if (StoreFamily == 9 ) CPUFamily = " 80387 ";
   if (StoreFamily == 10 ) CPUFamily = " 80487 ";
   if (StoreFamily == 11 ) CPUFamily = " Pentium® brand ";
   if (StoreFamily == 12 ) CPUFamily = " Pentium® Pro ";
   if (StoreFamily == 13 ) CPUFamily = " Pentium® II ";
   if (StoreFamily == 14 ) CPUFamily = " Pentium® processor with MMX technology ";
   if (StoreFamily == 15 ) CPUFamily = " Celeron™ ";
   if (StoreFamily == 16 ) CPUFamily = " Pentium® II Xeon ";
   if (StoreFamily == 17 ) CPUFamily = " Pentium® III ";
   if (StoreFamily == 18 ) CPUFamily = " M1 Family ";
   if (StoreFamily == 19 ) CPUFamily = " M2 Family ";
   if (StoreFamily == 24 ) CPUFamily = " K5 Family ";
   if (StoreFamily == 25 ) CPUFamily = " K6 Family ";
   if (StoreFamily == 26 ) CPUFamily = " K6-2 ";
   if (StoreFamily == 27 ) CPUFamily = " K6-3 ";
   if (StoreFamily == 28 ) CPUFamily = " AMD Athlon™ Processor Family ";
   if (StoreFamily == 29 ) CPUFamily = " AMD® Duron™ Processor ";
   if (StoreFamily == 30 ) CPUFamily = " AMD2900 Family ";
   if (StoreFamily == 31 ) CPUFamily = " K6-2+ ";
   if (StoreFamily == 32 ) CPUFamily = " Power PC Family ";
   if (StoreFamily == 33 ) CPUFamily = " Power PC 601 ";
   if (StoreFamily == 34 ) CPUFamily = " Power PC 603 ";
   if (StoreFamily == 35 ) CPUFamily = " Power PC 603+ ";
   if (StoreFamily == 36 ) CPUFamily = " Power PC 604 ";
   if (StoreFamily == 37 ) CPUFamily = " Power PC 620 ";
   if (StoreFamily == 38 ) CPUFamily = " Power PC X704 ";
   if (StoreFamily == 39 ) CPUFamily = " Power PC 750 ";
   if (StoreFamily == 48 ) CPUFamily = " Alpha Family ";
   if (StoreFamily == 49 ) CPUFamily = " Alpha 21064 ";
   if (StoreFamily == 50 ) CPUFamily = " Alpha 21066 ";
   if (StoreFamily == 51 ) CPUFamily = " Alpha 21164 ";
   if (StoreFamily == 52 ) CPUFamily = " Alpha 21164PC ";
   if (StoreFamily == 53 ) CPUFamily = " Alpha 21164a ";
   if (StoreFamily == 54 ) CPUFamily = " Alpha 21264 ";
   if (StoreFamily == 55 ) CPUFamily = " Alpha 21364 ";
   if (StoreFamily == 64 ) CPUFamily = " MIPS Family ";
   if (StoreFamily == 65 ) CPUFamily = " MIPS R4000 ";
   if (StoreFamily == 66 ) CPUFamily = " MIPS R4200 ";
   if (StoreFamily == 67 ) CPUFamily = " MIPS R4400 ";
   if (StoreFamily == 68 ) CPUFamily = " MIPS R4600 ";
   if (StoreFamily == 69 ) CPUFamily = " MIPS R10000 ";
   if (StoreFamily == 80 ) CPUFamily = " SPARC Family ";
   if (StoreFamily == 81 ) CPUFamily = " SuperSPARC ";
   if (StoreFamily == 82 ) CPUFamily = " microSPARC II ";
   if (StoreFamily == 83 ) CPUFamily = " microSPARC IIep ";
   if (StoreFamily == 84 ) CPUFamily = " UltraSPARC ";
   if (StoreFamily == 85 ) CPUFamily = " UltraSPARC II ";
   if (StoreFamily == 86 ) CPUFamily = " UltraSPARC IIi ";
   if (StoreFamily == 87 ) CPUFamily = " UltraSPARC III ";
   if (StoreFamily == 88 ) CPUFamily = " UltraSPARC IIIi ";
   if (StoreFamily == 96 ) CPUFamily = " 68040 ";
   if (StoreFamily == 97 ) CPUFamily = " 68xxx Family ";
   if (StoreFamily == 98 ) CPUFamily = " 68000 ";
   if (StoreFamily == 99 ) CPUFamily = " 68010 ";
   if (StoreFamily == 100 ) CPUFamily = " 68020 ";
   if (StoreFamily == 101 ) CPUFamily = " 68030 ";
   if (StoreFamily == 112 ) CPUFamily = " Hobbit Family ";
   if (StoreFamily == 120 ) CPUFamily = " Crusoe™ TM5000 Family ";
   if (StoreFamily == 121 ) CPUFamily = " Crusoe™ TM3000 Family ";
   if (StoreFamily == 128 ) CPUFamily = " Weitek ";
   if (StoreFamily == 130 ) CPUFamily = " Itanium™ Processor ";
   if (StoreFamily == 144 ) CPUFamily = " PA-RISC Family ";
   if (StoreFamily == 145 ) CPUFamily = " PA-RISC 8500 ";
   if (StoreFamily == 146 ) CPUFamily = " PA-RISC 8000 ";
   if (StoreFamily == 147 ) CPUFamily = " PA-RISC 7300LC ";
   if (StoreFamily == 148 ) CPUFamily = " PA-RISC 7200 ";
   if (StoreFamily == 149 ) CPUFamily = " PA-RISC 7100LC ";
   if (StoreFamily == 150 ) CPUFamily = " PA-RISC 7100 ";
   if (StoreFamily == 160 ) CPUFamily = " V30 Family ";
   if (StoreFamily == 176 ) CPUFamily = " Pentium® III Xeon™ ";
   if (StoreFamily == 177 ) CPUFamily = " Pentium® III Processor with Intel® SpeedStep™ Technology ";
   if (StoreFamily == 178 ) CPUFamily = " Pentium® 4 ";
   if (StoreFamily == 179 ) CPUFamily = " Intel® Xeon™ ";
   if (StoreFamily == 180 ) CPUFamily = " AS400 Family ";
   if (StoreFamily == 181 ) CPUFamily = " Intel® Xeon™ processor MP ";
   if (StoreFamily == 182 ) CPUFamily = " AMD AthlonXP™ Family ";
   if (StoreFamily == 183 ) CPUFamily = " AMD AthlonMP™ Family ";
   if (StoreFamily == 184 ) CPUFamily = " Intel® Itanium® 2 ";
   if (StoreFamily == 185 ) CPUFamily = " AMD Opteron™ Family ";
   if (StoreFamily == 190 ) CPUFamily = " K7 ";
   if (StoreFamily == 200 ) CPUFamily = " IBM390 Family ";
   if (StoreFamily == 201 ) CPUFamily = " G4 ";
   if (StoreFamily == 202 ) CPUFamily = " G5 ";
   if (StoreFamily == 250 ) CPUFamily = " i860 ";
   if (StoreFamily == 251 ) CPUFamily = " i960 ";
   if (StoreFamily == 260 ) CPUFamily = " SH-3 ";
   if (StoreFamily == 261 ) CPUFamily = " SH-4 ";
   if (StoreFamily == 280 ) CPUFamily = " ARM ";
   if (StoreFamily == 281 ) CPUFamily = " StrongARM ";
   if (StoreFamily == 300 ) CPUFamily = " 6x86 ";
   if (StoreFamily == 301 ) CPUFamily = " MediaGX ";
   if (StoreFamily == 302 ) CPUFamily = " MII ";
   if (StoreFamily == 320 ) CPUFamily = " WinChip ";
   if (StoreFamily == 350 ) CPUFamily = " DSP ";
   if (StoreFamily == 500 ) CPUFamily = " Video Processor ";
   TotalMessage += (CPUFamily  + cr 
   + cpuInstance.Manufacturer + " " + cpuInstance.Caption + cr 
   + "Current Clock Speed: " + cpuInstance.CurrentClockSpeed + " MHz" + cr 
   + "Maximum Clock Speed: " + cpuInstance.MaxClockSpeed + " MHz" +cr 
   + "LoadPercentage: " + cpuInstance.LoadPercentage + "%" + cr + cr );
   counter += 1;
   }
return (TotalMessage);
}

function GetNIC()
{
   var Locator = new ActiveXObject ("WbemScripting.SWbemLocator");
   var adapter = Locator.connectserver();
   var TotalMessage = "\r\n"; var cr = "\r\n";

   try
   {
	var adapters = adapter.Get ("Win32_NetworkAdapterConfiguration");
        var iadapter = new Enumerator (adapters.Instances_());
   }
   catch ( e )
   {
        return ("Network Card: " + e.description);
   }

   for (;!iadapter.atEnd();iadapter.moveNext())
   {
       var NIC = iadapter.item();
          if (NIC.IPEnabled == 1)
	  { 
     	     TotalMessage += "Network Card " + (NIC.Index + 1) + cr;	
	     TotalMessage += "MAC: " + NIC.MACAddress + cr;	     	     
	     TotalMessage += "IP: " +NIC.IPAddress(0) + cr;
     	     TotalMessage += "Mask: " +NIC.IPSubnet(0) + cr;
     	     TotalMessage += "Gateway: " + NIC.DefaultIPGateway(0) + cr;
	     TotalMessage += "DHCP Server:  " + NIC.DHCPServer + cr;
	     TotalMessage += "DNS Domain: " + NIC.DNSDomain + cr;
	     TotalMessage += "DNS Domain Search Order: " + GetDNSSearch() + cr; 	
	     TotalMessage += "DNS Enabled For WINS Resolution: " + NIC.DNSEnabledForWINSResolution + cr;
	     TotalMessage += "DNS Hostname:  " + NIC.DNSHostname + cr;
 	     TotalMessage += "DNS Server Search Order: " + NIC.DNSServerSearchOrder(0) + cr;
	     TotalMessage += "DNS Server Search Order: " + NIC.DNSServerSearchOrder(1) + cr;
	     TotalMessage += "Domain DNS Registration Enabled:  " + NIC.DomainDNSRegistrationEnabled + cr;
	     TotalMessage += "WINS Primary Server:  " + NIC.WINSPrimaryServer + cr;
	     TotalMessage += "WINS Secondary Server: " + NIC.WINSSecondaryServer + cr + cr;
        }	 
   }
return (TotalMessage);
}

function GetDNSSearch()
{
    	var TotalMessage = "";
	try
    { 
        var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\System\\CurrentControlSet\\Services\\Tcpip\\Parameters\\SearchList";
        TotalMessage = shell.RegRead( regkey1 ) ;
	return( TotalMessage );
    }
    catch( e )
    {
        return( "" );
    }
}

function GetOTMInfo()
{
    try
    { 
        var cr = "\r\n";
	var TotalMessage = "";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Nortel\\NorMat\\SMP\\Navigator\\Version";
	TotalMessage += "OTM Version: ";
	TotalMessage += (shell.RegRead( regkey1 ) + cr );
	return( TotalMessage + cr );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}









