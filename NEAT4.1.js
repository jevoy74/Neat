/**********************************************************************************************************
 *    Name:       NEAT4.1.js 
 *    Version:    4.1
 *    Author:     Jason Evoy - jevoy@nortel.com
 *    Date:       July 16, 2009
 *    Notes:	  This is a script that polls the registry & WMI for PC data & outputs to a file
 *    Min Req:    Windows 2000, 2003, XP, Internet Explorer 5.5 
 **********************************************************************************************************/
var TotalMessage = "";
NetStat();
NetRoute();
PopupContinue();
MakeFile();
PopupDone();
CLEANUP();

function PopupContinue()
{
   var InfoIcon = 64; var OkButton = 0; var Timeout = 15;
   var Title = "Nortel Enterprise Audit Tool Ver 4.1                       - Jason Evoy";
   var MessageLine1 = "Creating log file...........................\n\nPlease wait until finished message appears before opening file\nThe audit usually takes less than 1 minute to complete";  
   var WshShell = WScript.CreateObject("WScript.Shell");
   var ButtonCode = WshShell.Popup(MessageLine1, Timeout, Title, OkButton + InfoIcon);
}

function PopupDone()
  {
    var OkButton = 0; var Timeout = 15; var InfoPic = 64;
    var Title = "Nortel Enterprise Audit Tool Ver 4.1                     - Jason Evoy";
    var shell = WScript.CreateObject("WScript.Shell");
    var DialogBox =  shell.Popup("Completed.  Log file saved as: " + GetComputerName() + ".txt", Timeout, Title, OkButton + InfoPic);
    if (DialogBox == OkButton)
       {
        WScript.Quit();
       }
  }  

function GetComputerName()
{
   var cr = "\r\n";
   var WSHshell = WScript.CreateObject("WScript.Network");
   var ComputerName = "";
   ComputerName += (WSHshell.ComputerName);
   return ( ComputerName );
}

function MakeFile()
   {
   var forWriting = 2, forAppending = 8, forReading = 1, cr = "\r\n";
   var shell = new ActiveXObject( "Scripting.FileSystemObject" );
   shell.CreateTextFile( GetComputerName()+ ".txt" );
   var os = shell.GetFile( GetComputerName() + ".txt" );
   os = os.OpenAsTextStream( forAppending, 0 );
   var LineBreak = "+------------------------------------------------------------------------------+";
   os.write(LineBreak + cr);
   os.write("¦Nortel Enterprise Audit Tool Results - ");
   os.write(Date());
   os.write(cr);
   os.write(LineBreak + cr +cr);   
   os.write(GetICCMINFO());
   os.write(GetCC6Info());
   os.write(GetSybaseVer());
   os.write(GetSybaseSECC());
   os.write(GetVoiceServices());
   os.write(GetNBNMINFO());
   os.write(GetCORES());
   os.write(GetM1Info());
   os.write(GetM1Serial());
   os.write(GetAccessLink());
   os.write(GetSWCINFO());
   os.write(GetCCMA6Info());
   os.write(GetLMInfo());
   os.write(GetREPINFO());
   os.write(GetREPIP());
   os.write(GetCCMS60PEPS());	
   os.write(GetServerUtilityPEPS());
   os.write(GetLMPEPS());
   os.Write(GetCCMAPEPS());
   os.Write(GetCPVer());
   os.Write(GetCPCLAN());
   os.Write(GetCPELAN());
   os.write(GetCPPEPS());
   os.write(GetCCMMPEPS());
   os.write(GetCCTPEPS());
   os.write(cr);
   os.write(GetMulticast());
   os.write(LineBreak + cr);
   os.write("¦Operating System Info:                                                        ¦" + cr);
   os.write(LineBreak + cr + cr);
   os.write(GetOS_NEW());
   os.write(GetOS());
   os.write(GetChasis());
   os.write(cr);
   os.write(chkpagefile());
   os.write(ShowNetwork());
   os.write(cr);
   os.write(GetMasterBrowser());
   os.write(GetDAYLIGHTSAVINGSTIME());
   os.write(GetMEDIASENSING());
   os.write(GetIPFORWARDING());
   os.write(IIS_Inst());
   os.write(GetAPIPA());
   os.write(GetNIC());
   os.write(chkcpu());
   os.write(ShowDriveList());
   os.write(cr);
   os.write(GetPATHLENGTH());  
   os.write(GetPATH());
   os.write(cr);
   os.write(LineBreak + cr);
   os.write("¦Running Processes                                                             ¦" + cr);
   os.write(LineBreak + cr + cr);
   os.write(GetPROCESS() +cr );
   os.write(LineBreak + cr);
   os.write("¦Services                                                                      ¦" + cr);
   os.write(LineBreak + cr + cr);
   os.write(chkservices());
   os.write(LineBreak + cr);
   os.write("¦Installed Software (Through Windows Installer)                                ¦" + cr);
   os.write(LineBreak + cr + cr);
   os.write(GetInstalledPrograms() +cr);	
   os.write(LineBreak + cr);
   os.write("¦Registry Run Keys                                                             ¦" + cr);
   os.write(LineBreak + cr + cr);
   TotalMessage = "";
   os.write(GetRunKey());
   os.write(cr);
   os.write(LineBreak + cr);
   os.write("¦HOSTS File (Comments Excluded)                                                ¦" + cr);
   os.write(LineBreak + cr);
   os.write(PrintHOSTS() + cr);
   os.write(LineBreak + cr);
   os.write("¦LMHOSTS File (Comments Excluded)                                              ¦" + cr);
   os.write(LineBreak + cr);
   os.write(PrintLMHOSTS());
   os.write(LineBreak + cr);
   os.write("¦TCP/IP Connections                                                            ¦" + cr);
   os.write(LineBreak + cr);
   os.write(PrintFile("C:\\netstat.log"));
   os.write(PrintFile("C:\\netstat.rte"));
   os.write(cr + cr + cr);
   os.write(LineBreak + cr);
   os.write("¦Application Event Log (Last 25 Events)                                        ¦" + cr);
   os.write(LineBreak + cr + cr + cr);
   os.write(GetAppEvents() + cr + cr + cr);  
   os.write(LineBreak + cr);
   os.write("¦System Event Log (Last 25 Events)                                             ¦" + cr);
   os.write(LineBreak + cr + cr);
   os.write(GetSysEvents());
   os.Close();
   }

function CLEANUP()
{
   try
   {   
   var shell = new ActiveXObject( "Scripting.FileSystemObject" );
   var PINGFILE = "C:\\temp.txt";
   var NetStatFile = "C:\\netstat.log";
   var NetStatFile2 = "C:\\netstat.rte";
   if (shell.FileExists(PINGFILE)) { shell.DeleteFile(PINGFILE);}
   if (shell.FileExists(NetStatFile)) {shell.DeleteFile(NetStatFile); }
   if (shell.FileExists(NetStatFile2)) {shell.DeleteFile(NetStatFile2); }
   }
   catch ( e )
   {
   }
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
	var regkey4 = "HKLM\\Software\\Nortel\\Setup\\NINAMESERVER\\SiteName";
	var regkey6 = "HKLM\\System\\CurrentControlSet\\Services\\NBNM_Service\\ImagePath";
        TotalMessage += "Site Name: ";
	TotalMessage += (shell.RegRead( regkey4 ) + cr);
	TotalMessage += "NBNM IP: ";
	TotalMessage += (shell.RegRead( regkey3 ) + " Path:" + shell.RegRead( regkey6 ) + cr + cr);	
        TotalMessage += "CLAN IP: ";
	IP = (shell.RegRead( regkey1 ));
        TotalMessage += IP + cr;
	TotalMessage += (PingIPs(IP)) + cr + cr;
	TotalMessage += "ELAN IP: ";
	IP = (shell.RegRead( regkey2 ));
	TotalMessage += IP + cr;
	TotalMessage += (PingIPs(IP)) + cr + cr;
	return( TotalMessage );
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
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetREPINFO()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Replication Server Name: ";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Nortel\\Setup\\RepServerName";
 	var regkey2 = "HKLM\\Software\\Nortel\\Setup\\RepServerType";
  	TotalMessage += (shell.RegRead( regkey1 ) + cr);
	TotalMessage += "Warm Standby Type: ";
	TotalMessage += (shell.RegRead( regkey2 ) + cr );
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

function GetM1Info()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey2 = "HKLM\\Software\\Nortel\\Setup\\Switch\\Name";
	var regkey3 = "HKLM\\Software\\Nortel\\Setup\\Switch\\IP";
 	var regkey4 = "HKLM\\Software\\Nortel\\Setup\\Switch\\Customer";
	TotalMessage += ("CS 1000 Customer: " + shell.RegRead( regkey4 ) + cr);
	TotalMessage += ("CS 1000 HostName: " + shell.RegRead( regkey2 ) + cr);
	TotalMessage += ("CS 1000 IP: ");
	IP = (shell.RegRead( regkey3 ));
        TotalMessage += IP + cr;
	TotalMessage += (PingIPs(IP)) + cr + cr;
	
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetM1Serial()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = ""+cr;
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Nortel\\Setup\\Switch\\Serial";
	TotalMessage += ("CS 1000 Serial: " + shell.RegRead( regkey1 ) +cr);
	return( TotalMessage );
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

function GetOS_NEW()
{
   var TotalMessage = "", cr = "\r\n";
   var wbemFlagReturnImmediately = 0x10;
   var wbemFlagForwardOnly = 0x20;
   
   try
   {
  
   var objWMIService = GetObject("winmgmts:\\\\.\\root\\CIMV2");
   var colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem", "WQL", wbemFlagReturnImmediately | wbemFlagForwardOnly);
   var enumItems = new Enumerator(colItems);
   for (; !enumItems.atEnd(); enumItems.moveNext())
      {
      var objItem = enumItems.item();
      

	TotalMessage +="Operating System: \t" + objItem.Caption + " " + cr;
	TotalMessage +="CDS Version: \t\t"    + objItem.CSDVersion    + cr;
	TotalMessage +="Serial Number: \t\t"  + objItem.SerialNumber  + cr;
	TimeGenerated = objItem.InstallDate;
	TotalMessage += "Install Date: \t\t" + (ConvertINTtoTime(TimeGenerated)) + cr;
	TotalMessage +="Licensed Users: \t"      + objItem.NumberOfLicensedUsers + cr;
	TotalMessage +="Encryption Level: \t"      + objItem.EncryptionLevel + " bit" + cr;
	TimeGenerated = objItem.LocalDateTime;
	TotalMessage += "Local Time: \t\t" + (ConvertINTtoTime(TimeGenerated)) + cr;
	TimeGenerated = objItem.LastBootUpTime;
	TotalMessage += "Last Bootup Time: \t" + (ConvertINTtoTime(TimeGenerated)) + cr;
	TotalMessage +="Total Physical Memory: \t" + objItem.TotalVisibleMemorySize + " KB" + cr;
        TotalMessage +="Free Physical Memory: \t"  + objItem.FreePhysicalMemory + " KB" + cr;
        TotalMessage +="Total Virtual Memory: \t"  + objItem.TotalVirtualMemorySize + " KB" + cr;
        TotalMessage +="Free Virtual Memory: \t"   + objItem.FreeVirtualMemory + " KB" + cr;
      }

   }
   catch ( e )
   {
      return ("Could not get Operating System info: " + e.description);
   }

   
return (TotalMessage);
}

function ConvertINTtoTime(TimeGenerated)
{
    var TotalMessage = "";

    var ReadableDate = parseInt(TimeGenerated);    // Change DateTime Variable to an Integer
    var TheYear = Math.floor(ReadableDate / 10000000000);
    ReadableDate = (ReadableDate - (TheYear * 10000000000));
    var TheMonth = Math.floor(ReadableDate / 100000000);
    TotalMessage += TheYear + "-";
    if (TheMonth < 10) { TotalMessage+= "0" }          // So it will show a 08 instead of 8
    TotalMessage += TheMonth + "-";
    ReadableDate = (ReadableDate - (TheMonth * 100000000));
    var TheDay = Math.floor(ReadableDate / 1000000);
    if (TheDay < 10) { TotalMessage+= "0" }
    TotalMessage += TheDay + " ";
    ReadableDate = (ReadableDate - (TheDay * 1000000));
    var TheHour = Math.floor(ReadableDate / 10000);
    if (TheHour < 10) { TotalMessage+= "0" }
    TotalMessage += TheHour + ":";
    ReadableDate = (ReadableDate - (TheHour * 10000));
    var TheMinutes = Math.floor(ReadableDate / 100);
    if (TheMinutes < 10) { TotalMessage+= "0" }
    TotalMessage += TheMinutes + ":";
    ReadableDate = (ReadableDate - (TheMinutes * 100));
    var TheSeconds = ReadableDate;
    if (TheSeconds < 10) { TotalMessage+= "0" }
    TotalMessage += TheSeconds + "  ";

    return(TotalMessage);
} 

function GetOS()
{
    try
    {
        var cr = "\r\n"; var TotalMessage = ""; 
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey4 = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Hostname";
	var regkey5 = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Domain";
        TotalMessage += ("Computer Name: \t\t" + GetComputerName()+ cr);
	TotalMessage += ("Domain: \t\t" + shell.RegRead( regkey5 ) + cr);
	TotalMessage += "Hostname: \t\t";
	IP = (shell.RegRead( regkey4 ));
        TotalMessage += IP + cr;
	TotalMessage += (PingIPs(IP)) + cr;
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = e;
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
   var Locator = new ActiveXObject ("WbemScripting.SWbemLocator");
   var drive = Locator.connectserver();
   var TotalMessage = ""; var cr = "\r\n"; var OneMeg = "1048576"; var MB; var DiskDrive = ""; var tab ="\t\t";

   try
   {
	var drives = drive.Get ("Win32_LogicalDisk");
        var DriveInstance = new Enumerator (drives.Instances_());
        TotalMessage += "Drive Information:" + cr;
    for (;!DriveInstance.atEnd();DriveInstance.moveNext())
   {
        var iDrive = DriveInstance.item();
        TotalMessage += iDrive.Name;	
	DiskDrive = iDrive.DriveType ;
        if (DiskDrive == 1) {TotalMessage += " - No Root Directory "};
        if (DiskDrive == 2) {TotalMessage += " - Removable Disk "};
        if (DiskDrive == 3) {TotalMessage += " - Local Disk "};
        if (DiskDrive == 4) {TotalMessage += " - Network Drive "};
        if (DiskDrive == 5) {TotalMessage += " - Compact Disk "};
        if (DiskDrive == 6) {TotalMessage += " - RAM Disk "}; 
	TotalMessage += "(" + iDrive.FileSystem + ")" + tab;
	MB = (iDrive.Size / OneMeg);
      	MB = (MB * 10);
      	MB = Math.round(MB);
      	MB = (MB / 10);	
      	MB += " MB";
      	TotalMessage += MB + tab;
        MB = (iDrive.FreeSpace / OneMeg);
      	MB = (MB * 10);
      	MB = Math.round(MB);
      	MB = (MB / 10);	
      	MB += " MB Free";
      	TotalMessage += MB;
  	TotalMessage += cr;
   }
return (TotalMessage);
  }
   catch ( e )
   {
        return ("");
   }

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
      return ("Could not get Services info because: " + error.description);
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
	     TotalMessage += PRO.Name + " " + tab;
	     TotalMessage += "PID: " + PRO.ProcessID + tab;       
 	     TotalMessage += "Handles: " + PRO.HandleCount + tab;
	     TotalMessage += "Threads: " + PRO.ThreadCount + tab;
	     TotalMessage += "Memory: " + ((PRO.WorkingSetSize / 1024) + "k");
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
		TotalMessage += " " + shell.RegRead( regkey3 ) + " baud" + cr + cr;
           }
	if (TCPIP == 1)
	   {   		
		IP = shell.RegRead( regkey2 )
		TotalMessage += IP;
		TotalMessage += " port " + shell.RegRead( regkey1 ) + cr;
		TotalMessage += (PingIPs(IP)) + cr + cr;
	   }	
        return( TotalMessage );
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
        TotalMessage ="";
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
      return ("Could not get pagefile info because: " + error.description);
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
      TotalMessage += cpuInstance.DeviceID + cr;
      TotalMessage += cpuInstance.Name + cr;
      TotalMessage += "Current Clock Speed: " + cpuInstance.CurrentClockSpeed + " MHz" + cr ;
      TotalMessage += "Maximum Clock Speed: " + cpuInstance.MaxClockSpeed + " MHz" +cr ;
      TotalMessage += "LoadPercentage: " + cpuInstance.LoadPercentage + "%" + cr + cr;
   }
return (TotalMessage + cr);
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
          if ((NIC.IPEnabled == 1) && (NIC.IPAddress(0) != ""))
	  { 
     	     TotalMessage += "Network Card " + (NIC.Index + 1) + cr;	
	     TotalMessage += "MAC: " + NIC.MACAddress + cr;	     	     
	     TotalMessage += "IP: " +NIC.IPAddress(0) + cr;
     	     TotalMessage += "Mask: " +NIC.IPSubnet(0) + cr;
     	     TotalMessage += "Gateway: " + NIC.DefaultIPGateway(0) + cr;
	     TotalMessage += "DHCP Server:  " + NIC.DHCPServer + cr;
	     TotalMessage += "DNS Domain: " + NIC.DNSDomain + cr;
	     TotalMessage += "DNS Domain Search Order: " + GetDNSSearch() + cr; 	
	     TotalMessage += "DNS Enabled For WINS Resolution: ";
 	     if ((NIC.DNSEnabledForWINSResolution == true)) { TotalMessage += ("Yes" + cr);}
	     else {TotalMessage += ("No" + cr);}
	     TotalMessage += "DNS Hostname:  " + NIC.DNSHostname + cr;
 	     TotalMessage += "DNS Server Search Order: " + NIC.DNSServerSearchOrder(0) + cr;
	     TotalMessage += "Domain DNS Registration Enabled: " ;
	     if ((NIC.DomainDNSRegistrationEnabled == true)) { TotalMessage += ("Yes" + cr);}
	     else {TotalMessage += ("No" + cr);}
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

function GetVoiceServices()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Voice Services Card: ";
	TotalMessage += cr;
	TotalMessage += "VS Card Name: ";
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey1 = "HKLM\\Software\\Nortel\\SECCVPS\\Switch 1\\VPS_CARD_1\\Card_Name";
	var regkey2 = "HKLM\\Software\\Nortel\\SECCVPS\\Switch 1\\VPS_CARD_1\\ipAddress";
	var regkey3 = "HKLM\\Software\\Nortel\\SECCVPS\\Switch 1\\VPS_CARD_1\\Keycode";
	var regkey4 = "HKLM\\Software\\Nortel\\SECCVPS\\Switch 1\\VPS_CARD_1\\Transfer_DN";
	var regkey7 = "HKLM\\Software\\Nortel\\SECCVPS\\TapiServer\\ipAddress";
	var regkey8 = "HKLM\\Software\\Nortel\\SECCVPS\\TapiServer\\PortNumber";
	var regkey9 = "HKLM\\Software\\Nortel\\SECCVPS\\TapiServer\\TapiInstalled";
	TotalMessage += (shell.RegRead( regkey1 ));
	TotalMessage += ( cr + "VS Card IP Address: ");
	IP = (shell.RegRead( regkey2 ));
        TotalMessage += IP + cr;
	TotalMessage += (PingIPs(IP));	
	TotalMessage += ( cr + "VS Card Keycode: ");
	TotalMessage += (shell.RegRead( regkey3 ));
	TotalMessage += ( cr + "VS Default DN: ");
	TotalMessage += (shell.RegRead( regkey4 ));
	TotalMessage += ( cr + "TAPI Installed: ");
	if ((shell.RegRead( regkey9 )) == 1) { TotalMessage += ("Yes");}
	else {TotalMessage += ("No");}
	TotalMessage += ( cr + "TAPI CLAN IP: ");
	IP = (shell.RegRead( regkey7 ));
        TotalMessage += IP + cr;
	TotalMessage += (PingIPs(IP));
	TotalMessage += ( cr + "TAPI Port Address: ");
	TotalMessage += (shell.RegRead( regkey8 ));
	return( TotalMessage + cr + cr );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetSybaseVer()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Sybase Version: ";
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey1 = "HKLM\\Software\\SYBASE\\SQLServer\\CurrentVersion";
	var regkey2 = "HKLM\\Software\\ODBC\\ODBC.INI\\Blue\\NetworkAddress";
	var regkey3 = "HKLM\\Software\\SYBASE\\SQLServer\\DSLISTEN";
	TotalMessage += (shell.RegRead( regkey1 ));
	TotalMessage += ( cr + "SQL Listening On: " );
	TotalMessage += (shell.RegRead( regkey3 ));
	TotalMessage += ( cr + "ODBC Blue Address: " );
	TotalMessage += (shell.RegRead( regkey2 ));
	return( TotalMessage + cr );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetCORES()
{
    try
    { 
        var cr = "\r\n";
	var TotalMessage = "";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Nortel\\Co-Residency\\Symposium Web Client";
	var regkey2 = "HKLM\\Software\\Nortel\\Co-Residency\\TAPI";
       
        TotalMessage += "Co-Res SWC: ";
	TotalMessage += (shell.RegRead( regkey1 ) + cr );
	TotalMessage += "Co-Res TAPI: ";
	TotalMessage += (shell.RegRead( regkey2 ) + cr );
	
	return( TotalMessage + cr);
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function PingIPs(IP)
{
   var TotalMessage = "";
   var cr = "\r\n";
   var forReading = 1, forWriting = 2, forAppending = 8;
   var Wshell = new ActiveXObject("WScript.Shell");
   var shell = new ActiveXObject("Scripting.FileSystemObject");
   
   var command = "cmd.exe /c ping.exe -a -n 1 -w 1000 "
   command += IP;
   command += " >C:\\temp.txt"
   Wshell.run (command, 0, true);
   var os = shell.GetFile("C:\\temp.txt");
   os = os.OpenAsTextStream( forReading, 0 );
   while( !os.AtEndOfStream )
   {
      var sReadLine = os.ReadLine();
      PingResponse = sReadLine.substr(0,5)

      if (PingResponse == "Pingi")
      {
         TotalMessage += (sReadLine.substr(0) + cr);
      }
      if (PingResponse == "Reply")
      {
         TotalMessage += (sReadLine.substr(0));
      }
      if (PingResponse == "Reque")
      {
         TotalMessage += (sReadLine.substr(0));
      }	
	if (PingResponse == "Desti")
      {
         TotalMessage += (sReadLine.substr(0));
      }	
      if (PingResponse == "Unkno")
      {
         TotalMessage += (sReadLine.substr(0));
      }	 	
   }

    return(TotalMessage);
} 

function PrintHOSTS()
{
   try
   {
   var TotalMessage = "";
   var cr = "\r\n";
   var forReading = 1, forWriting = 2, forAppending = 8;
   var shell = new ActiveXObject("Scripting.FileSystemObject");
   var wshell = new ActiveXObject("WScript.Shell");
   var regkey1 = "HKLM\\Software\\Microsoft\\Windows NT\\CurrentVersion\\SystemRoot";
   var HOSTPATH = wshell.RegRead(regkey1);
   HOSTPATH += "\\system32\\drivers\\ETC\\hosts." 
   var os = shell.GetFile(HOSTPATH);
   os = os.OpenAsTextStream( forReading, 0 );
   while( !os.AtEndOfStream )
   {
      var sReadLine = os.ReadLine();
      CommentCheck = sReadLine.substr(0,1);
 
      if (CommentCheck != "#")
      {
         TotalMessage += (sReadLine.substr(0));
	 TotalMessage += cr;
      }
   }
}
    catch( e )
    {
        return( "" );
    }
   return(TotalMessage);
} 

function PrintLMHOSTS()
{
   try
   {
   var TotalMessage = "";
   var cr = "\r\n";
   var forReading = 1, forWriting = 2, forAppending = 8;
   var shell = new ActiveXObject("Scripting.FileSystemObject");
   var wshell = new ActiveXObject("WScript.Shell");
   var regkey1 = "HKLM\\Software\\Microsoft\\Windows NT\\CurrentVersion\\SystemRoot";
   var HOSTPATH = wshell.RegRead(regkey1);
   HOSTPATH += "\\system32\\drivers\\ETC\\lmhosts.sam" 
   var os = shell.GetFile(HOSTPATH);
   os = os.OpenAsTextStream( forReading, 0 );
   while( !os.AtEndOfStream )
   {
      var sReadLine = os.ReadLine();
      CommentCheck = sReadLine.substr(0,1);
 
      if (CommentCheck != "#")
      {
         TotalMessage += (sReadLine.substr(0));
	 TotalMessage += cr;
      }
   }
   }
    catch( e )
    {
        return( "" );
    }
   return(TotalMessage);
} 

function GetSybaseSECC()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Sybase Version: ";
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey1 = "HKLM\\Software\\SYBASE\\SQLServer\\CurrentVersion";
	var regkey2 = "HKLM\\Software\\ODBC\\ODBC.INI\\Blue\\ServerName";
	var regkey3 = "HKLM\\Software\\SYBASE\\SQLServer\\DSLISTEN";
	TotalMessage += (shell.RegRead( regkey1 ));
	TotalMessage += ( cr + "SQL Listening On: " );
	TotalMessage += (shell.RegRead( regkey3 ));
	TotalMessage += ( cr + "ODBC Blue Address: " );
	TotalMessage += (shell.RegRead( regkey2 ));
	return( TotalMessage + cr );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetMasterBrowser()
{
    	var TotalMessage = "Maintain Server List (Browser Election) is set to ";
	var cr = "\r\n";
	try
    { 
        var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Browser\\Parameters\\MaintainServerList";
        TotalMessage += shell.RegRead( regkey1 ) ;
	return( TotalMessage + cr );
    }
    catch( e )
    {
        return( "" );
    }
}

function GetCC6Info()
{
    try
    {
        var cr = "\r\n";
        var LineBreak = "+------------------------------------------------------------------------------+";
        TotalMessage = LineBreak + cr;
        TotalMessage += "¦Contact Center Info:                                                          ¦" + cr;
        TotalMessage += LineBreak + cr + cr;
	
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey2 = "HKLM\\Software\\Nortel\\Setup\\SerialNumber";
  	var regkey3 = "HKLM\\Software\\Nortel\\Setup\\Platform Version";
	var regkey4 = "HKLM\\Software\\Nortel\\Setup\\Build";
  	var regkey5 = "HKLM\\Software\\Nortel\\Setup\\UserCompany";
	var regkey6 = "HKLM\\Software\\Nortel\\Setup\\UserName";
	var regkey8 = "HKLM\\Software\\Nortel\\Setup\\InstallTimeSU";
	TotalMessage += "Version: ";
	TotalMessage += (shell.RegRead( regkey3 ) );
	TotalMessage += " Build: ";
	TotalMessage += (shell.RegRead( regkey4 ) + cr );
	TotalMessage += "Install Time PEP: ";
	TotalMessage += (shell.RegRead( regkey8 ) + cr );
	TotalMessage += "Serial Number: ";
	TotalMessage += (shell.RegRead( regkey2 ) + cr );
	TotalMessage += "Company Name: ";
	TotalMessage += (shell.RegRead( regkey5 ) + cr );
	TotalMessage += "Customer Name: ";
	TotalMessage += (shell.RegRead( regkey6 ) + cr );
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetLMInfo()
{
    try
    {
        var cr = "\r\n";
	var LineBreak = "+------------------------------------------------------------------------------+";
        TotalMessage = LineBreak + cr;
        TotalMessage += "¦License Manager Info:                                                         ¦" + cr;
        TotalMessage += LineBreak + cr + cr;
        var IP
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Nortel\\LM\\Type";
       	var regkey2 = "HKLM\\Software\\Nortel\\Setup\\LM_IP_PORT_Primary";
  	var regkey3 = "HKLM\\Software\\Nortel\\Setup\\LM_IP_PORT_Secondary";
	var regkey4 = "HKLM\\Software\\Nortel\\Setup\\LM_IP_Primary";
  	var regkey5 = "HKLM\\Software\\Nortel\\Setup\\LM_IP_Secondary";
	var regkey6 = "HKLM\\Software\\Nortel\\Setup\\LM_Package";
	var regkey7 = "HKLM\\Software\\Nortel\\Setup\\LM_StandbyServer";
	var regkey8 = "HKLM\\Software\\Nortel\\Setup\\LM_OpenQueue";
	var regkey9 = "HKLM\\Software\\Nortel\\LM\\Server\\LicFile";
	var regkey10= "HKLM\\Software\\Nortel\\Setup\\SerialNumber";
	
       
	TotalMessage += "Package: ";
	TotalMessage += (shell.RegRead( regkey6 ) + cr );
	TotalMessage += "Type: ";
	TotalMessage += (shell.RegRead( regkey1 ) + cr );
        TotalMessage += "Serial Number: ";
	TotalMessage += (shell.RegRead( regkey10 ) + cr );
        TotalMessage += "License File: ";
	TotalMessage += (shell.RegRead( regkey9 ) + cr );
        TotalMessage += "Open Queue: ";
	TotalMessage += (shell.RegRead( regkey8 ) + cr );
	TotalMessage += "Standby Server: ";
	TotalMessage += (shell.RegRead( regkey7 ) + cr + cr );
	TotalMessage += "Primary IP: ";
	IP = (shell.RegRead( regkey4 ) );
	TotalMessage += IP + "\:";
	TotalMessage += (shell.RegRead( regkey2 ) + cr );
	TotalMessage += (PingIPs(IP)) + cr + cr;
	TotalMessage += "Secondary IP: ";
	IP = (shell.RegRead( regkey5 ) );
           if (IP == regkey5)  { IP = "Unconfigured" };
        TotalMessage += IP + "\:";
       	TotalMessage += (shell.RegRead( regkey3 ) + cr + cr);
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetREPIP()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Replication Server IP: ";
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey3 = "HKLM\\Software\\Nortel\\Setup\\RepServerIP";
	var regkey4 = "HKLM\\Software\\Nortel\\Setup\\RepServerPortNum";
	IP = (shell.RegRead( regkey3 ) );
	TotalMessage += IP + "\:";
	TotalMessage += (shell.RegRead( regkey4 ) + cr );
	TotalMessage += (PingIPs(IP)) + cr + cr;
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetCCMAPEPS()
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

	TotalMessage = "";
	cr = "\r\n", TAB = "\t";
	
	RegistryPath = "SOFTWARE\\Nortel\\WClient\\Records"

	try
	{
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
	   GetPatchKey (KeyName); 
	   TotalMessage += "CCM Administration" + TAB + KeyName + cr;
   	} 
	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
return (TotalMessage);
}
 
function GetLMPEPS()
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

	TotalMessage = "";
	cr = "\r\n", TAB = "\t";
	
	RegistryPath = "SOFTWARE\\Nortel\\LM\\Server\\Records"

	try
	{
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
	   GetPatchKey (KeyName); 
	   TotalMessage += "License Manager" + TAB + KeyName + cr;
   	} 
	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
return (TotalMessage);
}


function GetServerUtilityPEPS()
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

	TotalMessage = "";
	cr = "\r\n", TAB = "\t";
	
	RegistryPath = "SOFTWARE\\Nortel\\Nortel SMI\\Records"

	try
	{
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
	   GetPatchKey (KeyName); 
	   TotalMessage += "Server Utility" + TAB + KeyName + cr;
   	} 
	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
return (TotalMessage);
}

function GetCCMS60PEPS()
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

        var LineBreak = "+------------------------------------------------------------------------------+";
        cr = "\r\n", TAB = "\t";
	TotalMessage = LineBreak + cr;
        TotalMessage +="¦Installed Patches                                                             ¦" + cr;
        TotalMessage += LineBreak + cr + cr;
	
	
	RegistryPath = "SOFTWARE\\Nortel\\SCCS\\Records"

	try
	{
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
	   GetPatchKey (KeyName); 
	   TotalMessage += "Contact Center" + TAB + KeyName + cr;
   	} 
	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
return (TotalMessage);
}

function GetPatchKey (KeyName)        
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
            if (subkeyval == "Date" || subkeyval == "Time" || subkeyval == "LogonUser")
            {
   	    oMethod = oReg.Methods_.Item("GetStringValue");
            oInParam = oMethod.InParameters.SpawnInstance_();
            oInParam.hDefKey = HKLM;
            oInParam.sSubKeyName = newreg;
            oInParam.sValueName = subkeyval;
            oOutParam = oReg.ExecMethod_(oMethod.Name, oInParam);
            cNames = oOutParam.sValue
            TotalMessage += cNames;
            TotalMessage += TAB;
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



function GetCPPEPS()
{
	HKLM = 0x80000002;
	HKCR = 0x80000000;
	HKCU = 0x80000001;
	HKUS = 0x80000003;
	HKCC = 0x80000005;

	var LineBreak = "+------------------------------------------------------------------------------+";
        cr = "\r\n", TAB = "\t";
	TotalMessage = LineBreak + cr;
        TotalMessage +="¦Installed Patches                                                             ¦" + cr;
        TotalMessage += LineBreak + cr + cr;
	
	RegistryPath = "SOFTWARE\\Nortel\\Setup\\CPServer\\"

	try
	{
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
	   GetCPPEPSoftwareKey(KeyName);
   	} 

	return (TotalMessage);
	}
  catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}
 
function GetCPPEPSoftwareKey (KeyName)
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
            if (subkeyval != "(Default)")
            {
   	    oMethod = oReg.Methods_.Item("GetStringValue");
            oInParam = oMethod.InParameters.SpawnInstance_();
            oInParam.hDefKey = HKLM;
            oInParam.sSubKeyName = newreg;
            oInParam.sValueName = subkeyval;
            TotalMessage += subkeyval;
            TotalMessage += cr;
            }
        }
   }
}
   catch (e)
   {
   } 
}

function NetStat()
{
   var Wshell = new ActiveXObject("WScript.Shell");
   var command = "cmd.exe /c netstat.exe -na>C:\\netstat.log"
   Wshell.run (command, 0, true);
   return;
}  

function NetRoute()
{
   var Wshell = new ActiveXObject("WScript.Shell");
   var FILENAME = "C:\\netstat.rte"
   var command = "cmd.exe /c netstat.exe -r >" + FILENAME
   Wshell.run (command, 0, true);
   return;
} 

function PrintFile(FILENAME)
{
   try
   {
   var TotalMessage = ""; 
   var cr = "\r\n";
   var forReading = 1, forWriting = 2, forAppending = 8;
   var shell = new ActiveXObject("Scripting.FileSystemObject");
   var os = shell.GetFile(FILENAME);
   os = os.OpenAsTextStream( forReading, 0 );
     while( !os.AtEndOfStream )
     {
	TotalMessage += os.ReadLine();
        TotalMessage += cr;
     }
   os.Close();
   }
   catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
   return(TotalMessage);
} 

function GetChasis()
{
   var Locator = new ActiveXObject ("WbemScripting.SWbemLocator");
   var chasis = Locator.connectserver();
   var TotalMessage = ""; var cr = "\r\n";

   try
   {
	var chasiss = chasis.Get ("Win32_SystemEnclosure");
        var ichasis = new Enumerator (chasiss.Instances_());
    for (;!ichasis.atEnd();ichasis.moveNext())
   {
        var CHAS = ichasis.item();
        TotalMessage += cr + "Computer Info:" + cr + "Make: " + CHAS.Manufacturer + cr;	
	FindType = CHAS.ChassisTypes(0) + cr;
 	if (FindType == 1)  { TotalMessage += "Type: Other" + cr };
	if (FindType == 3)  { TotalMessage += "Type: Desktop" + cr };
	if (FindType == 4)  { TotalMessage += "Type: Low Profile Desktop" + cr };
	if (FindType == 5)  { TotalMessage += "Type: Pizza Box" + cr };
	if (FindType == 6)  { TotalMessage += "Type: Mini Tower" + cr };
	if (FindType == 7)  { TotalMessage += "Type: Tower" + cr };
	if (FindType == 8)  { TotalMessage += "Type: Portable" + cr };
	if (FindType == 9)  { TotalMessage += "Type: Laptop" + cr };
	if (FindType == 10) { TotalMessage += "Type: Notebook" + cr };
	if (FindType == 11) { TotalMessage += "Type: Hand Held" + cr };
	if (FindType == 12) { TotalMessage += "Type: Docking Station" + cr };
	if (FindType == 13) { TotalMessage += "Type: All In One" + cr };
	if (FindType == 14) { TotalMessage += "Type: Sub Notebook" + cr };
	if (FindType == 15) { TotalMessage += "Type: Space-Saving" + cr };
	if (FindType == 16) { TotalMessage += "Type: Lunch Box" + cr };
	if (FindType == 17) { TotalMessage += "Type: Main System Chassis" + cr };
	if (FindType == 18) { TotalMessage += "Type: Expansion Chassis" + cr };
	if (FindType == 19) { TotalMessage += "Type: Sub Chassis" + cr };
	if (FindType == 20) { TotalMessage += "Type: Bus Expansion Chassis" + cr };
	if (FindType == 21) { TotalMessage += "Type: Peripheral Chassis" + cr };
	if (FindType == 22) { TotalMessage += "Type: Storage Chassis" + cr };
	if (FindType == 23) { TotalMessage += "Type: Rack Mount Chassis" + cr };
	if (FindType == 24) { TotalMessage += "Type: Sealed-Case PC" + cr };
   }
return (TotalMessage);
  }
   catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }

}

function GetAppEvents()
{

try
{
   var wbemFlagReturnImmediately = 0x10 ;
   var wbemFlagForwardOnly = 0x20 ;
   var TotalMessage = "" ;
   var EventCounter = 0;
   var NumberOfEvents = 25;
   var cr = "\r\n";
   var objWMIService = GetObject("winmgmts:\\\\.\\root\\CIMV2") ;
   var colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NTLogEvent where LogFile='Application'", "WQL",
                                          wbemFlagReturnImmediately | wbemFlagForwardOnly) ;
   var enumItems = new Enumerator(colItems) ;
   
while ((EventCounter != NumberOfEvents) && (!enumItems.atEnd()))
   {
      var EVT = enumItems.item() ;

    var ReadableDate = parseInt(EVT.TimeGenerated);    // Change DateTime Variable to an Integer
    var TheYear = Math.floor(ReadableDate / 10000000000);
    ReadableDate = (ReadableDate - (TheYear * 10000000000));
    var TheMonth = Math.floor(ReadableDate / 100000000);
    TotalMessage += TheYear + "-";
    if (TheMonth < 10) { TotalMessage+= "0" }          // So it will show a 08 instead of 8
    TotalMessage += TheMonth + "-";
    ReadableDate = (ReadableDate - (TheMonth * 100000000));
    var TheDay = Math.floor(ReadableDate / 1000000);
    if (TheDay < 10) { TotalMessage+= "0" }
    TotalMessage += TheDay + " ";
    ReadableDate = (ReadableDate - (TheDay * 1000000));
    var TheHour = Math.floor(ReadableDate / 10000);
    if (TheHour < 10) { TotalMessage+= "0" }
    TotalMessage += TheHour + ":";
    ReadableDate = (ReadableDate - (TheHour * 10000));
    var TheMinutes = Math.floor(ReadableDate / 100);
    if (TheMinutes < 10) { TotalMessage+= "0" }
    TotalMessage += TheMinutes + ":";
    ReadableDate = (ReadableDate - (TheMinutes * 100));
    var TheSeconds = ReadableDate;
    if (TheSeconds < 10) { TotalMessage+= "0" }
    TotalMessage += TheSeconds + "  ";
 
    TotalMessage += "Event: " + EVT.EventCode + ": " 
    var EventTypeCode = EVT.EventType; 
        if (EventTypeCode == 1 ) TotalMessage += " Error ";
        if (EventTypeCode == 2 ) TotalMessage += " Warning ";
        if (EventTypeCode == 3 ) TotalMessage += " Information ";
        if (EventTypeCode == 4 ) TotalMessage += " Audit Success ";
        if (EventTypeCode == 5 ) TotalMessage += " Audit Failure ";
   TotalMessage += "\tSource: " + EVT.SourceName;
   
	var CheckNULLMessage = EVT.Message; 
        if (CheckNULLMessage == null )
        {
         try { (TotalMessage += cr + "Description: " + EVT.InsertionStrings.toArray().join(",")); }
         catch(e) { return; }
        }
        if (CheckNULLMessage != null) { TotalMessage += cr + "Description: " + EVT.Message }; 
	TotalMessage += cr + "--------------------------------------------------------------------------------" + cr;
      enumItems.moveNext();
      EventCounter += 1;
   } 
}
catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
return (TotalMessage);
}

function GetSysEvents()
{

try
{
   var wbemFlagReturnImmediately = 0x10 ;
   var wbemFlagForwardOnly = 0x20 ;
   var TotalMessage = "" ;
   var EventCounter = 0;
   var NumberOfEvents = 25;
   var cr = "\r\n";
   var objWMIService = GetObject("winmgmts:\\\\.\\root\\CIMV2") ;
   var colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NTLogEvent where LogFile='System'", "WQL", wbemFlagReturnImmediately | wbemFlagForwardOnly) ;
   var enumItems = new Enumerator(colItems) ;
 while ((EventCounter != NumberOfEvents) && (!enumItems.atEnd()))
   {
    var EVT = enumItems.item() ;
    var ReadableDate = parseInt(EVT.TimeGenerated);    // Change DateTime Variable to an Integer
    var TheYear = Math.floor(ReadableDate / 10000000000);
    ReadableDate = (ReadableDate - (TheYear * 10000000000));
    var TheMonth = Math.floor(ReadableDate / 100000000);
    TotalMessage += TheYear + "-";
    if (TheMonth < 10) { TotalMessage+= "0" }          // So it will show a 08 instead of 8
    TotalMessage += TheMonth + "-";
    ReadableDate = (ReadableDate - (TheMonth * 100000000));
    var TheDay = Math.floor(ReadableDate / 1000000);
    if (TheDay < 10) { TotalMessage+= "0" }
    TotalMessage += TheDay + " ";
    ReadableDate = (ReadableDate - (TheDay * 1000000));
    var TheHour = Math.floor(ReadableDate / 10000);
    if (TheHour < 10) { TotalMessage+= "0" }
    TotalMessage += TheHour + ":";
    ReadableDate = (ReadableDate - (TheHour * 10000));
    var TheMinutes = Math.floor(ReadableDate / 100);
    if (TheMinutes < 10) { TotalMessage+= "0" }
    TotalMessage += TheMinutes + ":";
    ReadableDate = (ReadableDate - (TheMinutes * 100));
    var TheSeconds = ReadableDate;
    if (TheSeconds < 10) { TotalMessage+= "0" }
    TotalMessage += TheSeconds + "  ";
    TotalMessage += "Event: " + EVT.EventCode + ": " 
    var EventTypeCode = EVT.EventType; 
        if (EventTypeCode == 1 ) TotalMessage += " Error ";
        if (EventTypeCode == 2 ) TotalMessage += " Warning ";
        if (EventTypeCode == 3 ) TotalMessage += " Information ";
        if (EventTypeCode == 4 ) TotalMessage += " Audit Success ";
        if (EventTypeCode == 5 ) TotalMessage += " Audit Failure ";
   TotalMessage += "\tSource: " + EVT.SourceName;
   
	var CheckNULLMessage = EVT.Message; 
        if (CheckNULLMessage == null )
        {
         try { (TotalMessage += cr + "Description: " + EVT.InsertionStrings.toArray().join(",")); }
         catch(e) { return; }
        }
        if (CheckNULLMessage != null) { TotalMessage += cr + "Description: " + EVT.Message }; 
	TotalMessage += cr + "-----------------------------------------------------------------------------------------------" + cr;
   enumItems.moveNext();
      EventCounter += 1;
   } 
}
catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
return (TotalMessage);

}

function IIS_Inst()
{
        var TotalMessage = "", cr = "\r\n";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\Setup\\OC Manager\\Subcomponents\\iis_common";  
 	var regkey2 = "HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\Setup\\OC Manager\\Subcomponents\\iis_smtp";  
 	var regkey3 = "HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\Setup\\OC Manager\\Subcomponents\\snmp";  
	var regkey4 = "HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\Setup\\OC Manager\\Subcomponents\\aspnet";  
 	var regkey5 = "HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\Setup\\OC Manager\\Subcomponents\\iis_www";  
 	
        TotalMessage = "IIS \t"; 
        var CHECK = shell.RegRead( regkey1 );
	if (CHECK == 1) TotalMessage += "is installed" + cr;
        else TotalMessage += "is not installed" + cr;
        TotalMessage +="SMTP \t"
	var CHECK = shell.RegRead( regkey2 );
	if (CHECK == 1) TotalMessage += "is installed" + cr;
        else TotalMessage += "is not installed" + cr;
	TotalMessage += "SNMP \t"
	var CHECK = shell.RegRead( regkey3 );
	if (CHECK == 1) TotalMessage += "is installed" + cr;
        else TotalMessage += "is not installed" + cr;
	TotalMessage += "WWW \t"
	var CHECK = shell.RegRead( regkey5 );
	if (CHECK == 1) TotalMessage += "is installed" + cr;
        else TotalMessage += "is not installed" + cr;
	
try
{
        TotalMessage += "ASP.NET "
	var CHECK = shell.RegRead( regkey4 );
	if (CHECK == 1) TotalMessage += "is installed" + cr;
        else TotalMessage += "is not installed" + cr;
	return( TotalMessage);
    }
    catch( e )
    {
        return( TotalMessage += "is not installed" + cr );
    }
}	

function GetMulticast()
{
    try
    {
        var cr = "\r\n";
	var TAB = "\t";
        var LineBreak = "+------------------------------------------------------------------------------+";
        TotalMessage = LineBreak + cr;
        TotalMessage += "¦RTD Multicast Configuration:                                                  ¦" + cr;
        TotalMessage += LineBreak + cr + cr;
	
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey1 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MultiCastAddr";
  	var regkey2 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MultiCastTTL";
	var regkey3 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\RSMCompression";
	var regkey4 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\rtdCompression";
	
        var regkeyI1 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\IAgentPort";
  	var regkeyI2 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\IApplicationPort";
	var regkeyI3 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\ISkillsetPort";
	var regkeyI4 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\INodalPort";
        var regkeyI5 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\IIVRPort";
  	var regkeyI6 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\IRoutePort";
	
	var regkeyI1r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\IAgentRate";
  	var regkeyI2r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\IApplicationRate";
	var regkeyI3r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\ISkillsetRate";
	var regkeyI4r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\INodalRate";
        var regkeyI5r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\IIVRRate";
  	var regkeyI6r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\IRouteRate";
	
	var regkeyM1 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWAgentPort";
  	var regkeyM2 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWApplicationPort";
	var regkeyM3 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWSkillsetPort";
	var regkeyM4 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWNodalPort";
        var regkeyM5 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWIVRPort";
  	var regkeyM6 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWRoutePort";
	
	var regkeyM1r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWAgentRate";
  	var regkeyM2r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWApplicationRate";
	var regkeyM3r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWSkillsetRate";
	var regkeyM4r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWNodalRate";
        var regkeyM5r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWIVRRate";
  	var regkeyM6r = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MWRouteRate";
     
	TotalMessage += "Multicast IP group: " +  (shell.RegRead(regkey1)) + TAB + TAB + "Multicast time to live(TTL): " + (shell.RegRead(regkey2)) + " sec" + cr + cr;
	
	TotalMessage += "Interval To Date" + TAB + "IP Port:" + TAB + "Multicast Rate:" + cr;

	TotalMessage += "Agent:" + TAB + TAB + TAB + (shell.RegRead(regkeyI1)) + TAB + TAB + (shell.RegRead(regkeyI1r)) + " ms" + cr;
	TotalMessage += "Application:" + TAB + TAB + (shell.RegRead(regkeyI2)) + TAB + TAB + (shell.RegRead(regkeyI2r)) + " ms" + cr;
	TotalMessage += "Skillset:" + TAB + TAB +    (shell.RegRead(regkeyI3)) + TAB + TAB + (shell.RegRead(regkeyI3r)) + " ms" + cr;
	TotalMessage += "Nodal:" + TAB + TAB + TAB + (shell.RegRead(regkeyI4)) + TAB + TAB + (shell.RegRead(regkeyI4r)) + " ms" + cr;
	TotalMessage += "IVR:" + TAB + TAB + TAB +   (shell.RegRead(regkeyI5)) + TAB + TAB + (shell.RegRead(regkeyI5r)) + " ms" + cr;
	TotalMessage += "Route:" + TAB + TAB + TAB + (shell.RegRead(regkeyI6)) + TAB + TAB + (shell.RegRead(regkeyI6r)) + " ms" + cr +cr; 

	TotalMessage += "Moving Window" + TAB + TAB + "IP Port:" + TAB + "Multicast Rate:" + cr;
	TotalMessage += "Agent:" + TAB + TAB + TAB + (shell.RegRead(regkeyM1)) + TAB + TAB + (shell.RegRead(regkeyM1r)) + " ms" + cr;
	TotalMessage += "Application:" + TAB + TAB + (shell.RegRead(regkeyM2)) + TAB + TAB + (shell.RegRead(regkeyM2r)) + " ms" + cr;
	TotalMessage += "Skillset:" + TAB + TAB +    (shell.RegRead(regkeyM3)) + TAB + TAB + (shell.RegRead(regkeyM3r)) + " ms" + cr;
	TotalMessage += "Nodal:" + TAB + TAB + TAB + (shell.RegRead(regkeyM4)) + TAB + TAB + (shell.RegRead(regkeyM4r)) + " ms" + cr;
	TotalMessage += "IVR:" + TAB + TAB + TAB +   (shell.RegRead(regkeyM5)) + TAB + TAB + (shell.RegRead(regkeyM5r)) + " ms" + cr;
	TotalMessage += "Route:" + TAB + TAB + TAB + (shell.RegRead(regkeyM6)) + TAB + TAB + (shell.RegRead(regkeyM6r)) + " ms" + cr + cr; 

        TotalMessage += "RTD Compression: " +  (shell.RegRead(regkey4)) + TAB + TAB + "RSM Compression: " + (shell.RegRead(regkey3)) + cr + cr;

	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}


function GetCCMMPEPS()
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

	TotalMessage = "All Installed CCMM PEPS: \r\n";
	cr = "\r\n";
	
	RegistryPath = "SOFTWARE\\Nortel\\Contact Center Multimedia\\Records"

	try
	{
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
	   TotalMessage += KeyName + cr;
   	} 
	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
return (TotalMessage);
}

function GetCCTPEPS()
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

	TotalMessage = "All Installed CCT PEPS: \r\n";
	cr = "\r\n";
	
	RegistryPath = "SOFTWARE\\Nortel\\Communication Control Toolkit\\Records"

	try
	{
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
	   TotalMessage += KeyName + cr;
   	} 
	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
return (TotalMessage);
}

function GetCCMA6Info()
{
    try
    {
	var TEMP;
        var cr = "\r\n";
        var shell = new ActiveXObject( "WScript.Shell" );
        var LineBreak = "+------------------------------------------------------------------------------+";
        TotalMessage = LineBreak + cr;
        TotalMessage += "¦CCMA Real Time Reporting Settings:                                            ¦" + cr;
        TotalMessage += LineBreak + cr + cr;
	var regkey1  = "HKLM\\Software\\Nortel\\RTD\\IP Receive";
       	var regkey2  = "HKLM\\Software\\Nortel\\RTD\\IP Send";
	var regkey2a = "HKLM\\Software\\Nortel\\EmergencyHelp\\IP Send";
  	var regkey3  = "HKLM\\Software\\Nortel\\RTD\\Output Rate";
	var regkey4  = "HKLM\\Software\\Nortel\\RTD\\Transform Rate";
  	var regkey5  = "HKLM\\Software\\Nortel\\RTD\\OAM Timeout";
	var regkey6  = "HKLM\\Software\\Nortel\\RTD\\Transmission Option";
       	var regkey7  = "HKLM\\Software\\Nortel\\RTD\\Max Unicast Sessions";
  	var regkey8  = "HKLM\\Software\\Nortel\\RTD\\Compression";
	var regkey9  = "HKLM\\Software\\Nortel\\RTD\\Multicast Filter";
  	var regkey0  = "HKLM\\Software\\Nortel\\RTD\\NCC Server";
	var regkey0a = "HKLM\\Software\\Nortel\\Ngen comm\\localsitename";
	var regkey0b = "HKLM\\Software\\Nortel\\WClient\\ServerSoapName";
	var regkey0c = "HKLM\\Software\\Nortel\\WClient\\Version";
	TotalMessage += "IP Receive Address: \t";
	TotalMessage += (shell.RegRead( regkey1 ) + cr );
	TotalMessage += "IP Send Address: \t";
	TotalMessage += (shell.RegRead( regkey2 ) + cr );
	TotalMessage += "Emergency Help: \t";
	TotalMessage += (shell.RegRead( regkey2a ) + cr + cr);
	TotalMessage += "Output rate: \t\t";
	TotalMessage += (shell.RegRead( regkey3 ) + " milliseconds" + cr );
	TotalMessage += "Transform Rate: \t";
	TotalMessage += (shell.RegRead( regkey4 ) + " milliseconds" + cr );
	TotalMessage += "OAM Timeout: \t\t";
	TotalMessage += (shell.RegRead( regkey5 ) + " milliseconds" + cr + cr );
	TotalMessage += "Transmission Options: \t";
	TEMP = shell.RegRead( regkey6 );
	   if (TEMP == 0) { TotalMessage += "Multicast"; } 
	   if (TEMP == 1) { TotalMessage += "Unicast";   }	
	   if (TEMP == 2) { TotalMessage += "Multicast and Unicast"; }	
        TotalMessage += cr;
	TotalMessage += "Max Unicast Sessions: \t";
	TotalMessage += (shell.RegRead( regkey7 ) + cr );
	TotalMessage += "Compress RTD Packets: \t";
	TEMP = shell.RegRead( regkey8 );
	   if (TEMP == 0) { TotalMessage += "OFF"; }
	   if (TEMP == 1) { TotalMessage += "ON";  }
	TotalMessage += cr;
	TotalMessage += "Multicast Filter: \t";
	TEMP = shell.RegRead( regkey9 );
	   if (TEMP == 0) { TotalMessage += "OFF"; }
	   if (TEMP == 1) { TotalMessage += "ON";  }
	TotalMessage += cr + cr;
	TotalMessage += "Local Site Name: \t";
	TotalMessage += (shell.RegRead( regkey0a ) + cr );
	TotalMessage += "Server SOAP Name: \t";
	TotalMessage += (shell.RegRead( regkey0b ) + cr );
	TotalMessage += "NCC Server: \t\t";
	TotalMessage += (shell.RegRead( regkey0 ) + cr );
	TotalMessage += "CCMA Version: \t\t";
	TotalMessage += (shell.RegRead( regkey0c ) + cr + cr );
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}


function GetCPCLAN()
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

	TotalMessage = "";
	cr = "\r\n";
	
	RegistryPath = "Software\\Nortel\\CallPilot\\ConfigWizard\\NIC\\"
	RegistryKey = "CLAN IP Address"
	
	try
	{
	Locator = new ActiveXObject("WbemScripting.SWbemLocator");
	oSvc = Locator.ConnectServer("127.0.0.1", "root\\default", "", "");
	oReg = oSvc.Get("StdRegProv");
	oMethod = oReg.Methods_.Item("GetMultiStringValue");
	oInParam = oMethod.InParameters.SpawnInstance_();
	oInParam.hDefKey = HKLM;
	oInParam.sSubKeyName = RegistryPath;
	oInParam.sValueName = RegistryKey;
	oOutParam = oReg.ExecMethod_(oMethod.Name, oInParam);
	aNames = oOutParam.sValue.toArray();

	for (x = 0; x <= aNames.length -1; x++)
	{
	   var MultiStringValue = aNames[x]
	   TotalMessage += "Call Pilot CLAN: " + MultiStringValue;
   	} 
       
	return (TotalMessage + cr);
	}
  catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}
 


function GetCPELAN()
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

	TotalMessage = "";
	cr = "\r\n";
	
	RegistryPath = "Software\\Nortel\\CallPilot\\ConfigWizard\\NIC\\"
	RegistryKey = "ELAN IP Address"
	
	try
	{
	Locator = new ActiveXObject("WbemScripting.SWbemLocator");
	oSvc = Locator.ConnectServer("127.0.0.1", "root\\default", "", "");
	oReg = oSvc.Get("StdRegProv");
	oMethod = oReg.Methods_.Item("GetMultiStringValue");
	oInParam = oMethod.InParameters.SpawnInstance_();
	oInParam.hDefKey = HKLM;
	oInParam.sSubKeyName = RegistryPath;
	oInParam.sValueName = RegistryKey;
	oOutParam = oReg.ExecMethod_(oMethod.Name, oInParam);
	aNames = oOutParam.sValue.toArray();

	for (x = 0; x <= aNames.length -1; x++)
	{
	   var MultiStringValue = aNames[x]
	   TotalMessage += "Call Pilot ELAN: " + MultiStringValue;
   	} 
       
	return (TotalMessage + cr);
	}
  catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}


function GetCPVer()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Call Pilot Version: ";
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey1 = "HKLM\\Software\\Nortel\\Setup\\CPServer\\Version";
	TotalMessage += (shell.RegRead( regkey1 ));
	return( TotalMessage + cr );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

