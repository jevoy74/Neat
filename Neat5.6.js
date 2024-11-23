/************************************************************************************************
 *    Name:       NEAT5.6.js 
 *    Version:    5.6
 *    Author:     Jason Evoy - jevoy@avaya.com
 *    Date:       April 4th, 2012
 *    Notes:	  This is a script that polls the registry & WMI for PC data & outputs to a file
 *                Designed for AACC 6.0 and above use Neat 4.3.js for older releases of CCMS, SCCS
 *    Min Req:    Windows XP, 7, 2008 Internet Explorer 6.0 or higher
 ************************************************************************************************/
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
   var Title = "Avaya Enterprise Audit Tool Ver 5.6                       - Jason Evoy";
   var MessageLine1 = "Creating log file...........................\n\nPlease wait until finished message appears before opening file\nThe audit usually takes less than 1 minute to complete";  
   var WshShell = WScript.CreateObject("WScript.Shell");
   var ButtonCode = WshShell.Popup(MessageLine1, Timeout, Title, OkButton + InfoIcon);
}

function PopupDone()
  {
    var OkButton = 0; var Timeout = 15; var InfoPic = 64;
    var Title = "Avaya Enterprise Audit Tool Ver 5.6                     - Jason Evoy";
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
   var LineBreak = "+--------------------------------------------------------------------+";
   os.write(LineBreak + cr);
   os.write("¦Avaya Enterprise Audit Tool v5.6 Results - ");
   os.write(Date());
   os.write(cr);
   os.write(LineBreak + cr +cr);   
   os.write(GetAACCInfo());
   os.write(GetNBNMINFO61());
   os.write(GetCORES61());
   os.write(GetSIPInfo());

//   os.write(GetM1Info());      // Needs work  Re: AACC with M1 connection
//   os.write(GetM1Serial());    // needs work
//   os.write(GetAccessLink());  // Needs work
 
   os.write(GetCCMA61());
   os.write(GetLMInfo61());

   os.write(GetAACC6PEPS());
   os.write(GetCCCC6PEPS());
   os.write(GetCCMA6PEPS());
   os.write(GetCCLM6PEPS());
   os.write(GetCCMM6PEPS());
   os.write(GetCCMSU6PEPS());
   os.write(GetCCT6PEPS());
   os.write(GetCCWS6PEPS());
    
   os.write(GetMulticast61());
   os.write(LineBreak + cr);
   os.write("¦Operating System Info:                                              ¦" + cr);
   os.write(LineBreak + cr + cr);
   os.write(GetOS_NEW());
   os.write(LineBreak + cr);
   os.write("¦Computer Info:                                              ¦" + cr);
   os.write(LineBreak + cr + cr);
   os.write(GetCompInfo());
   os.write(cr);
   os.write(GetJavaVer());
   os.write(PrintFile("C:\\Java.log"));
   os.write(GetJavaUpdate());
   os.write(cr);
   os.write(ShowNetwork());
   os.write(cr);
   os.write(GetMasterBrowser());
   os.write(GetMEDIASENSING());
   os.write(GetIPFORWARDING());
   os.write(GetAPIPA());
   os.write(GetIPv6());
   os.write(GetNIC());
   os.write(chkcpu());
   os.write(ShowDriveList());
   os.write(cr);
   os.write(GetPATHLENGTH());  
   os.write(GetPATH());
   os.write(cr);
   os.write(LineBreak + cr);
   os.write("¦Running Processes                                                   ¦" + cr);
   os.write(LineBreak + cr + cr);
   os.write(GetPROCESS() +cr );
   os.write(LineBreak + cr);
   os.write("¦Services                                                            ¦" + cr);
   os.write(LineBreak + cr + cr);
   os.write(chkservices());
   os.write(LineBreak + cr);
   os.write("¦Installed Software (Through Windows Installer)                      ¦" + cr);
   os.write(LineBreak + cr + cr);
   os.write(GetInstalledPrograms() + cr);
   os.write(GetInstalledPrograms32() + cr);
   os.write(GetHotFixList() + cr);	
   os.write(LineBreak + cr);
   os.write("¦Registry Run Keys                                                   ¦" + cr);
   os.write(LineBreak + cr + cr);
   TotalMessage = "";
   os.write(GetRunKey());
   os.write(cr);
   os.write(LineBreak + cr);
   os.write("¦HOSTS File (Comments Excluded)                                      ¦" + cr);
   os.write(LineBreak + cr);
   os.write(PrintHOSTS() + cr);
   os.write(LineBreak + cr);
   os.write("¦LMHOSTS File (Comments Excluded)                                    ¦" + cr);
   os.write(LineBreak + cr);
   os.write(PrintLMHOSTS());
   os.write(LineBreak + cr);
   os.write("¦TCP/IP Connections                                                  ¦" + cr);
   os.write(LineBreak + cr);
   os.write(PrintFile("C:\\netstat.log"));
   os.write(PrintFile("C:\\netstat.rte"));
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
   var JavaFile = "C:\\Java.log";
   if (shell.FileExists(PINGFILE)) { shell.DeleteFile(PINGFILE);}
   if (shell.FileExists(NetStatFile)) {shell.DeleteFile(NetStatFile); }
   if (shell.FileExists(NetStatFile2)) {shell.DeleteFile(NetStatFile2); }
   if (shell.FileExists(JavaFile)) {shell.DeleteFile(JavaFile); }
   }
   catch ( e )
   {
   }
}

function GetNBNMINFO61()
{
    try
    { 
        var cr = "\r\n";
	var TotalMessage = "";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\NINAMESERVER\\CLan";
	var regkey2 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\NINAMESERVER\\ELan";
        var regkey3 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\NINAMESERVER\\NSAddr";
	var regkey4 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\NINAMESERVER\\SiteName";
	var regkey6 = "HKLM\\System\\CurrentControlSet\\Services\\NBNM_Service\\ImagePath";
        TotalMessage += "Site Name: ";
	TotalMessage += (shell.RegRead( regkey4 ) + cr);
	TotalMessage += "NBNM IP: ";
	TotalMessage += (shell.RegRead( regkey3 ) + cr + "Path:" + shell.RegRead( regkey6 ) + cr + cr);	
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

function GetAACCInfo()
{
    try
    {
        var cr = "\r\n";
        var LineBreak = "+--------------------------------------------------------------------+";
        TotalMessage = LineBreak + cr;
        TotalMessage += "¦Contact Center Info:                                                ¦" + cr;
        TotalMessage += LineBreak + cr + cr;
	
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey2 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\SerialNumber";
  	var regkey3 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\Platform Version";
	var regkey4 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\Revision";
  	var regkey5 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\UserCompany";
	var regkey6 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\UserName";
	var regkey8 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\CCMSHostname";
	var regkey9 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\CCTHostname";
	TotalMessage += "Version: ";
	TotalMessage += (shell.RegRead( regkey3 ) );
	TotalMessage += " Build: ";
	TotalMessage += (shell.RegRead( regkey4 ) + cr );
	TotalMessage += "AACC Hostname: ";
	TotalMessage += (shell.RegRead( regkey8 ) + cr );
	TotalMessage += "CCT Hostname:  ";
	TotalMessage += (shell.RegRead( regkey9 ) + cr );
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
   var TotalMessage = "", cr = "\r\n", TimeGenerated = "", ProdType = "", DEP = "";
      
   try
   {
  
   var objWMIService = GetObject("winmgmts:\\\\.\\root\\CIMV2");
   var colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem", "WQL", 0x10 | 0x20);
   var enumItems = new Enumerator(colItems);
   for (; !enumItems.atEnd(); enumItems.moveNext())
      {
      var objItem = enumItems.item();
      

	TotalMessage +="Operating System: \t" + objItem.Caption + " " + cr;
	TotalMessage +="Service Pack: \t\t"   + objItem.CSDVersion    + cr;
        ProdType = objItem.ProductType;
    	TotalMessage +="Product Type: \t\t"   + (ConvertProdType(ProdType)) + cr;
	TotalMessage +="Product ID: \t\t"  + objItem.SerialNumber  + cr;
        TotalMessage +="Architecture: \t\t" + objItem.OSArchitecture + cr;
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
        DEP = objItem.DataExecutionPrevention_SupportPolicy
	TotalMessage +="Data Execution Prev: \t";
	    if (DEP == 1) { TotalMessage+= "DEP is enabled for all processes (AlwaysOn)" }   
  	    if (DEP == 2) { TotalMessage+= "DEP is enabled for Windows services only (OptIn - Windows Default)" } 
 	    if (DEP == 3) { TotalMessage+= "DEP is enabled for all processes (OptOut)" }       
   	    if (DEP == 0) { TotalMessage+= "DEP is disabled for all processes (AlwaysOff)" } 
	TotalMessage += cr;
      }

   }
   catch ( e )
   {
        var ErrorMessage = "";
        ErrorMessage += "Send screen cap of this error to jevoy@avaya.com" + cr;
        ErrorMessage += "Function: GetOS_New()" + cr;
        ErrorMessage += e + cr;
        ErrorMessage += (e.number &0xFFFF) + cr;
	ErrorMessage += e.description + cr;

return (ErrorMessage);
   }

   
return (TotalMessage);
}

function ConvertProdType(ProdType)
{
    var TotalMessage = "";

  
    if (ProdType == 1) { TotalMessage+= "Workstation" }   
    if (ProdType == 2) { TotalMessage+= "Domain Controller" } 
    if (ProdType == 3) { TotalMessage+= "Server" }       
    if (ProdType >3)   { TotalMessage+= "Unknown" } 

    return(TotalMessage);
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

function GetCompInfo()
{
   var TotalMessage = "", cr = "\r\n";
      
   try
   {
  
   var objWMIService = GetObject("winmgmts:\\\\.\\root\\CIMV2");
   var colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", 0x10 | 0x20);
   var enumItems = new Enumerator(colItems);
   for (; !enumItems.atEnd(); enumItems.moveNext())
      {
      var objItem = enumItems.item();

	TotalMessage +="Computer Name: \t\t" + objItem.Name  + cr;
	IP = objItem.DNSHostName;
        TotalMessage += "DNSHostName:\t\t" + IP + cr; 
 	TotalMessage += (PingIPs(IP)) + cr;
        TotalMessage +="Domain: \t\t"        + objItem.Domain + cr;
	TotalMessage +="Workgroup: \t\t"     + objItem.Workgroup + cr;
        TotalMessage +="Manufacturer: \t\t"  + objItem.Manufacturer + cr;
	TotalMessage +="Model: \t\t\t"         + objItem.Model + cr;
	TotalMessage +="Daylight Savings Time in Effect (Null is off): " + objItem.DaylightInEffect + cr;
	TotalMessage +="Processors: \t\t"    + objItem.NumberOfProcessors + cr;
        TotalMessage +="Logical Processors: \t" + objItem.NumberOfLogicalProcessors + cr;
        
      }

   }
   catch ( e )
   {
      return (e + e.description);
   }

   
return (TotalMessage);
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
try
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
	oSvc = Locator.ConnectServer();
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
catch (e)
   {
	return (e.description);
   } 
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
 return (e.description);
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

function chkcpu()
{
   var Locator = new ActiveXObject ("WbemScripting.SWbemLocator");
   var CPU = Locator.connectserver();
   var cpuerr = "0", TotalMessage = "", cr = "\r\n";
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

function GetCORES61()
{
   try
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
		
	RegistryPath1 = "SOFTWARE\\Wow6432Node\\Nortel\\"

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
           if (KeyName1 == "Co-Residency") { GetRunSoftwareKey (KeyName1); }
   	}
   }
  catch( e )
    {
        return( "" );
    } 
return (TotalMessage);
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

function NetStat()
{
   var Wshell = new ActiveXObject("WScript.Shell");
   var command = "cmd.exe /c netstat.exe -nao>C:\\netstat.log"
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

function GetCCMA61()
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
	var regkey1  = "HKLM\\Software\\Wow6432Node\\Nortel\\RTD\\IP Receive";
       	var regkey2  = "HKLM\\Software\\Wow6432Node\\Nortel\\RTD\\IP Send";
  	var regkey3  = "HKLM\\Software\\Wow6432Node\\Nortel\\RTD\\Output Rate";
	var regkey4  = "HKLM\\Software\\Wow6432Node\\Nortel\\RTD\\Transform Rate";
  	var regkey5  = "HKLM\\Software\\Wow6432Node\\Nortel\\RTD\\OAM Timeout";
	var regkey6  = "HKLM\\Software\\Wow6432Node\\Nortel\\RTD\\Transmission Option";
       	var regkey7  = "HKLM\\Software\\Wow6432Node\\Nortel\\RTD\\Max Unicast Sessions";
  	var regkey8  = "HKLM\\Software\\Wow6432Node\\Nortel\\RTD\\Compression";
  	
	var regkey0a = "HKLM\\Software\\Wow6432Node\\Nortel\\WClient\\Version";
	var regkey0b = "HKLM\\Software\\Wow6432Node\\Nortel\\WClient\\LDAPPORT";
	var regkey0c = "HKLM\\Software\\Wow6432Node\\Nortel\\WClient\\SSLPORT";
	
	TotalMessage += "IP Receive Address: \t";
	TotalMessage += (shell.RegRead( regkey1 ) + cr );
	TotalMessage += "IP Send Address: \t";
	TotalMessage += (shell.RegRead( regkey2 ) + cr );
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
	
	TotalMessage += "CCMA Version: \t\t";
	TotalMessage += (shell.RegRead( regkey0a ) + cr);
	TotalMessage += "LDA Port: \t\t";
	TotalMessage += (shell.RegRead( regkey0b ) + cr);
	TotalMessage += "SSLPort: \t\t";
	TotalMessage += (shell.RegRead( regkey0c ) + cr + cr );

	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}

function GetLMInfo61()
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
       
	var regkey1a = "HKLM\\Software\\Wow6432Node\\Nortel\\LM\\MajorLicenseUsage";
	var regkey1b = "HKLM\\Software\\Wow6432Node\\Nortel\\LM\\CriticalLicenseUsage";
	var regkey2 =  "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\LM_IP_PORT_Primary";
  	var regkey3 =  "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\LM_IP_PORT_Secondary";
	var regkey4 =  "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\LM_IP_Primary";
  	var regkey5 =  "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\LM_IP_Secondary";
	var regkey6 =  "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\LM_Package";
	var regkey8 =  "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\LM_OpenQueue";
	var regkey7 =  "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\LM_HeteroNetwork";
	var regkey9 =  "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\LM_HeteroNetworkAdmin";
	var regkey10 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\SerialNumber";
	TotalMessage += "Package: ";
	TotalMessage += (shell.RegRead( regkey6 ) + cr );
	TotalMessage += "Serial Number: ";
	TotalMessage += (shell.RegRead( regkey10 ) + cr );
        TotalMessage += "Open Queue: ";
	TotalMessage += (shell.RegRead( regkey8 ) + cr );
        TotalMessage += "Hetero Network: ";
	TotalMessage += (shell.RegRead( regkey7 ) + cr );
        TotalMessage += "Hetero Network Admin: ";
	TotalMessage += (shell.RegRead( regkey9 ) + cr );
	TotalMessage += "Major License Usage: ";
	TotalMessage += (shell.RegRead( regkey1a ) + "%" + cr );
        TotalMessage += "Critical License Usage: ";
	TotalMessage += (shell.RegRead( regkey1b ) + "%" + cr );
	TotalMessage += "Primary IP: ";
	IP = (shell.RegRead( regkey4 ) );
	TotalMessage += IP + "\:";
	TotalMessage += (shell.RegRead( regkey2 ) + cr );
	TotalMessage += (PingIPs(IP)) + cr + cr;
	TotalMessage += "Secondary IP: ";
	IP = (shell.RegRead( regkey5 ) );
           if (IP != regkey5)  { IP = "Unconfigured" };
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
 
function GetHotFixList()
{
   try
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
	
	RegistryPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Component Based Servicing\\ApplicabilityEvaluationCache"
        TotalMessage =""; cr = "\r\n"; FINDKB = "";
	Locator = new ActiveXObject("WbemScripting.SWbemLocator");
	oSvc = Locator.ConnectServer(null, "root\\default", "", "");
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
           if (KeyName.substr(12,2) == "KB") { TotalMessage += (KeyName.substr(12,9) + cr);}
   	} 
   }
   catch( e )
    {
        TotalMessage = "";
	return( TotalMessage);
    }
return (TotalMessage);
}

function GetInstalledPrograms32()
{
   try
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
	
	RegistryPath = "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
        TotalMessage ="";
	Locator = new ActiveXObject("WbemScripting.SWbemLocator");
	oSvc = Locator.ConnectServer(null, "root\\default", "", "");
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
   }
     catch( e )
    {
        return( "" );
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
      return (e + e.Description);
   } 
}

function GetSIPInfo()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "\r\nLocal SIP Provider: \r\n";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey1 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\Switch\\SIPServerType";
	var regkey2 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\Switch\\MSLocale";
 	var regkey3 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\Switch\\DomainName";
	var regkey4 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\Switch\\TcpUdpPort";
	var regkey5 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\Switch\\TlsPort";
	var regkey6 = "HKLM\\Software\\Wow6432Node\\Nortel\\Setup\\Switch\\SIPIP";
	
	TotalMessage += ("SIP Server Type: " + shell.RegRead( regkey1 ) + cr);
	TotalMessage += ("MS Locale: " + shell.RegRead( regkey2 ) + cr);
	TotalMessage += ("Domain Name: " + shell.RegRead( regkey3 ) + cr);
	TotalMessage += ("CLAN Local Listening TCP/UDP Port: " + shell.RegRead( regkey4 ) + cr);
	TotalMessage += ("CLAN Local Listening TLS Port:  " + shell.RegRead( regkey5 ) + cr);
	
	TotalMessage += ("SIP IP: ");
	IP = (shell.RegRead( regkey6 ));
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

function GetMulticast61()
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
       	var regkey1 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MultiCastAddr";
  	var regkey2 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MultiCastTTL";
	var regkey3 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\RSMCompression";
	var regkey4 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\rtdCompression";
	
        var regkeyI1 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\IAgentPort";
  	var regkeyI2 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\IApplicationPort";
	var regkeyI3 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\ISkillsetPort";
	var regkeyI4 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\INodalPort";
        var regkeyI5 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\IIVRPort";
  	var regkeyI6 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\IRoutePort";
	
	var regkeyI1r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\IAgentRate";
  	var regkeyI2r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\IApplicationRate";
	var regkeyI3r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\ISkillsetRate";
	var regkeyI4r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\INodalRate";
        var regkeyI5r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\IIVRRate";
  	var regkeyI6r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\IRouteRate";
	
	var regkeyM1 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWAgentPort";
  	var regkeyM2 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWApplicationPort";
	var regkeyM3 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWSkillsetPort";
	var regkeyM4 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWNodalPort";
        var regkeyM5 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWIVRPort";
  	var regkeyM6 = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWRoutePort";
	
	var regkeyM1r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWAgentRate";
  	var regkeyM2r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWApplicationRate";
	var regkeyM3r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWSkillsetRate";
	var regkeyM4r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWNodalRate";
        var regkeyM5r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWIVRRate";
  	var regkeyM6r = "HKLM\\Software\\Wow6432Node\\Nortel\\ICCM\\SDP\\MWRouteRate";
     
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


function GetAACC6PEPS()
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
        cr = "\r\n", TAB = "\t\t";
	TotalMessage = LineBreak + cr;
        TotalMessage +="¦Installed Patches as of: " + Date() + "                             ¦" + cr;
        TotalMessage += LineBreak + cr + cr;
	RegistryPath = "SOFTWARE\\Wow6432Node\\Nortel\\Contact Center\\Product Updates\\CCMS"

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
	   TotalMessage += KeyName + TAB;
	   GetPatchKey (KeyName); 
           TotalMessage += cr;
   	} 
	}
       	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
TotalMessage += cr;
return (TotalMessage);
}

function GetCCCC6PEPS()
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
	TotalMessage = "",cr = "\r\n", TAB = "\t\t";
	RegistryPath = "SOFTWARE\\Wow6432Node\\Nortel\\Contact Center\\Product Updates\\CommonComponents"

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
	   TotalMessage += KeyName + TAB;
	   GetPatchKey (KeyName); 
           TotalMessage += cr;
   	} 
      	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
TotalMessage += cr;
return (TotalMessage);
}

function GetCCMA6PEPS()
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
	TotalMessage = "",cr = "\r\n", TAB = "\t\t";
	RegistryPath = "SOFTWARE\\Wow6432Node\\Nortel\\Contact Center\\Product Updates\\CCMA"

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
	   TotalMessage += KeyName + TAB;
	   GetPatchKey (KeyName); 
           TotalMessage += cr;
   	} 
      	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
TotalMessage += cr;
return (TotalMessage);
}

function GetCCLM6PEPS()
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
	TotalMessage = "",cr = "\r\n", TAB = "\t\t";
	RegistryPath = "SOFTWARE\\Wow6432Node\\Nortel\\Contact Center\\Product Updates\\LM"

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
	   TotalMessage += KeyName + TAB;
	   GetPatchKey (KeyName); 
	   TotalMessage += cr;
   	} 
      	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
TotalMessage += cr;
return (TotalMessage);
}

function GetCCMM6PEPS()
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
	TotalMessage = "", cr = "\r\n", TAB = "\t\t";
	RegistryPath = "SOFTWARE\\Wow6432Node\\Nortel\\Contact Center\\Product Updates\\CCMM"

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
	   TotalMessage += KeyName + TAB;
	   GetPatchKey (KeyName); 
	   TotalMessage += cr;
   	} 
      	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
TotalMessage += cr;
return (TotalMessage);
}

function GetCCMSU6PEPS()
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
	TotalMessage = "",cr = "\r\n", TAB = "\t\t";
	RegistryPath = "SOFTWARE\\Wow6432Node\\Nortel\\Contact Center\\Product Updates\\CCMSU"

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
	   TotalMessage += KeyName + TAB;
	   GetPatchKey (KeyName);
           TotalMessage += cr;
   	} 
      	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
TotalMessage += cr;
return (TotalMessage);
}

function GetCCT6PEPS()
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
	TotalMessage = "",cr = "\r\n", TAB = "\t\t";
	RegistryPath = "SOFTWARE\\Wow6432Node\\Nortel\\Contact Center\\Product Updates\\CCT"

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
	   TotalMessage += KeyName + " " + TAB;
	   GetPatchKey (KeyName); 
	   TotalMessage += cr;
   	} 
      	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
TotalMessage += cr;
return (TotalMessage);
}

function GetCCWS6PEPS()
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
        TotalMessage = "",cr = "\r\n", TAB = "\t\t";
	RegistryPath = "SOFTWARE\\Wow6432Node\\Nortel\\Contact Center\\Product Updates\\CCWS"

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
	   TotalMessage += KeyName + TAB;
	   GetPatchKey (KeyName); 
           TotalMessage += cr;
   	} 
      	}
	catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
TotalMessage += cr;
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
            TotalMessage += "   ";
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

function GetIPv6()
{

    try
    {
        var TotalMessage = "IPv6 is ", cr = "\r\n";
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\TCPIP6\\Parameters\\DisabledComponents";       
        var IPv6 = shell.RegRead( regkey );
	if (IPv6 == 255) // 0x000000ff
           { TotalMessage += "Disabled in the registry" + cr};
	else if (IPv6 == -1) // 0xffffffff
 	   { TotalMessage += "Disabled in the registry"+ cr};
        else {TotalMessage += "Not disabled in the registry - Check individual NICs" + cr};
	return( TotalMessage + cr );
    }
    catch( e )
    {
        return( TotalMessage + "Not disabled in the registry - Check individual NICs" + cr);  // if DisabledComponents Key is missing
    }
}

function GetJavaVer()
{
   var Wshell = new ActiveXObject("WScript.Shell");
   var command = "cmd.exe /c java -version 2>\\java.log"
   Wshell.run (command, 0, true);
   return;
}  

function GetJavaUpdate()
{
   try
   {
   var TotalMessage = "Java Automatic Updates is ";
   var cr = "\r\n";
   var JREUpdate = "";
   var shell = new ActiveXObject( "WScript.Shell" );
   var regkey1 = "HKLM\\Software\\Wow6432Node\\JavaSoft\\Java Update\\Policy\\EnableJavaUpdate";
   JREUpdate = shell.RegRead( regkey1 );
	if (JREUpdate == 0)
	   {
		TotalMessage += "Disabled";
           }
	if (JREUpdate == 1)
	   {   		
		TotalMessage += "Enabled";
	   }	
        return( TotalMessage + cr );
   }
   catch( e )
    {
        TotalMessage = "";
	return( TotalMessage);
    }
   return(TotalMessage);
}	