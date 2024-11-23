/**********************************************************************************************************
 *    Name:       NEAT3.4.js 
 *    Version:    3.4
 *    Author:     Jason Evoy - jevoy@nortel.com
 *    Date:       Aug 24, 2006
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
   var Title = "Nortel Enterprise Audit Tool Ver 3.4                       - Jason Evoy";
   var MessageLine1 = "Creating log file...........................\n\nPlease wait until finished message appears before opening file\nThe audit usually takes less than 1 minute to complete";  
   var WshShell = WScript.CreateObject("WScript.Shell");
   var ButtonCode = WshShell.Popup(MessageLine1, Timeout, Title, OkButton + InfoIcon);
}

function PopupDone()
  {
    var OkButton = 0; var Timeout = 15; var InfoPic = 64;
    var Title = "Nortel Enterprise Audit Tool Ver 3.4                     - Jason Evoy";
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
   os.write("�Nortel Enterprise Audit Tool Results - ");
   os.write(Date());
   os.write(cr);
   os.write(LineBreak + cr +cr);   
   os.write(GetSWCINFO());
   os.write(GetOTMInfo());
   os.write(GetICCMINFO());
   os.write(GetCC6Info());
   os.write(GetLMInfo());
   os.write(GetREPINFO());
   os.write(GetREPIP());
   os.write(GetSU1());
   os.write(GetSU2());
   os.write(GetSUDATETIME());
   os.write(cr);
   os.write(GetPEPS());	
   os.write(GetCPCLAN());   
   os.write(GetCPELAN());
   os.write(cr);
   os.write(GetCPPEPS());
   os.write(cr);
   os.write(GetClientSU());
   os.write(GetSybaseVer());
   os.write(GetSybaseSECC());
   os.write(GetNBNMINFO());
   os.write(GetCORES());
   os.write(GetVoiceServices());
   os.write(GetAccessLink());
   os.write(GetM1Info());
   os.write(GetOS());
   os.write(GetChasis());
   os.write(cr);
   os.write(chkcpu());
   os.write(chkmemory());
   os.write(chkpagefile());
   os.write(ShowNetwork());
   os.write(GetMasterBrowser());
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
   os.write(LineBreak + cr);
   os.write("�Running Processes                                                             �" + cr);
   os.write(LineBreak + cr + cr);
   os.write(GetPROCESS() +cr );
   os.write(LineBreak + cr);
   os.write("�Services                                                                      �" + cr);
   os.write(LineBreak + cr + cr);
   os.write(chkservices());
   os.write(LineBreak + cr);
   os.write("�Installed Software (Through Windows Installer)                                �" + cr);
   os.write(LineBreak + cr + cr);
   os.write(GetInstalledPrograms() +cr);	
   os.write(LineBreak + cr);
   os.write("�Registry Run Keys                                                             �" + cr);
   os.write(LineBreak + cr + cr);
   TotalMessage = "";
   os.write(GetRunKey());
   os.write(cr);
   os.write(LineBreak + cr);
   os.write("�HOSTS File (Comments Excluded)                                                �" + cr);
   os.write(LineBreak + cr);
   os.write(PrintHOSTS() + cr);
   os.write(LineBreak + cr);
   os.write("�LMHOSTS File (Comments Excluded)                                              �" + cr);
   os.write(LineBreak + cr);
   os.write(PrintLMHOSTS());
   os.write(LineBreak + cr);
   os.write("�TCP/IP Connections                                                            �" + cr);
   os.write(LineBreak + cr);
   os.write(PrintFile("C:\\netstat.log"));
   os.write(PrintFile("C:\\netstat.rte"));
   os.write(cr + cr + cr);
   os.write(LineBreak + cr);
   os.write("�Application Event Log (Last 25 Events)                                        �" + cr);
   os.write(LineBreak + cr + cr + cr);
   os.write(GetAppEvents() + cr + cr + cr);  
   os.write(LineBreak + cr);
   os.write("�System Event Log (Last 25 Events)                                             �" + cr);
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
	var regkey5 = "HKLM\\Software\\Nortel\\ICCM\\SDP\\MultiCastAddr";
	var regkey4 = "HKLM\\Software\\Nortel\\Setup\\NINAMESERVER\\SiteName";
	var regkey6 = "HKLM\\System\\CurrentControlSet\\Services\\NBNM_Service\\ImagePath";
        TotalMessage += "Site Name: ";
	TotalMessage += (shell.RegRead( regkey4 ) + cr);
	TotalMessage += "Multicast IP: ";
	TotalMessage += (shell.RegRead( regkey5 ) + cr );
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

function GetSUDATETIME()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Service Update installed on ";
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
	TotalMessage += ("Customer: " + shell.RegRead( regkey4 ) + cr);
	TotalMessage += ("M1 HostName: " + shell.RegRead( regkey2 ) + cr);
	TotalMessage += ("M1 IP: ");
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
 	var regkey4 = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Hostname";
	var regkey5 = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Domain";
        TotalMessage += (shell.RegRead( regkey1 ) + cr );
	TotalMessage += ("Service Pack: " + GetSP());
	TotalMessage += ("Computer Name: " + GetComputerName()+ cr);
	TotalMessage += ("Domain: " + shell.RegRead( regkey5 ) + cr);
	TotalMessage += "Hostname: ";
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

function GetSP()
{
    try
    {
        var cr = "\r\n"; var TotalMessage = ""; 
	var shell = new ActiveXObject( "WScript.Shell" );
        var regkey2 = "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\CSDVersion";
 	TotalMessage += (shell.RegRead( regkey2 ) + cr);
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "None" + cr;
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
      return ("Could not get memory info because: " + error.description);
      }

   for (;!imemory.atEnd();imemory.moveNext())
   {
      var memoryInstance = imemory.item();
      TotalMessage 
      += ( "Total Physical Memory: " + memoryInstance.TotalVisibleMemorySize + cr	      
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
      return ("Could not get CPU info because: " + error.description);
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
   if (StoreFamily == 11 ) CPUFamily = " Pentium� brand ";
   if (StoreFamily == 12 ) CPUFamily = " Pentium� Pro ";
   if (StoreFamily == 13 ) CPUFamily = " Pentium� II ";
   if (StoreFamily == 14 ) CPUFamily = " Pentium� processor with MMX technology ";
   if (StoreFamily == 15 ) CPUFamily = " Celeron� ";
   if (StoreFamily == 16 ) CPUFamily = " Pentium� II Xeon ";
   if (StoreFamily == 17 ) CPUFamily = " Pentium� III ";
   if (StoreFamily == 18 ) CPUFamily = " M1 Family ";
   if (StoreFamily == 19 ) CPUFamily = " M2 Family ";
   if (StoreFamily == 24 ) CPUFamily = " K5 Family ";
   if (StoreFamily == 25 ) CPUFamily = " K6 Family ";
   if (StoreFamily == 26 ) CPUFamily = " K6-2 ";
   if (StoreFamily == 27 ) CPUFamily = " K6-3 ";
   if (StoreFamily == 28 ) CPUFamily = " AMD Athlon� Processor Family ";
   if (StoreFamily == 29 ) CPUFamily = " AMD� Duron� Processor ";
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
   if (StoreFamily == 120 ) CPUFamily = " Crusoe� TM5000 Family ";
   if (StoreFamily == 121 ) CPUFamily = " Crusoe� TM3000 Family ";
   if (StoreFamily == 128 ) CPUFamily = " Weitek ";
   if (StoreFamily == 130 ) CPUFamily = " Itanium� Processor ";
   if (StoreFamily == 144 ) CPUFamily = " PA-RISC Family ";
   if (StoreFamily == 145 ) CPUFamily = " PA-RISC 8500 ";
   if (StoreFamily == 146 ) CPUFamily = " PA-RISC 8000 ";
   if (StoreFamily == 147 ) CPUFamily = " PA-RISC 7300LC ";
   if (StoreFamily == 148 ) CPUFamily = " PA-RISC 7200 ";
   if (StoreFamily == 149 ) CPUFamily = " PA-RISC 7100LC ";
   if (StoreFamily == 150 ) CPUFamily = " PA-RISC 7100 ";
   if (StoreFamily == 160 ) CPUFamily = " V30 Family ";
   if (StoreFamily == 176 ) CPUFamily = " Pentium� III Xeon� ";
   if (StoreFamily == 177 ) CPUFamily = " Pentium� III Processor with Intel� SpeedStep� Technology ";
   if (StoreFamily == 178 ) CPUFamily = " Pentium� 4 ";
   if (StoreFamily == 179 ) CPUFamily = " Intel� Xeon� ";
   if (StoreFamily == 180 ) CPUFamily = " AS400 Family ";
   if (StoreFamily == 181 ) CPUFamily = " Intel� Xeon� processor MP ";
   if (StoreFamily == 182 ) CPUFamily = " AMD AthlonXP� Family ";
   if (StoreFamily == 183 ) CPUFamily = " AMD AthlonMP� Family ";
   if (StoreFamily == 184 ) CPUFamily = " Intel� Itanium� 2 ";
   if (StoreFamily == 185 ) CPUFamily = " AMD Opteron� Family ";
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

function GetClientSU()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Client Service Update: ";
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey1 = "HKLM\\Software\\Nortel\\SECC\\SU Record\\SUID";
	var regkey2 = "HKLM\\Software\\Nortel\\SECC\\SU Record\\Time";
	var regkey3 = "HKLM\\Software\\Nortel\\SECC\\SU Record\\Date";
	
	TotalMessage += (shell.RegRead( regkey1 ));
	TotalMessage += ( " installed on ");
	TotalMessage += (shell.RegRead( regkey3 ));
	TotalMessage += " at " + (shell.RegRead( regkey2 ) + cr );
	return( TotalMessage );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
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

function GetSU1()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Service Update: ";
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey1 = "HKLM\\Software\\Nortel\\SCCS\\SU Record\\SUID";
	TotalMessage += (shell.RegRead( regkey1 ));
	return( TotalMessage + cr );
    }
    catch( e )
    {
        TotalMessage = "";
	return( TotalMessage );
    }
}	

function GetSU2()
{
    try
    {
        var cr = "\r\n";
	var TotalMessage = "Service Update: ";
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey1 = "HKLM\\Software\\Nortel\\SCCS\\PEP Record\\SUID";
	TotalMessage += (shell.RegRead( regkey1 ));
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
	var TotalMessage = "CCMS Info:" + cr;
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
	TotalMessage += (shell.RegRead( regkey6 ) + cr +cr );
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
	var TotalMessage = "License Manager Info:" + cr;
	var shell = new ActiveXObject( "WScript.Shell" );
       	var regkey2 = "HKLM\\Software\\Nortel\\Setup\\LM_IP_PORT_Primary";
  	var regkey3 = "HKLM\\Software\\Nortel\\Setup\\LM_IP_PORT_Secondary";
	var regkey4 = "HKLM\\Software\\Nortel\\Setup\\LM_IP_Primary";
  	var regkey5 = "HKLM\\Software\\Nortel\\Setup\\LM_IP_Secondary";
	var regkey6 = "HKLM\\Software\\Nortel\\Setup\\LM_Package";
	var regkey7 = "HKLM\\Software\\Nortel\\Setup\\LM_StandbyServer";
	var regkey8 = "HKLM\\Software\\Nortel\\Setup\\LM_OpenQueue";
	
       
	TotalMessage += "Package: ";
	TotalMessage += (shell.RegRead( regkey6 ) + cr );
	TotalMessage += "Primary IP: ";
	IP = (shell.RegRead( regkey4 ) );
	TotalMessage += IP + "\:";
	TotalMessage += (shell.RegRead( regkey2 ) + cr );
	TotalMessage += (PingIPs(IP)) + cr ;
	TotalMessage += "Secondary IP: ";
	TotalMessage += (shell.RegRead( regkey5 ) );
	TotalMessage += "\: ";
	TotalMessage += (shell.RegRead( regkey3 ) + cr );
	TotalMessage += "Open Queue: ";
	TotalMessage += (shell.RegRead( regkey8 ) + cr );
	TotalMessage += "Standby Server: ";
	TotalMessage += (shell.RegRead( regkey7 ) + cr + cr );
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

function GetPEPS()
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

	TotalMessage = "All Installed PEPS:";
	cr = "\r\n";
	
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
	   TotalMessage += KeyName + cr;
   	} 
	}
	catch (e)
	{ return ""; }
return (TotalMessage);
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
        return( "" );
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
        return( "" );
    }
}

function GetCPPEPS()
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

	TotalMessage = "Call Pilot Installed Patches: \r\n";
	cr = "\r\n";
	
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
        return( "" );
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
function NetStat()
{
   var Wshell = new ActiveXObject("WScript.Shell");
   var command = "cmd.exe /c netstat.exe -n>C:\\netstat.log"
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
        return( "" );
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
   catch ( e )
   {
        return ("");
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
catch (e)
{
return ("");
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
catch (e)
{
return ("");
}
return (TotalMessage);

}