var g_SL_Version = "10006";
var g_SL_SubVersion = "0";
var g_SL_LeadProgrammer = "Sergej N. Chagovets";
var g_SL_SoftwareTesters = "Helen U. Bogatyreva";
var g_SL_TestOrganization = "JSC Tulamashzavod";
var g_SL_SoftwareLicense = "OSS - Open Source Software";
var g_SL_FullVersion = g_SL_Version + "." + g_SL_SubVersion;
var gLanguage="en";
var gSysLibLocation="";
////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////
var gSysLibDBMSSrv="172.17.0.1;172.17.0.5";
var gStartupWait = 30;
var gScriptStartupWait = 30;
var gScriptSLStartupWait = 30;
var gScriptSLSysLibProcessWait = 900;	//15 minutes...
var gScriptEngine = "wscript.exe";
var gSystemSysLibLocation = "\\\\sirius.space\\SYSVOL\\space\\Policies\\{31B2F340-016D-11D2-945F-00C04FB984F9}\\MACHINE\\Scripts\\Startup\\SysLib.vbs";
var gUserSysLibLocation = "\\\\sirius.space\\SYSVOL\\space\\Policies\\{31B2F340-016D-11D2-945F-00C04FB984F9}\\USER\\Scripts\\Logon\\SysLib.vbs";
//(At IIS server for *.vbs file required specification of MIME type for vbs file extension as (with file charset specification): text/html;charset=windows-1251)
var gSysLibAltrnativeLocation = "https://sirius.space/syslib.vbs";
var gRemoveSysLibDownloadedFromAlternateLoc = 1;
var gSysLibSrcName = "SysLib.vbs";
var gSysLibExecTryCount = 3;
////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////
var g_oSink;
var g_oSvc;
var g_ScriptEngine;
var g_ComputerName;
var g_User_SID_Cached;
var g_OSVersion;
//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//' Constants for opening files
//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
var OpenFileForReading = 1;
var OpenFileForWriting = 2;
var OpenFileForAppending = 8;
var TristateUseDefault = -2;	//Opens the file using the system default.
var TristateTrue = -1;		//Opens the file as Unicode.
var TristateFalse =  0;		//Opens the file as ASCII
//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//' Constants for Windows Registry
//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
var HKEY_CLASSES_ROOT = 0x80000000;
var HKEY_CURRENT_CONFIG = 0x80000005;
var HKEY_CURRENT_USER = 0x80000001;
var HKEY_DYN_DATA = 0x80000006;
var HKEY_LOCAL_MACHINE = 0x80000002;
var HKEY_PERFORMANCE_DATA = 0x80000004;
var HKEY_USERS = 0x80000003;
var KEY_WOW64_64KEY = 0x00000100;
var KEY_WOW64_32KEY = 0x00000200;
var REG_NONE = 0;                        //' No value type
var REG_SZ = 1;                          //' Unicode nul terminated string
var REG_EXPAND_SZ = 2;                   //' Unicode nul terminated string
var REG_BINARY = 3;                      //' Free form binary
var REG_DWORD = 4;                       //' 32-bit number
var REG_DWORD_LITTLE_ENDIAN = 4;         //' 32-bit number (same as REG_DWORD)
var REG_DWORD_BIG_ENDIAN = 5;            //' 32-bit number
var REG_LINK = 6;                        //' Symbolic Link (unicode)
var REG_MULTI_SZ = 7;                    //' Multiple Unicode strings
var REG_RESOURCE_LIST = 8;               //' Resource list in the resource map
var REG_FULL_RESOURCE_DESCRIPTOR = 9;    //' Resource list in the hardware description
var REG_RESOURCE_REQUIREMENTS_LIST = 10;
var DBG_STACK_PREFIX = "	";
var DBG_LOG_FILESIZE = 4096000;
var MOD_ADLER = 65521;
////////////////////////////////////////////
//Ulian calendar support constants
////////////////////////////////////////////
var RTOH = 3.819718634205491995951628829999;
var DTR_Const = 1.74532925199433E-02;
var RTD_Const = 57.2957795130823;
var RTS_Const = 206264.806247096;       	/* arc seconds per radian */
var STR_Const = 4.84813681109536E-06;    	/* radians per arc second */
var Pi_Const = 3.14159265358979;
////////////////////////////////////////////
//WMI constants
////////////////////////////////////////////
var wbemFlagReturnImmediately = 0x10;
var wbemFlagForwardOnly = 0x20;
function f_GetFileDirectory(In_FullFileName)
{
    	var currentDir = new String;
	try 
	{
		currentDir = In_FullFileName;
	    	currentDir = currentDir.match(/.+\\/g)[0];
    		currentDir = currentDir.substring(0,currentDir.length-1);
	}
	catch(ll_Error)
	{
		errorHandler(ll_Error, true, "f_GetFileDirectory");
	}
	return currentDir;
}
function errorHandler(err, in_exit_script, in_location)
{
	var l_exit_script = (typeof in_exit_script === "undefined") ? true : in_exit_script;
	switch (gLanguage.toLowerCase())
	{
	case "en":
		if (in_location != "" & in_location != undefined & in_location != null) {
			WScript.Echo ("Error occured in: [" + in_location + "].\n" + "Number:        " + decimalToHexString(err.number) + "\n" + "Description:    " + err.description);
		} else {
			WScript.Echo ("Error.\n" + "Number:        " + decimalToHexString(err.number) + "\n" + "Description:    " + err.description);
		}
		break;
	case "ru":
		if (in_location != "" & in_location != undefined & in_location != null) {
			WScript.Echo ("Ïðîèçîøëà îøèáêà â : [" + in_location + "].\n" + "Íîìåð:        " + decimalToHexString(err.number) + "\n" + "Îïèñàíèå:    " + err.description);
		} else {
			WScript.Echo ("Ïðîèçîøëà îøèáêà.\n" + "Íîìåð:        " + decimalToHexString(err.number) + "\n" + "Îïèñàíèå:    " + err.description);
		}
		break;
	}
	if (l_exit_script)
	{
        	WScript.Quit(err.number);
	}
}
function decimalToHexString(number, hex_prefix)
{
	if (number < 0)
	{
		number = 0xFFFFFFFF + number + 1;
	}
	hex_prefix = (typeof hex_prefix === "undefined") ? "0x" : hex_prefix;
	if (hex_prefix !== "")
	{
		return hex_prefix + number.toString(16).toUpperCase();
	}
	else
	{
		return number.toString(16).toUpperCase();
	}
}
function f_StartProcessMonitor()
{
	try 
	{
//	 	var objShell = new ActiveXObject('Shell.Application');
// 		var objWShell = new ActiveXObject('WScript.Shell');
		//' Create an object sink
		g_oSink = WScript.CreateObject('WbemScripting.SWbemSink','SLS_');
		//' Connect to WMI and the cimv2 namespace, and obtain
		//' an SWbemServices object
//		var oSvc = GetObject("winmgmts:\\\\.\\root\\cimv2");
		g_oSvc = GetObject("winmgmts:\\\\.\\root\\CIMV2");
//		var locator = new ActiveXObject("WbemScripting.SWbemLocator");
//		g_oSvc = locator.ConnectServer(null, "root\\CIMV2");
		//' Query for all Win32_Process objects
//		g_oSvc.ExecQueryAsync (g_oSink, 'select * from __instancecreationevent  within 3 where TargetInstance isa "Win32_Process"');
		g_oSvc.ExecQueryAsync (g_oSink, 'SELECT * FROM __InstanceOperationEvent WITHIN 2 WHERE TargetInstance ISA "Win32_Process"', 'WQL', 128, null, null);
//		oSvc.ExecQueryAsync (g_oSink, "Select * from Win32_Process");
//		g_oSvc.ExecQueryAsync (g_oSink, "Select * from Win32_Process");
		WScript.sleep (10);	
	}
	catch(ll_Error)
	{
		errorHandler(ll_Error, true, "f_StartProcessMonitor");
	}
	return;
}
function SLS_OnObjectReady(oInst, octx)
{
	try 
	{
//		var objItems = String(oInst.Name);
//		WScript.Echo (objItems);
		WScript.Echo (oInst.Path_.Class);
		WScript.Echo ("Got Instance: " + oInst.Name);
	}
	catch(ll_Error)
	{
		WScript.Echo ("Error!");
		return;
	}
	return;
}
function SLS_OnCompleted(HResult, oErr, oCtx)
{
	WScript.Echo ("ExecQueryAsync completed");
	g_oSink.Cancel();
	g_oSink = null;
	g_oSvc = null;
	return;
}
function SLS_OnProgress(iUpperBound, iCurrent, strMessage, objWbemAsyncContext)
{
	WScript.Echo ("ExecQueryAsync in progress: " + strMessage);
	return;
}
/*
' The sink subroutine to handle the OnObjectReady 
' event. This is called as each object returns.
sub sink_OnObjectReady(oInst, octx)
    WScript.Echo "Got Instance: " & oInst.Name
end sub
' The sink subroutine to handle the OnCompleted event.
' This is called when all the objects are returned. 
' The oErr parameter obtains an SWbemLastError object,
' if available from the provider.
sub sink_OnCompleted(HResult, oErr, oCtx)
    WScript.Echo "ExecQueryAsync completed"
    bdone = true
end sub
*/
function f_GetWin32Processes(In_PCName, In_ShowPID, In_ShowProcessName, In_ShowCommandLine, In_GetProcOwner, In_GetProcOwnerSID, In_ProcElemDelim, In_RowDelim)
{
	var objWMI, colItems, objItem, output, output_SID;
	var list = "Ñïèñîê ïðîöåññîâ Windows\n\n";
	try
	{
		if (In_PCName == null || In_PCName == "")
		{
			l_PCName = "."
		} else {	
			l_PCName = In_PCName
		}       	
		if (In_ProcElemDelim == null || In_ProcElemDelim == "")
		{
			l_ProcElemDelim = "\t\t"
		} else {	
			l_ProcElemDelim = In_ProcElemDelim
		}       	
		if (In_RowDelim == null || In_RowDelim == "")
		{
			l_RowDelim = "\n"
		} else {	
			l_RowDelim = In_RowDelim
		}       	
		objWMI = GetObject("winmgmts:\\\\" + l_PCName + "\\root\\cimv2");
		//Ôîðìèðóåì êîëëåêöèþ ïðîöåññîâ ñ ïîìîùüþ êëàññà Win32_Process
		colItems = new Enumerator(objWMI.ExecQuery("Select * from Win32_Process"));
		for (; !colItems.atEnd(); colItems.moveNext())
		{
			objItem = colItems.item();
			if (In_GetProcOwner)
			{
				output_user = objItem.ExecMethod_("GetOwner").User;
				output_domain = objItem.ExecMethod_("GetOwner").Domain;
				output = output_domain + "\\" + output_user;
	
			} else {
				output_user = "";
				output_domain = "";
				output = "";
			}
			if (In_GetProcOwnerSID)
			{
				output_SID = objItem.ExecMethod_("GetOwnerSid").Sid + l_ProcElemDelim;
			} else {
				output_SID = "";
			}
			if (In_ShowPID)
			{
				output_PID = objItem.Handle + l_ProcElemDelim;
			} else {
				output_PID = "";
			}
			if (In_ShowCommandLine)
			{
				output_CMD = objItem.CommandLine + l_ProcElemDelim;
			} else {
				output_CMD = "";
			}
			if (In_ShowProcessName)
			{
				output_Name = objItem.Name + l_ProcElemDelim;
			} else {
				output_Name = "";
			}
			list += output_PID + output_Name + output_CMD + output_SID + output + l_RowDelim;
		}
		return list;
	}
	catch(ll_Error)
	{
		errorHandler(ll_Error, true, "f_GetWin32Processes");
		return "";
	}
}
function f_FindProcessesByCmdLine(In_PCName, In_SearchCriteria)
{
	var objWMI, colItems, objItem, output, output_SID;
	try
	{
		if (In_PCName == null || In_PCName == "")
		{
			l_PCName = "."
		} else {	
			l_PCName = In_PCName
		}       	
		objWMI = GetObject("winmgmts:\\\\" + l_PCName + "\\root\\cimv2");
		colItems = new Enumerator(objWMI.ExecQuery("Select Handle, Name, CommandLine from Win32_Process"));
		for (; !colItems.atEnd(); colItems.moveNext())
		{
			objItem = colItems.item();
			var llTestName = trim(objItem.Name.toString().toLowerCase());
			var llTestCmd = "";
			if (objItem.CommandLine != null) {
				llTestCmd = trim(objItem.CommandLine.toString().toLowerCase());
			}
			var llTestHandle = parseInt(objItem.Handle, 10);
			var llTestCriteria = trim(In_SearchCriteria.toString().toLowerCase());
			if (llTestName.indexOf(llTestCriteria) != -1) {
				return llTestHandle;
			}
			if (llTestCmd.indexOf(llTestCriteria) != -1) {
				return llTestHandle;
			}
		}
		return 0;
	}
	catch(ll_Error)
	{
		errorHandler(ll_Error, true, "f_FindProcessesByCmdLine");
		return "";
	}
}
function f_IsFileExist(In_File_Name) 
{
	var l_FSO;
	try
	{
		l_FSO = WScript.CreateObject("Scripting.FileSystemObject");
		if (l_FSO.FileExists(In_File_Name))
		{
			return true;
		} else {
			return false;
		}
	}
	catch (ll_Error)
	{
		errorHandler(ll_Error, true, "f_IsFileExist");
		return false;
	}
}
function f_IsFolderExist(In_Folder_Name)
{
	var l_FSO
	try
	{
		l_FSO = WScript.CreateObject("Scripting.FileSystemObject");
		if (l_FSO.FolderExists(In_Folder_Name))
		{
			l_FSO = null;
			return true;
		} else {
			l_FSO = null;
			return false;
		}
	}
	catch (ll_Error)
	{
		l_FSO = null;
		errorHandler(ll_Error, true, "f_IsFolderExist");
		return false;
	}
}
function ReadRegistry (In_Registry_Path) {
	var WshShell;
	try {
		WshShell = WScript.CreateObject("WScript.Shell");
		return WshShell.RegRead(In_Registry_Path);
	}
	catch (ll_Error)
	{
		if (ll_Error.number == -2147024894) {
			return "";
		} else {
			errorHandler(ll_Error, true, "ReadRegistry");
			return false;
		}
	}
}
function WriteRegistry (In_Registry_Path, In_Value, In_Type) {
	var WshShell;
	try 
	{
		WshShell = WScript.CreateObject("WScript.Shell");
		WshShell.RegWrite (In_Registry_Path, In_Value, In_Type);
	}
	catch (ll_Error)
	{
		errorHandler(ll_Error, true, "WriteRegistry");
		return false;
	}
}
function ltrim(InString) {
	try{
		if (InString !== null & InString !== undefined) {
		var lBuffer = "";
		var lTextBegins = false;
		for (var is_local=0; is_local < InString.toString().length; ++is_local) {
			if ((InString.toString().charAt(is_local) == " " || InString.toString().charAt(is_local) == "\t") && !lTextBegins) {
			} else {
				lTextBegins = true;
				lBuffer += InString.toString().charAt(is_local);
			}
		}
		return lBuffer;
		} else {
			return "";
		}
	}
	catch (ll_Error)
	{
		errorHandler(ll_Error, true, "ltrim");
		return "";
	}
}
function rtrim(InString) {
	try{
		if (InString !== null & InString !== undefined) {
			var lBuffer = "";
			var lTextBegins = false;
			for (var is_local=(InString.toString().length-1); is_local >= 0 ; --is_local) {
				if ((InString.toString().charAt(is_local) == " " || InString.toString().charAt(is_local) == "\t") && !lTextBegins) {
				} else {
					lTextBegins = true;
					lBuffer = InString.toString().charAt(is_local) + lBuffer;
				}
			}
			return lBuffer;
		} else {
			return "";
		}
	}
	catch (ll_Error)
	{
		errorHandler(ll_Error, true, "rtrim");
		return "";
	}
}
function trim(InString) {
	try{
		return ltrim(rtrim(InString));
	}
	catch (ll_Error)
	{
		errorHandler(ll_Error, true, "trim");
		return "";
	}
}
function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}
function f_DecodeWindowsProductType (In_Type) {
//	f_DecodeWindowsProductType = "N/A"
	try
	{
		if (isNumeric(In_Type)) {
			switch (Math.floor(In_Type)) {
			case 1:
				return "Work station";
			case 2:
				return "Domain controller";
			case 3:
				return "Server";
			default:
				return "N/A";
			}
		} else {
			//Old case (WinNT, ServerNT, LanManNT)
			switch (In_Type.toLowerCase()) {
			case "winnt":
				return "Work station";
			case "lanmannt":
				return "Domain controller";
			case "servernt":
				return "Server";
			default:
				return "N/A";
			}
		}
	}
	catch (ll_Error)
	{
		errorHandler(ll_Error, true, "f_DecodeWindowsProductType");
		return false;
	}
}
function f_validate_str_param (In_Test_Param, In_DefaultValue) {
	//This functio allow password char multiplication, if In_Hide_Char present several chars...
	//By default - default string value is empty string...
	var l_DefaultValue = "";
	if (In_DefaultValue === null || In_DefaultValue === undefined) {
		l_DefaultValue = "";
	} else {
		if (trim(In_DefaultValue) == "") {
			l_DefaultValue = "";
		} else {
			l_DefaultValue = In_DefaultValue;
		}
	}
	if (In_Test_Param === null || In_Test_Param === undefined) {
		return l_DefaultValue;
	} else {
		if (trim(In_Test_Param) == "") {
			return l_DefaultValue;
		} else {
			return In_Test_Param;
		}
	}
}
function f_string(strLength, strChar){
	try{
		var lBuffer = "";
		for (var i_local=1; i_local<=Math.floor(strLength); ++i_local) {
			lBuffer += strChar;
		}
		return lBuffer;
	}
	catch (ll_Error)
	{
		errorHandler(ll_Error, true, "f_string");
		return false;
	}
}
function f_Normalize_OS_Version_Number (In_Version_Number, In_MajorLen, In_MinorLen, In_RevisionLen) {
	var ll_RetVal="";
	try {
	//f_rnd_tmr = timer
	l_Version_Number = f_validate_str_param (In_Version_Number, "0.0.0");
	l_MajorLen = Math.floor(f_validate_str_param (In_MajorLen, "2"));
	l_MinorLen = Math.floor(f_validate_str_param (In_MinorLen, "2"));
	l_RevisionLen = Math.floor(f_validate_str_param (In_RevisionLen, "5"));
	ll_RetVal = "";
	ll_OS_VerArr = l_Version_Number.split(".");
	if (ll_OS_VerArr.length > 0) {
		for (ia_local = 0; ia_local < ll_OS_VerArr.length; ++ia_local) {
			switch (ia_local) {
			case 0:
				ll_RetVal = f_string(Math.floor(l_MajorLen) - ll_OS_VerArr[ia_local].toString().length, "0") + ll_OS_VerArr[ia_local].toString();
				break;
			case 1:
				ll_RetVal = ll_RetVal + ("." + f_string(Math.floor(l_MinorLen) - ll_OS_VerArr[ia_local].toString().length,"0") + ll_OS_VerArr[ia_local].toString());
				break;
			case 2:
				ll_RetVal = ll_RetVal + "." + f_string(Math.floor(l_RevisionLen) - ll_OS_VerArr[ia_local].toString().length,"0") + ll_OS_VerArr[ia_local].toString();
				break;
			default:
				ll_RetVal = ll_RetVal + "." + ll_OS_VerArr[ia_local].toString();
				break;
			}
		}
	}
		ll_LastErrorSuffix = "";
	}
	catch (ll_Error)
	{
		errorHandler(ll_Error, false);
		ll_LastErrorSuffix = " Last error ID: [" + ll_Error.number + " (" + ll_Error.description + ")].";
	}
	if (ll_LastErrorSuffix.lenght > 0 ) {
		WScript.Echo (ll_LastErrorSuffix);
	}
	return ll_RetVal;
}
function f_DecodeWindowsSerialNumber(In_DigitalProductID) {
	var digitalProductId = In_DigitalProductID;
        var keyStartIndex = 51;	
        var keyEndIndex = keyStartIndex + 15;
        var digits="B,C,D,F,G,H,J,K,M,P,Q,R,T,V,W,X,Y,2,3,4,6,7,8,9".split(",")	// Possible alpha-numeric characters in product key.
        var decodeLength = 29;	// Length of decoded product key
        var decodeStringLength = 15; // Length of decoded product key in byte-form
        var BinaryKey = new Array (15);// product key in byte-form
        var J = 0;
	var sn="";
        for (var i = keyStartIndex + 1; i <= keyEndIndex; ++i) {
            BinaryKey[J] = digitalProductId[i];
            J = J + 1;
        }
        for (var i = decodeLength; i >= 0; --i) {
            if ((i + 1) % 6 == 0) {
                if (i + 1 != 30) {
                    sn = "-" + sn;
                }
            } else {
                var digitMapIndex = 0;
		var byteValue;
                for (J = decodeStringLength - 1; J >= 0; --J) {
                    byteValue = ((digitMapIndex * 256) | BinaryKey[J]);
                    BinaryKey[J] = parseInt(byteValue / 24, 10);
                    digitMapIndex = byteValue % 24;
                }
                sn = digits[digitMapIndex] + sn;
            }
        }
	return sn;
}
function f_Read_Registry_Ex(strComputer, hKEY_PTR, strKeyPath, strValueName, strDataType, iError) {
	if (strComputer == "" | strComputer === null | strComputer === undefined) {
		strComputer = ".";
	}
	var ll_SimpleReg = false;
	var ll_RetVal = "";
	var ll_Phase = 0;
	try {
		ll_Phase = 100;
		lStdRegProv = GetObject("winmgmts:{impersonationLevel=impersonate,(Security,Restore)}!\\\\" + strComputer + "\\root\\default:StdRegProv");
	}
	catch (ll_Error)
	{
		ll_SimpleReg = true;
	}
	if (ll_SimpleReg) {
		ll_Phase = 101;
		lStdRegProv = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\" + strComputer + "\\root\\default:StdRegProv");
	}
	try {
		ll_Phase = 102;
		oReg=lStdRegProv;
		switch (strDataType.toString()) {
		case "4":
		case "REG_DWORD":
			//GetDWORDValue
			ll_Phase = 4;
			var objInParams = oReg.Methods_("GetDWORDValue").InParameters.SpawnInstance_();
	    		objInParams.hDefKey = hKEY_PTR;
    			objInParams.sSubKeyName = strKeyPath;
    			objInParams.sValueName = strValueName;
	    		var objOutParams = oReg.ExecMethod_("GetDWORDValue", objInParams);
			ll_RetVal = objOutParams.uValue;
			break;
		case "2":
		case "REG_EXPAND_SZ":
			ll_Phase = 2;
			//GetExpandedStringValue
			var objInParams = oReg.Methods_("GetExpandedStringValue").InParameters.SpawnInstance_();
	    		objInParams.hDefKey = hKEY_PTR;
    			objInParams.sSubKeyName = strKeyPath;
    			objInParams.sValueName = strValueName;
	    		var objOutParams = oReg.ExecMethod_("GetExpandedStringValue", objInParams);
			ll_RetVal = objOutParams.sValue;
			break;
		case "1":
		case "REG_SZ":
			ll_Phase = 1;
			//GetStringValue
			var objInParams = oReg.Methods_("GetStringValue").InParameters.SpawnInstance_();
	    		objInParams.hDefKey = hKEY_PTR;
    			objInParams.sSubKeyName = strKeyPath;
    			objInParams.sValueName = strValueName;
	    		var objOutParams = oReg.ExecMethod_("GetStringValue", objInParams);
			ll_RetVal = objOutParams.sValue;
			break;
		case "7":
		case "REG_MULTI_SZ":
			//GetMultiStringValue
			ll_Phase = 7;
			var objInParams = oReg.Methods_("GetMultiStringValue").InParameters.SpawnInstance_();
    			objInParams.hDefKey = hKEY_PTR;
    			objInParams.sSubKeyName = strKeyPath;
	    		objInParams.sValueName = strValueName;
    			var objOutParams = oReg.ExecMethod_("GetMultiStringValue", objInParams);
			if (objOutParams.sValue !== null & objOutParams.sValue !== undefined) {
				ll_RetVal = objOutParams.sValue.toArray();
			} else {
				ll_RetVal = "";
			}
			break;
		default:
			//REG_BINARY, OTHERS
			//Assume that this is binary value
			ll_Phase = 3;
			var objInParams = oReg.Methods_("GetBinaryValue").InParameters.SpawnInstance_();
    			objInParams.hDefKey = hKEY_PTR;
	    		objInParams.sSubKeyName = strKeyPath;
    			objInParams.sValueName = strValueName;
    			var objOutParams = oReg.ExecMethod_("GetBinaryValue", objInParams);
			if (objOutParams.uValue !== null & objOutParams.uValue !== undefined) {
				ll_RetVal = objOutParams.uValue.toArray();
			} else {
				ll_RetVal = 0;
			}
			break;
		}
		return ll_RetVal;
	}
	catch (ll_Error){
		errorHandler(ll_Error, true, "f_Read_Registry_Ex, phase:[" + ll_Phase + "]")
		return "";
	}
}
//https://msdn.microsoft.com/en-us/library/aa384833%28v=vs.85%29.aspx
//http://forum.shelek.ru/index.php/topic,24216.0.html
//call a provider method using the Scripting API and SWbemServices.ExecMethod
function WMIMethod(wmiObject, methodName) {
  var _method;
  var _inputParameters;
  var _outputParameters;
  //output parameters.
  this.getOutputParameters = function() {
    return _outputParameters;
  }
  //input parameters
  this.getInputParameters = function() {
    return _inputParameters;
  }
  //Setup input parameters.
  this.setInputParameters = function(parameters) {
    for(key in parameters) {
      eval("_inputParameters." + key + " = parameters[key]");
    }
  }
  //Execution selected method
  this.execute = function() {
    _outputParameters = wmiObject.ExecMethod_(_method.Name, _inputParameters);
  }
  //initialization...
  _method = wmiObject.Methods_.Item(methodName);
  _inputParameters = _method.InParameters.SpawnInstance_();
  _outputParameters = null;
}
function Registry() {
  var _registryProvider = null;
  //Methods.
  //getSubKeys section
  this.getSubKeys = function(rootKey, keyPath) {
    var keys = new Array();
    var enumKeys = new WMIMethod(_registryProvider, "EnumKey");
    enumKeys.setInputParameters({ hDefKey: rootKey, sSubKeyName: keyPath });
    enumKeys.execute();
    with(enumKeys.getOutputParameters()) {
      if(sNames != null) {
        keys = sNames.toArray();
      }
    }
    return keys;
  }
  //getValues section
  this.getValues = function(rootKey, keyPath, valueName) {
    var result = new Array();
    var valuesNames, valuesTypes;
    var enumValues = new WMIMethod(_registryProvider, "EnumValues");
    enumValues.setInputParameters({ hDefKey: rootKey, sSubKeyName: keyPath });
    enumValues.execute();
    with(enumValues.getOutputParameters()) {
      if(sNames != null && Types != null) {
        valuesNames = sNames.toArray();
        valuesTypes = Types.toArray();
        for(var i = 0; i < valuesNames.length; ++i) {
          result.push({ name: valuesNames[i], type: valuesTypes[i] });
        }
      }
    }
    return result;
  }
  //getSpecified value
  this.read = function(rootKey, keyPath, valueName) {
    var values = this.getValues(rootKey, keyPath);
    var valueType = null;
    for(i in values) {
      if(values[i].name == valueName) {
        valueType = values[i].type;
        break;
      }
    }
    var getMethodName = "";
    switch(valueType) {
      case REG_SZ:
        getMethodName = "GetStringValue";
        break;
      case REG_EXPAND_SZ:
        getMethodName = "GetExpandedStringValue";
        break;
      case REG_BINARY:
        getMethodName = "GetBinaryValue";
        break;
      case REG_DWORD:
        getMethodName = "GetDWORDValue";
        break;
      case REG_MULTI_SZ:
        getMethodName = "GetMultiStringValue";
        break;
    }
    var get = new WMIMethod(_registryProvider, getMethodName);
    get.setInputParameters({ hDefKey: rootKey, sSubKeyName: keyPath, sValueName: valueName });
    get.execute();
    with(get.getOutputParameters()) {
      switch(valueType) {
        case REG_SZ:
        case REG_EXPAND_SZ:
          return sValue;
        case REG_BINARY:
          return uValue.toArray();
        case REG_DWORD:
          return uValue;
        case REG_MULTI_SZ:
          return sValue.toArray();
      }
    }
  }
  //Creating registry provider
  function _createRegistryProvider() {
    var locator = new ActiveXObject("WbemScripting.SWbemLocator");
    var connection = locator.ConnectServer(null, "root\\default");
    return connection.Get("StdRegProv");
  }
  //Initialization
  _registryProvider = _createRegistryProvider();
}
/*
// Êîíñòðóêòîð îáúåêòà äëÿ ðàáîòû ñ ðååñòðîì.
// Êîíñòàíòû.
// Êîðåíü ëîêàëüíîé ìàøèíû.
Registry.prototype.HKLM = 0x80000002;
// Òèïû çíà÷åíèé.
Registry.prototype.REG_SZ = 1;
Registry.prototype.REG_EXPAND_SZ = 2;
Registry.prototype.REG_BINARY = 3;
Registry.prototype.REG_DWORD = 4;
Registry.prototype.REG_MULTI_SZ = 7;
function Registry() {
  // Àòðèáóòû îáúåêòà.
  // COM-îáúåêò - WMI ïðîâàéäåð ðååñòðà.
  var _registryProvider = null;
  // Ìåòîäû îáúåêòà.
  // Âîçâðàùàåò íåïîñðåäñòâåííûå êëþ÷è óêàçàííîé âåòâè äåðåâà.
  this.getSubKeys = function(rootKey, keyPath) {
    var keys = new Array();
    var enumKeys = new WMIMethod(_registryProvider, "EnumKey");
    enumKeys.setInputParameters({ hDefKey: rootKey, sSubKeyName: keyPath });
    enumKeys.execute();
    with(enumKeys.getOutputParameters()) {
      if(sNames != null) {
        keys = sNames.toArray();
      }
    }
    return keys;
  }
  // Âîçâðàùàåò çíà÷åíèÿ êëþ÷à óêàçàííîé âåòâè äåðåâà.
  this.getValues = function(rootKey, keyPath, valueName) {
    var result = new Array();
    var valuesNames, valuesTypes;
    var enumValues = new WMIMethod(_registryProvider, "EnumValues");
    enumValues.setInputParameters({ hDefKey: rootKey, sSubKeyName: keyPath });
    enumValues.execute();
    with(enumValues.getOutputParameters()) {
      if(sNames != null && Types != null) {
        valuesNames = sNames.toArray();
        valuesTypes = Types.toArray();
        for(var i = 0; i < valuesNames.length; ++i) {
          result.push({ name: valuesNames[i], type: valuesTypes[i] });
        }
      }
    }
    return result;
  }
  // Âîçâðàùàåò çíà÷åíèå èç êëþ÷à.
  this.read = function(rootKey, keyPath, valueName) {
    var values = this.getValues(rootKey, keyPath);
    var valueType = null;
    for(i in values) {
      if(values[i].name == valueName) {
        valueType = values[i].type;
        break;
      }
    }
    var getMethodName = "";
    switch(valueType) {
      case Registry.prototype.REG_SZ:
        getMethodName = "GetStringValue";
        break;
      case Registry.prototype.REG_EXPAND_SZ:
        getMethodName = "GetExpandedStringValue";
        break;
      case Registry.prototype.REG_BINARY:
        getMethodName = "GetBinaryValue";
        break;
      case Registry.prototype.REG_DWORD:
        getMethodName = "GetDWORDValue";
        break;
      case Registry.prototype.REG_MULTI_SZ:
        getMethodName = "GetMultiStringValue";
        break;
    }
    var get = new WMIMethod(_registryProvider, getMethodName);
    get.setInputParameters({ hDefKey: rootKey, sSubKeyName: keyPath, sValueName: valueName });
    get.execute();
    with(get.getOutputParameters()) {
      switch(valueType) {
        case Registry.prototype.REG_SZ:
        case Registry.prototype.REG_EXPAND_SZ:
          return sValue;
        case Registry.prototype.REG_BINARY:
          return uValue.toArray();
        case Registry.prototype.REG_DWORD:
          return uValue;
        case Registry.prototype.REG_MULTI_SZ:
          return sValue.toArray();
      }
    }
  }
  // Âñïîìîãàòåëüíûå ôóíêöèè.
  function _createRegistryProvider() {
    var locator = new ActiveXObject("WbemScripting.SWbemLocator");
    var connection = locator.ConnectServer(null, "root\\default");
    return connection.Get("StdRegProv");
  }
  // Èíèöèàëèçàöèÿ
  _registryProvider = _createRegistryProvider();
}
*/
//Usage: var subKeys = registry.getSubKeys(Registry.prototype.HKLM, "SOFTWARE\\Microsoft");
function f_GetOSInfo(In_PC_Name, In_Info_Class){
	var iError;
//	try 
//	{
		if ((In_PC_Name == null) || (In_PC_Name == ""))
		{
			strComputer = ".";
		} else {
			strComputer = In_PC_Name;
		}
		ll_NoCachedValueFound = true;
		l_Suffix = "";
		if (ll_NoCachedValueFound) {
			objWMIService = GetObject("winmgmts:"  + "{impersonationLevel=impersonate,(Shutdown,RemoteShutdown)}!\\\\" + strComputer + "\\root\\cimv2");
			colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem");
			var enumItems = new Enumerator(colOperatingSystems);
			for (; !enumItems.atEnd(); enumItems.moveNext()) {
				var ObjOperatingSystem = enumItems.item();
				switch (In_Info_Class)	{
				case 0:
//					return f_DecodeWindowsSerialNumber(f_Read_Registry_Ex(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", "DigitalProductId", "REG_BINARY", iError));
					var registry = new Registry();
					return f_DecodeWindowsSerialNumber(registry.read (HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", "DigitalProductId"));
					registry = null;
					break;
				case 1:
					if (trim(ObjOperatingSystem.OtherTypeDescription) != "") {
						return trim(ObjOperatingSystem.Caption) + " " + trim(ObjOperatingSystem.OtherTypeDescription)  + " " + trim(ObjOperatingSystem.CSDVersion) + " (" + trim(ObjOperatingSystem.Version) + "). Role: (" + trim(f_DecodeWindowsProductType(ObjOperatingSystem.ProductType)) + "). Serial number: (" + trim(ObjOperatingSystem.SerialNumber) + ")";
					} else {
						return trim(ObjOperatingSystem.Caption) + " " + trim(ObjOperatingSystem.CSDVersion) + " (" + trim(ObjOperatingSystem.Version) + "). Role: (" + trim(f_DecodeWindowsProductType(ObjOperatingSystem.ProductType)) + "). Serial number: (" + trim(ObjOperatingSystem.SerialNumber) + ")";
					}
					//' & vbcrlf & ObjOperatingSystem.Name
					break;
				case 2:
					return ObjOperatingSystem.Version;
					break;
				case 3:
					return f_DecodeWindowsProductType(ObjOperatingSystem.ProductType);
					if (err.number != 0) {
						f_GetOSInfo  = f_DecodeWindowsProductType(ReadRegistry ("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\ProductOptions\\ProductType"));
					}
					break;
				case 4:
					return ObjOperatingSystem.SerialNumber;
					break;
				case 5:
					if (trim(ObjOperatingSystem.OtherTypeDescription) != "") {
						return trim(ObjOperatingSystem.Caption) + " " + trim(ObjOperatingSystem.OtherTypeDescription) + " " + trim(ObjOperatingSystem.CSDVersion) + " (" + trim(ObjOperatingSystem.Version) + ")." ;
					} else {
						return trim(ObjOperatingSystem.Caption) + " " + trim(ObjOperatingSystem.CSDVersion) + " (" + trim(ObjOperatingSystem.Version) + ")." ;
					}
					break;
				case 6:
					//'OS Architechture
					//'Windows Vista attruibute...
					return ObjOperatingSystem.OSArchitechture;
					break;
				case 7:
					//'SP Major Version
					return ObjOperatingSystem.ServicePackMajorVersion;
					break;
				case 8:
					//'same as 3, but return integer value for certificates table...
					switch (ObjOperatingSystem.ProductType) {
					case 1:
						return 0;
					case 2:
						return 2;
					case 3:
						return 1;
					}
						switch (ReadRegistry ("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\ProductOptions\\ProductType").toLowerCase()) {
						case "winnt":
							return 0;
						case "servernt":
							return 1;
						case "lanmannt":
							return 2;
						default:
							//'Lets assume that this is minimal possible system. In our case - WinNT
							return 0;
						}
					break;
				case 9:
					return ReadRegistry ("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\ProductOptions\\ProductType").toLowerCase();
					break;
				case 10:
					return ObjOperatingSystem.Caption;
					break;
				case 11:
					return ObjOperatingSystem.OtherTypeDescription;
					break;
				case 12:
					return ObjOperatingSystem.CSDVersion;
					break;
				case 13:
					//'Get OS.Version as long 
					ll_OS_VerArr = ObjOperatingSystem.Version.split(".");
					if (ll_OS_VerArr.length > 1) {
						return (Math.floor(ll_OS_VerArr[0]) * Math.floor(10000000)) + (Math.floor(ll_OS_VerArr[1]) * Math.floor(100000)) + Math.floor(ll_OS_VerArr[2]);
					} else {
						return 0;
					}
					break;
				case 14:
					//'Lets return extended OS Version as string like 010.000.009926
					return f_Normalize_OS_Version_Number (ObjOperatingSystem.Version, 2, 2, 5);
					break;
				case 15:
					//'Lets return extended OS Version as string like 010.000.009926
					return f_Normalize_OS_Version_Number (ObjOperatingSystem.Version, 3, 3, 6);
					break;
				case 16:
					//'Major version
					ll_OS_VerArr = ObjOperatingSystem.Version.split(".");
					if (ll_OS_VerArr.length > 0) {
						return Math.floor(ll_OS_VerArr[0]);
					} else {
						f_GetOSInfo = 0
					}
					break;
				case 17:
					//'Minor version
					ll_OS_VerArr = ObjOperatingSystem.Version.split(".");
					if ((ll_OS_VerArr.length > 0)) {
						return Math.floor(ll_OS_VerArr[1]);
					} else {
						return 0;
					}
					break;
				case 18:
					//'Revision version
					ll_OS_VerArr = ObjOperatingSystem.Version.split(".");
					if ((ll_OS_VerArr.length > 1)) {
						return Math.floor(ll_OS_VerArr[2]);
					} else {
						return 0;
					}
					break;
				default:
					f_GetOSInfo = ObjOperatingSystem.PlusVersionNumber & vbcrlf
					break;
				}
			}
			objWMIService = null;
			colOperatingSystems = null;
		}
//	}
//	catch (ll_Error)
//	{
//		errorHandler(ll_Error, true, "");
		return false;
//	}
}
function GetScriptEngineInfo()
{
    var s;
    s = ""; // Build string with necessary info.
    s += ScriptEngine() + " Version ";
    s += ScriptEngineMajorVersion() + ".";
    s += ScriptEngineMinorVersion() + ".";
    s += ScriptEngineBuildVersion();
    return(s);
}
function f_calendar_to_julian (In_Year, In_Month, In_Day) {
//	'on error resume next
	var l_Year = clng(f_validate_str_param (In_Year, year(now())));
	var l_Month = clng(f_validate_str_param (In_Month, month(now())));
	var l_Day = clng(f_validate_str_param (In_Day, day(now())));
	var ll_RetVal = "";
	var ll_Y, ll_A, ll_b, ll_C, ll_E, ll_m;
	var ll_J;
	/* The origin should be chosen to be a century year
	 * that is also a leap year.  We pick 4801 B.C.
	 */
	ll_Y = (l_Year) + (4800);
	if ((l_Year) < 0) {
    		ll_Y = (ll_Y) + 1;
	}
	/* The following magic arithmetic calculates a sequence
	 * whose successive terms differ by the correct number of
	 * days per calendar month.  It starts at 122 = March'; January
	 * and February come after December.
	 */
	ll_m = l_Month;
	if (ll_m <= 2) {
	    ll_m = ll_m + 12;
	    ll_Y = ll_Y - 1;
	}
	ll_E = Math.floor(((306) * ((ll_m) + (1))) / (10));
	ll_A = Math.floor((ll_Y) / (100));  /* number of centuries */
	var l_CalendarType = 0;	//'0 - julius; 1 - gregor
	if (l_Year <= 1582) {
		if (l_Year = 1582) {
        		if (l_Month < 10) {
				l_CalendarType = 0;	//	'0 - julius; 1 - gregor
        		}
        		if (l_Month > 10) {
				l_CalendarType = 1;	//	'0 - julius; 1 - gregor
        		}
        		if (l_Day >= 15) {
				l_CalendarType = 1;	//	'0 - julius; 1 - gregor
        		}
    		}
		switch (parseInt(l_CalendarType, 10)) {
		case 0:
    			ll_b = -38;
			break;
		case 1:
			ll_b = Math.floor(((ll_A) / (4)) - (ll_A));
		}
	} else {
		/* -number of century years that are not leap years */
		ll_b = Math.floor(((ll_A) / (4)) - (ll_A));
	}
	ll_C = Math.floor(((36525) * (ll_Y)) / (100)); /* Julian calendar years and leap years */
	/* Add up these terms, plus offset from J 0 to 1 Jan 4801 B.C.
	 * Also fudge for the 122 days from the month algorithm.
	 */
	ll_J = (ll_b) + (ll_C) + (ll_E) + (l_Day) - ((32167) + ((5)/(10)));
	ll_RetVal = ll_J;
	return ll_RetVal;
}
function isNumeric(obj)
{
	return !isNaN(parseFloat(obj)) && isFinite(obj);
}
function cdbl(In_DoubleValue)
{
	if (isNumeric(In_DoubleValue)) {
		return parseFloat(In_DoubleValue);
	} else {
		return parseFloat(0);
	}
}
function clng (In_IntegerValue) {
	if (isNumeric(In_IntegerValue)) {
		return parseInt(In_IntegerValue, 10);
	} else {
		return parseInt(0, 10);
	}
}
/*
'WS2 DE-404 engine part...
''!! Warning function is incomplete!... Modification and full testing required!!!
*/
function f_julian_to_calendar_str (In_JulianDate, In_TimeZoneOffset, In_UT_TDT_to_Local, In_UT_to_Local) {
	//'on error resume next
	var l_JulianDate = (f_validate_str_param (In_JulianDate, "0"));
	var l_TimeZoneOffset = parseInt(f_validate_str_param (In_TimeZoneOffset, "180"), 10);
	var UT_TDT_to_Local = (f_validate_str_param (In_UT_TDT_to_Local, "False"));
	var UT_to_Local = (f_validate_str_param (In_UT_to_Local, "False"));
	var aTimeZoneOffset_J = (-1) * (l_TimeZoneOffset) / (1440);
	var ll_RetVal = "";
	var ll_month, ll_day;
	var ll_year, ll_A, ll_C, ll_D, ll_X, ll_Y, ll_jd;
	var BC;
	var DD;
	var J_Bkp;
	var aDeltaT;
	var T;
	var ll_months = new Array(12);
	var ll_days = new Array(7);
	var jtocal_string = "";
	var cyear = 1986;
	ll_month = 1;
	ll_day = 1;
	yerend = 0;
	ll_months[0] = "January";
	ll_months[1] = "February";
	ll_months[2] = "March";
	ll_months[3] = "April";
	ll_months[4] = "May";
	ll_months[5] = "June";
	ll_months[6] = "July";
	ll_months[7] = "August";
	ll_months[8] = "September";
	ll_months[9] = "October";
	ll_months[10] = "November";
	ll_months[11] = "December";
	ll_days[0] = "Sunday";
	ll_days[1] = "Monday";
	ll_days[2] = "Tuesday";
	ll_days[3] = "Wednesday";
	ll_days[4] = "Thursday";
	ll_days[5] = "Friday";
	ll_days[6] = "Saturday";
	var J_Bkp = l_JulianDate;
	UT_to_Local = (0);
	if (l_JulianDate == UT && (UT_to_Local)) {
    		T = cdbl(2000) + (l_JulianDate - J2000_Const) / cdbl((365) + (25)/(100));
    		aDeltaT = (deltat(T) / cdbl(86400));
	} else {
    		aDeltaT = 0;
	}
	if (UT_TDT_to_Local) {
    		l_JulianDate = l_JulianDate - aTimeZoneOffset_J + aDeltaT;
	} else {
    		l_JulianDate = l_JulianDate + aDeltaT;
	}
	if (l_JulianDate < 1721425.5) {
    		/* January 1.0, 1 A.D. */
    		BC = 1;
	} else {
    		BC = 0;
	}
	ll_jd = Math.floor(l_JulianDate + cdbl(5)/cdbl(10));  /* round Julian date up to integer */
	/* Find the number of Gregorian centuries
	 * since March 1, 4801 B.C.
	*/
	ll_A = fix((cdbl(100) * cdbl(ll_jd) + cdbl(3204500)) / cdbl(3652425));
	/* Transform to Julian calendar by adding in Gregorian century years
	 * that are not leap years.
	 * Subtract 97 days to shift origin of JD to March 1.
	 * Add 122 days for magic arithmetic algorithm.
	 * Add four years to ensure the first leap year is detected.
	*/
	ll_C = cdbl(ll_jd) + cdbl(1486);
	if (cdbl(ll_jd) >= cdbl(cdbl(2299160) + (cdbl(5)/cdbl(10)))) {
    		ll_C = Math.floor(ll_C + ll_A - (cdbl(ll_A) / cdbl(4)));
	} else {
    		ll_C = ll_C + 38;
	}
	/* Offset 122 days, which is where the magic arithmetic
	* month formula sequence starts (March 1 = 4 * 30.6 = 122.4).
	*/
	ll_D = Math.floor((cdbl(100) * cdbl(ll_C) - cdbl(12210)) / cdbl(36525));
	/* Days in that many whole Julian years */
	ll_X = Math.floor((cdbl(36525) * ll_D) / cdbl(100));
	/* Find month and day. */
	ll_Y = Math.floor(((ll_C - ll_X) * cdbl(100)) / cdbl(3061));
	ll_day = parseInt(ll_C - ll_X - fix(((cdbl(306) * cdbl(ll_Y)) / cdbl(10))), 10)
	ll_month = parseInt(ll_Y - 1, 10)
	if (ll_Y > 13) {
    		ll_month = ll_month - 12;
	}
	/* Get the year right. */
	ll_year = ll_D - cdbl(4715);
	if (ll_month > 2) {
    		ll_year = ll_year - 1;
	}
	/* Day of the week. */
	ll_A = (ll_jd + 1) % 7;
	/* Fractional part of day. */
	ll_DD = ll_day + l_JulianDate - ll_jd + 0.5;
	/* post the year. */
	cyear = ll_year;
	if (BC) {
    		ll_year = -1*ll_year + 1;
    		cyear = -ll_year;
        	jtocal_string = jtocal_string + ll_year + " B.C. ";
	} else {
        	jtocal_string = jtocal_string + ll_year + " ";
	}
	ll_day = parseInt(ll_DD, 10);
	jtocal_string = jtocal_string + ll_months[ll_month - 1] + ", " + ll_day + " " + ll_days[parseInt(ll_A, 10)] + " ";
	/* Flag last or first day of year */
	if (((ll_month = 1) && (ll_day = 1)) || ((ll_month = 12) && (ll_day = 31))) {
    		yerend = 1;
	} else {
    		yerend = 0;
	}
	/* Display fraction of calendar day
	 * as clock time.
	 */
	ll_A = parseInt(ll_DD, 10);
	ll_DD = ll_DD - ll_A;
	aTimeZoneOffset_J = (-1) * cdbl(l_TimeZoneOffset) / cdbl(1440);
	jtocal_string = jtocal_string + f_RadianToHMS(cdbl(2) * cdbl(Pi_Const) * cdbl(ll_DD), "", true);
	if (l_JulianDate == TDT) {
		jtocal_string = jtocal_string + "TDT" + "\n";
	} else {
		if (l_JulianDate == UT) {
			//'MsgBox "UT" ') /* Universal Time */
			jtocal_string = jtocal_string + "UT" + "\n";
		} else {
			//'MsgBox ("Undefined!")
			if (l_JulianDate + aTimeZoneOffset_J == TDT) {
				jtocal_string = jtocal_string + "Local time (TDT based)" + "\n";
			} else {
				if (l_JulianDate + aTimeZoneOffset_J == UT) {
					jtocal_string = jtocal_string + "Local time (UT based)" + "\n";
				} else {
					if ((l_JulianDate + aTimeZoneOffset_J - aDeltaT) == UT) {
						jtocal_string = jtocal_string + "Local time (UT=>TDT based)" + "\n";
					} else {
						jtocal_string = jtocal_string + "Undefined!" + "\n";
					}
				}
			}
		}
	}
	l_JulianDate = J_Bkp;
	ll_RetVal = jtocal_string;
	return ll_RetVal;
}
function f_RadianToHMS (In_RadianValue, In_HMS_Separator, In_ShowPartsOfSeconds) {
//	'on error resume next
	var l_RadianValue = f_validate_str_param (In_RadianValue, "0");
	var l_HMS_Separator = f_validate_str_param (In_HMS_Separator, "");
	var l_ShowPartsOfSeconds = (f_validate_str_param (In_ShowPartsOfSeconds, "True"));
	var ll_RetVal = "";
	var ll_H, ll_m;
	var ll_sint, ll_sfrac;
	var ll_s;
	ll_s = (l_RadianValue) * (RTOH);
	if ((ll_s) < 0) {
    		ll_s = cdbl(ll_s) + cdbl(24)
	}
	ll_H = parseInt(ll_s, 10);
	ll_s = (ll_s) - (ll_H);
	ll_s = (ll_s) * (60);
	ll_m = parseInt(ll_s, 10);
	ll_s = (ll_s) - (ll_m);
	ll_s = (ll_s) * (60);
	/* Handle shillings and pence roundoff. */
	var ll_sfrac = Math.floor(((100000) * (ll_s) + (0.5)));
	if ((ll_sfrac) >= (6000000)) {
    		ll_sfrac = (ll_sfrac) - (6000000);
    		ll_m = ll_m + 1;
    		if ((ll_m) >= (60)) {
        		ll_m = (ll_m) - (60);
        		ll_H = ll_H + 1;
      		}
  	}
	ll_sint = Math.floor((ll_sfrac) / (100000));
	ll_sfrac = (ll_sfrac) - (ll_sint) * (100000);
	if (trim(l_HMS_Separator) == "") {
		if (l_ShowPartsOfSeconds) {
			ll_RetVal = ll_H + "h " + String(2 - ll_m.toString().length, "0") + ll_m + "m " + String(2 - ll_sint.toString().length, "0") + ll_sint + "." + String(5 - ll_sfrac.toString().length, "0") + ll_sfrac + " sec ";
		} else {
			ll_RetVal = ll_H + "h " + String(2 - ll_m.toString().length, "0") + ll_m + "m " + String(2 - ll_sint.toString().length, "0") + ll_sint + " sec ";
		}
	} else {
		if (l_ShowPartsOfSeconds) {
			ll_RetVal = ll_H + trim(l_HMS_Separator) + String(2 - ll_m.toString().length, "0") + ll_m + trim(l_HMS_Separator) + String(2 - ll_sint.toString().length, "0") + ll_sint + "." + String(5 - ll_sfrac.toString().length, "0") + ll_sfrac;
		} else {
			ll_RetVal = ll_H + trim(l_HMS_Separator) + String(2 - ll_m.toString().length, "0") + ll_m + trim(l_HMS_Separator) + String(2 - ll_sint.toString().length, "0") + ll_sint;
		}
	}
	return ll_RetVal;
}
function f_ConvertHMSToJulianDate (In_Hour, In_Minute, In_Second, In_TimeZoneOffset) {
//	'on error resume next
	try {
		l_Hour = parseInt(f_validate_str_param (In_Hour, hour(now())), 10);
		l_Minute = parseInt(f_validate_str_param (In_Minute, minute(now())), 10);
		l_Second = parseInt(f_validate_str_param (In_Second, second(now())), 10);
		l_TimeZoneOffset = parseInt(f_validate_str_param (In_TimeZoneOffset, "180"), 10);
		ll_RetVal = (3600 * l_Hour + 60 * l_Minute + l_Second) / 86400;
		aTimeZoneOffset_J = (-1) * l_TimeZoneOffset / 1440;
		ll_RetVal = ll_RetVal + aTimeZoneOffset_J;
	}
	catch (ll_Error)
	{
		errorHandler(ll_Error, true, "f_ConvertHMSToJulianDate");
		return 0;
	}
	return ll_RetVal;
}
function f_SetHeartBeats() {
	//'on error resume next
	var HeartBeats = 0;
	HeartBeats = f_calendar_to_julian (year(now()), month(now()), day(now())) + f_ConvertHMSToJulianDate (hour(now()), minute(now()), second(now()), "")
	return HeartBeats;
}
function now () {
	return new Date();
}
function year(In_Date) {
	return new Date(In_Date).getFullYear();
}
function month(In_Date) {
	return ((new Date(In_Date).getMonth()) + 1);
}
function day(In_Date) {
	return new Date(In_Date).getDate();
}
function hour(In_Date) {
	return new Date(In_Date).getHours();
}
function minute(In_Date) {
	return new Date(In_Date).getMinutes();
}
function second(In_Date) {
	return new Date(In_Date).getSeconds();
}
function weekday(In_Date) {
	return new Date(In_Date).getDay();
}
function GetUserName() {
	try {
		var WshNetwork = WScript.CreateObject("WScript.Network");
		return WshNetwork.UserName;
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "GetUserName");
		return "N/A";
	}
}
function GetDomainName() {
	try {
		var WshNetwork = WScript.CreateObject("WScript.Network");
		return WshNetwork.UserDomain;
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "GetDomainName");
		return "N/A";
	}
}
function GetFullUserName() {
	try {
		var WshNetwork = WScript.CreateObject("WScript.Network");
		return WshNetwork.UserDomain + "\\" + WshNetwork.UserName;
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "GetFullUserName");
		return "N/A";
	}
}
function GetFullUserName_Ex(In_Full_Name_Style) {
	//0 - Domain\User, 1 - User@Domain, 2 - Domain, 3 - Name, 4 - PC Domain, 5 - PC Name, 6 - {Domain|PCName}\User, 1 - User@{Domain|PCName}
	try {
		var Name_Style, l_dotpos, l_Domain_Name, l_UserName;
		var ll_GetFullUserName_Ex = "N/A";
		if ((In_Full_Name_Style == 0) | (In_Full_Name_Style == null)) {
			Name_Style=0;
		} else {
			Name_Style=clng(In_Full_Name_Style);
		}
		var WshNetwork = WScript.CreateObject("WScript.Network");
		switch (clng(Name_Style)) {
		case 0:
			ll_GetFullUserName_Ex = WshNetwork.UserDomain + "\\" + WshNetwork.UserName;
			break;
		case 1:
			ll_GetFullUserName_Ex = WshNetwork.UserName + "@" + WshNetwork.UserDomain;
			break;
		case 2:
			//Alternate methods (only for PC Domain! Complex case in multi domain networks...):
			//HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MSMQ\Parameters\setup\MachineDomain REG_SZ, Also MachineDomainFQDN REG_SZ
			//HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AltDefaultDomainName REG_SZ, Also CachePrimaryDomain REG_SZ, Also DefaultDomainName REG_SZ
			//HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DomainCache\%DOMAIN_NAME%	REG_SZ (ALL CACHED Domains....)
			//!!	HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\Tcpip\Parameters\Domain	REG_SZ, Also NV Domain 	REG_SZ (FQDN domain name)
			ll_GetFullUserName_Ex = WshNetwork.UserDomain;
			if (trim(ll_GetFullUserName_Ex) == "") {
				//Lets try to detect default dpomain name...
				l_Domain_Name = ReadRegistry ("HKEY_CURRENT_USER\\Volatile Environment\\USERDNSDOMAIN");
				l_dotpos = l_Domain_Name.indexOf(".");
				if (clng(l_dotpos) != -1) {
					l_Domain_Name = l_Domain_Name.substr(1,clng(l_dotpos)-1);
				}
				ll_GetFullUserName_Ex = l_Domain_Name;
			}
			if (trim(ll_GetFullUserName_Ex) == "") {
				//Lets try to detect default dpomain name...
				l_Domain_Name = ReadRegistry ("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon\\CachePrimaryDomain");
				ll_GetFullUserName_Ex = l_Domain_Name;
			}
			if (trim(ll_GetFullUserName_Ex) == "") {
				//Lets try to detect default dpomain name...
				l_Domain_Name = ReadRegistry ("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Domain");
				l_dotpos = l_Domain_Name.indexOf(".");
				if (clng(l_dotpos) != -1) {
					l_Domain_Name = l_Domain_Name.substr(1,clng(l_dotpos)-1);
				}
				ll_GetFullUserName_Ex = l_Domain_Name;
			}
			if (trim(ll_GetFullUserName_Ex) == "") {
				//Lets try to detect default dpomain name...
				l_Domain_Name = ReadRegistry ("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Telephony\\DomainName");
				l_dotpos = l_Domain_Name.indexOf(".");
				if (clng(l_dotpos) != -1) {
					l_Domain_Name = l_Domain_Name.substr (1,clng(l_dotpos)-1);
				}
				GetFullUserName_Ex = l_Domain_Name;
			}
			if (trim(ll_GetFullUserName_Ex) == "") {
				//Lets try to detect default domain name...
				l_Domain_Name = ReadRegistry ("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\MSMQ\\Parameters\\setup\\MachineDomain");
				l_dotpos = l_Domain_Name.indexOf(".");
				if (clng(l_dotpos) != -1) {
					l_Domain_Name = l_Domain_Name.substr(1, clng(l_dotpos)-1);
				}
				ll_GetFullUserName_Ex = l_Domain_Name;
			}
			break;
		case 3:
			//#Try #1 (Main)
			ll_GetFullUserName_Ex = WshNetwork.UserName;
			//#Try #2
			if (trim(ll_GetFullUserName_Ex) == "") {
				//Lets try to detect user name...
				l_Domain_Name = ReadRegistry ("HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Logon User Name");
				ll_GetFullUserName_Ex = l_Domain_Name;
			}
			//#Try #3 
			if (trim(ll_GetFullUserName_Ex) == "") {
				//Lets try to detect user name... (Key exist from Windows 2000 (5.0; In Win NT 4.0 Missing UserName) to Windows 8 (6.2)
				l_Domain_Name = ReadRegistry ("HKEY_CURRENT_USER\\Software\\Microsoft\\Active Setup\\Installed Components\\{44BBA840-CC51-11CF-AAFA-00AA00B6015C}\\UserName")
				ll_GetFullUserName_Ex = l_Domain_Name;
			}
			break;
		case 4:
			l_Domain_Name = ReadRegistry ("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon\\CachePrimaryDomain");
			ll_GetFullUserName_Ex = l_Domain_Name;
			if (trim(GetFullUserName_Ex) == "") {
				//Lets try to detect default domain name...
				l_Domain_Name = ReadRegistry ("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Domain");
				l_dotpos = l_Domain_Name.indexOf(".");
				if (clng(l_dotpos) != -1) {
					l_Domain_Name = l_Domain_Name.substr(1, clng(l_dotpos)-1);
				}
				ll_GetFullUserName_Ex = l_Domain_Name;
			}
			if (trim(ll_GetFullUserName_Ex) == "") {
				//Lets try to detect default domain name...
				l_Domain_Name = ReadRegistry ("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Telephony\\DomainName");
				l_dotpos = l_Domain_Name.indexOf(".");
				if (clng(l_dotpos) != -1) {
					l_Domain_Name = l_Domain_Name.substr (1,clng(l_dotpos)-1);
				}
				ll_GetFullUserName_Ex = l_Domain_Name;
			}
			if (trim(ll_GetFullUserName_Ex) == "") {
				//Lets try to detect default domain name...
				l_Domain_Name = ReadRegistry ("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\MSMQ\\Parameters\\setup\\MachineDomain");
				l_dotpos = l_Domain_Name.indexOf(".");
				if (clng(l_dotpos) != -1) {
					l_Domain_Name = l_Domain_Name.substr(1,clng(l_dotpos)-1);
				}
				ll_GetFullUserName_Ex = l_Domain_Name;
			}
			break;
		case 5:
	        	ll_GetFullUserName_Ex = WshNetwork.ComputerName;
			//#Try #2
			if (trim(ll_GetFullUserName_Ex) == "") {
				//Lets try to detect computer name...
				ll_GetFullUserName_Ex = ReadRegistry ("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\ComputerName\\ComputerName\\ComputerName");
			}
			break;
		case 6:
			//Like 0, but if WshNetwork.UserDomain is empty - using PC Name
			l_DomainName = WshNetwork.UserDomain;
			if (trim(l_DomainName) == "") {
				l_DomainName = WshNetwork.ComputerName;
				//#Try #2
				if (trim(l_DomainName) == "") {
					//Lets try to detect computer name...
					l_DomainName = ReadRegistry ("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\ComputerName\\ComputerName\\ComputerName");
				}
			}
			//#Try #1 (Main)
			l_UserName = WshNetwork.UserName;
			//#Try #2
			if (trim(l_UserName) == "") {
				//Lets try to detect default domain name...
				l_UserName = ReadRegistry ("HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Logon User Name");
			}
			//#Try #3 
			if (trim(l_UserName) == "") {
				//Lets try to detect default domain name... (Key exist from Windows 2000 (5.0; In Win NT 4.0 Missing UserName) to Windows 8 (6.2)
				l_UserName = ReadRegistry ("HKEY_CURRENT_USER\\Software\\Microsoft\\Active Setup\\Installed Components\\{44BBA840-CC51-11CF-AAFA-00AA00B6015C}\\UserName");
			}
			ll_GetFullUserName_Ex = l_DomainName  + "\\" + l_UserName;
			break;
		case 7:
			//Like 1, but if WshNetwork.UserDomain is empty - using PC Name
			l_DomainName = WshNetwork.UserDomain;
			if (trim(l_DomainName) == "") {
				l_DomainName = WshNetwork.ComputerName;
				//#Try #2
				if (trim(l_DomainName) == "") {
					//Lets try to detect computer name...
					l_DomainName = ReadRegistry ("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\ComputerName\\ComputerName\\ComputerName");
				}
			}
			//#Try #1 (Main)
			l_UserName = WshNetwork.UserName;
			//#Try #2
			if (trim(l_UserName) == "") {
				//Lets try to detect user name...
				l_UserName = ReadRegistry ("HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Logon User Name");
			}
			//#Try #3 
			if (trim(l_UserName) == "") {
				//Lets try to detect user name... (Key exist from Windows 2000 (5.0; In Win NT 4.0 Missing UserName) to Windows 8 (6.2)
				l_UserName = ReadRegistry ("HKEY_CURRENT_USER\\Software\\Microsoft\\Active Setup\\Installed Components\\{44BBA840-CC51-11CF-AAFA-00AA00B6015C}\\UserName");
			}
			ll_GetFullUserName_Ex = l_UserName + "@" + l_DomainName;
			break;
		default:
			ll_GetFullUserName_Ex = WshNetwork.UserName + "@" + WshNetwork.UserDomain;
			break;
		}
		return ll_GetFullUserName_Ex;
	}
	catch (ll_Error){
		errorHandler(ll_Error, true, "GetFullUserName_Ex");
		return 0;
	}
}
function f_GetPCDomainName(In_PCName) {
	try {
		if (trim(In_PCName) == "") {
			l_PCName = ".";
		} else {
			l_PCName = In_PCName;
		}
		var l_GetPCDomainName = "", iError = 0;
		var l_Domain_Name = f_Read_Registry_Ex(l_PCName, HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon", "CachePrimaryDomain", "REG_SZ", iError);
		if ((iError == 0) & l_Domain_Name != "") {
			l_GetPCDomainName = l_Domain_Name;
		}
		if (trim(l_GetPCDomainName) == "") {
			//Lets try to detect default dpomain name...
			l_Domain_Name = ReadRegistry ("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Domain");
			l_dotpos = l_Domain_Name.indexOf(".")
			if (clng(l_dotpos) != -1) {
				l_Domain_Name = l_Domain_Name.substr (1,clng(l_dotpos)-1);
			}
			l_GetPCDomainName = l_Domain_Name;
		}
		if (trim(l_GetPCDomainName) == "") {
			//Lets try to detect default dpomain name...
			l_Domain_Name = ReadRegistry ("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Telephony\\DomainName")
			l_dotpos = l_Domain_Name.indexOf(".");
			if (clng(l_dotpos) != -1) {
				l_Domain_Name = l_Domain_Name.substr (1,clng(l_dotpos)-1);
			}
			l_GetPCDomainName = l_Domain_Name;
		}
		if (trim(l_GetPCDomainName) == "") {
			//Lets try to detect default dpomain name...
			l_Domain_Name = ReadRegistry ("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\MSMQ\\Parameters\\setup\\MachineDomain")
			l_dotpos = l_Domain_Name.indexOf(".");
			if (clng(l_dotpos) != -1) {
				l_Domain_Name = l_Domain_Name.substr (1,clng(l_dotpos)-1);
			}
			l_GetPCDomainName = l_Domain_Name;
		}
		if (trim(l_GetPCDomainName) == "") {
			l_GetPCDomainName = "N/A";
		}
		return l_GetPCDomainName;
	}
	catch(ll_Error) {
		errorHandler(ll_Error, true, "f_GetPCDomainName");
		return 0;
	}
}
function GetComputerName () {
	try {
		var lComputerName = "";
		var lFuncPhase = 0;
		if (trim(g_ComputerName) != "") {
			lComputerName = g_ComputerName;
			return lComputerName;
		}
		lFuncPhase = 1;
		var WshNetwork = WScript.CreateObject("WScript.Network");
        	lComputerName = WshNetwork.ComputerName;
		//#Try #2
		lFuncPhase = 2;
		if (trim(lComputerName) == "") {
			//Lets try to detect computer name...
			lComputerName = ReadRegistry ("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\ComputerName\\ComputerName\\ComputerName");
		}
		lFuncPhase = 3;
		if (trim(lComputerName) != "") {
			if (trim(g_ComputerName) == "") {
				g_ComputerName = lComputerName;
			}
		}
		lFuncPhase = 4;
		return lComputerName;
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "GetComputerName, phase: [" + lFuncPhase + "]");
		return 0;
	}
}
function f_GetFastScriptEngine() {
	var Script, ScriptRun;
	try {
		Script = WScript.FullName;
		ScriptRun = (Script.substr(Script.lastIndexOf("\\")+1)).toLowerCase();
		//Main Procedure
		if (ScriptRun != "cscript.exe") {
			return "WScript";
		} else {
			return "CScript";
		}
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "f_GetFastScriptEngine");
		return 0;
	}
}
function f_StartupLogo() {
	try {
		if (g_ScriptEngine.toLowerCase() == "cscript") {
			WScript.Echo ("");
			WScript.Echo ("=========================================================================================");
			WScript.Echo ("SL monitor version: \t[" + g_SL_FullVersion + "] starting at: ["+ Date() + "]");
			WScript.Echo ("Script engine: \t\t[" + GetScriptEngineInfo() + "]");
			WScript.Echo ("-----------------------------------------------------------------------------------------");
			WScript.Echo ("Lead programmer: \t[" + g_SL_LeadProgrammer + "]");
			WScript.Echo ("Software testers: \t[" + g_SL_SoftwareTesters + "]");
			WScript.Echo ("Software deployed for: \t[" + g_SL_TestOrganization + "]");
			WScript.Echo ("-----------------------------------------------------------------------------------------");
			WScript.Echo ("License:\t\t[" + g_SL_SoftwareLicense + "]");
			WScript.Echo ("=========================================================================================");
			WScript.Echo ("");
		}
		return 0;
	}
	catch (ll_Error){
		errorHandler(ll_Error, true, "f_StartupLogo");
		return 0;
	}
}
function f_LowLevelInit() {
	try {
		g_ScriptEngine = f_GetFastScriptEngine();
		g_OSVersion  = f_GetOSInfo(".", 14);
		return;
	}
	catch (ll_Error){
		errorHandler(ll_Error, true, "f_LowLevelInit");
		return 0;
	}
}
function printf(In_String) {
	try {
		if (g_ScriptEngine.toLowerCase() == "cscript") {
			WScript.Echo (In_String);
		}
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "printf")
		return 0;
	}
}
function f_GetUserSID (In_PCName, In_UserName) {
	var l_DomainName, l_Result, l_WMIUserObj, l_BSPosition, l_AtPosition, l_UpdateCache, l_GetUserSID;
	l_GetUserSID = "";
	l_UpdateCache = false;
	//In case NT AUTHORITY - DOMAIN_NAME = PC_NAME !! (HOST\SYSTEM)
	try {
		if ((trim(In_PCName) == "" | trim(In_PCName) == ".") & trim(In_UserName) == "") {
			if (trim(g_User_SID_Cached) != "" & g_User_SID_Cached !== null & g_User_SID_Cached !== undefined) {
				l_GetUserSID = g_User_SID_Cached;
				return l_GetUserSID;
			} else {
				l_UpdateCache = true;
			}
		}
		l_PCName = f_validate_str_param (In_PCName, ".");
		ll_UserName = f_validate_str_param (In_UserName, GetFullUserName_Ex(6));
		l_BSPosition = ll_UserName.indexOf("\\");
		l_AtPosition = ll_UserName.indexOf("@");
		if (l_BSPosition > -1) {
			l_DomainName = ll_UserName.substr(0, l_BSPosition);
    			l_UserName = ll_UserName.substr(l_BSPosition + 1);
		} else {
			if (l_AtPosition > -1) {
				l_UserName = ll_UserName.substr(l_AtPosition);
    				l_DomainName = ll_UserName.substr(l_AtPosition + 1);
			} else {
				l_UserName = ll_UserName;
    				l_DomainName = WScript.CreateObject("WScript.Network").UserDomain;
			}
		}
	}
	catch (ll_Error){
		//Some internal Error processing
	}
	try {
		l_WMIUserObj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\" + l_PCName + "\\root\\cimv2:Win32_UserAccount.Domain='" + l_DomainName + "',Name='" + l_UserName + "'");
		l_Result = l_WMIUserObj.SID;
		l_GetUserSID = l_Result;
	}
	catch (ll_Error) {
		try {
			//Additional call
			l_WMIUserObj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\" + l_PCName + "\\root\\cimv2:Win32_Account.Domain='" + l_DomainName + "',Name='" + l_UserName + "'");
			if (l_WMIUserObj !== null & l_WMIUserObj !== undefined) {
				l_Result = l_WMIUserObj.SID;
			} else {
				//Lets try change domain namespace to local PC namespace...
				ll_PCName = GetFullUserName_Ex(5);
				l_WMIUserObj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\" + l_PCName + "\\root\\cimv2:Win32_UserAccount.Domain='" + ll_PCName + "',Name='" + l_UserName + "'");
				if (l_WMIUserObj !== null & l_WMIUserObj !== undefined) {
					l_Result = l_WMIUserObj.SID;
				} else {
					//Additional call
					l_WMIUserObj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\" + l_PCName + "\\root\\cimv2:Win32_Account.Domain='" + ll_PCName + "',Name='" + l_UserName + "'");
					if (l_WMIUserObj !== null & l_WMIUserObj !== undefined) {
						l_Result = l_WMIUserObj.SID;
					} else {
						l_Result = "";
					}
				}
			}
		}
		catch (ll_ErrorInternal)
		{
			//Simple on error resume next...
			l_Result = "";
		}
		l_WMIUserObj = null;
		l_GetUserSID = l_Result;
	}
	if (l_UpdateCache) {
		g_User_SID_Cached = l_GetUserSID;
	}
	return l_GetUserSID;
}
function sleep(In_Interval) {
	try {
		WScript.sleep(In_Interval*1000);
		return;
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "sleep")
		return 0;
	}
}
function f_ping (In_ServerName) {
	//Required Windows XP+
	try {
		var l_ping = 0;	//false
		var f_ping_ttl = 0;
		var WMISvc;
		var cPingResults;
		//Lets perform fast check for cached values...
		var l_AsyncDetectedHosts = "";
		if (trim(g_OSVersion) == "") {
			g_OSVersion  = f_GetOSInfo(".", 14);
		}
		if (g_OSVersion >= "05.01.02600") {
			cPingResults = GetObject("winmgmts:\\\\.\\root\\CIMV2").ExecQuery("Select * from Win32_PingStatus Where Address='" + In_ServerName + "'", "WQL", wbemFlagReturnImmediately | wbemFlagForwardOnly);
   			var pingResults = new Enumerator(cPingResults);
			for (; !pingResults.atEnd(); pingResults.moveNext()) {
				var oPing = pingResults.item();
				if (oPing.StatusCode == 0) {
					f_ping_ttl = oPing.ResponseTimeToLive;
					l_ping = 1;
				}
			}
		} else {
			//Windows 2k, Windows XP beta, Windows NT 4.0 case...
			var WshShell = WScript.CreateObject("WScript.Shell");
			var objFSO = WScript.CreateObject("Scripting.FileSystemObject");
			var lWshEnv = WshShell.Environment("Process");
			var l_temp_folder = lWshEnv("Temp");
			var objFile;
			var lFile;
			if (l_temp_folder != "") {
				WshShell.Run ("cmd /c ping " + In_ServerName + " -n 3 -w 300 >\"" + l_temp_folder + "\\ping_status.dat\"", 0, true);
				objFile = objFSO.OpenTextFile(l_temp_folder + "\\ping_status.dat", 1);
			} else {
				WshShell.Run ("cmd /c ping " + In_ServerName + " -n 3 -w 300 >ping_status.dat", 0, true);
				objFile = objFSO.OpenTextFile("ping_status.dat", 1);
			}
			strContents = objFile.ReadAll();
			objFile.Close();
			if (strContents.toLowerCase().indexOf("ttl=") != -1) {
				l_ping=1;
			}
			if (l_temp_folder != "") {
	  			lFile = objFSO.GetFile(l_temp_folder + "\\ping_status.dat");
			} else {
	  			lFile = objFSO.GetFile("ping_status.dat");
			}
	  		//Delete file
  			lFile.Delete (true); //Lets force file deletion
  			lFile = null;
			WshShell = null;
			objFSO = null;
			objFile = null;
			lWshEnv = null;
		}
		return l_ping;
	}
	catch(ll_Error) {
		errorHandler(ll_Error, true, "f_ping")
		return 0;
	}
}
function f_Process_win32PingStatusReturnCode(In_StatusCode) {
	try {
		var l_Process_win32PingStatusReturnCode = "";
		l_StatusCode = clng(f_validate_str_param (In_StatusCode, "-1"));
		switch (l_StatusCode) {
		case 0:
			l_Process_win32PingStatusReturnCode = "Success";
			break;
		case 11001:
			l_Process_win32PingStatusReturnCode = "Buffer Too Small";
			break;
		case 11002:
			l_Process_win32PingStatusReturnCode = "Destination Net Unreachable";
			break;
		case 11003:
			l_Process_win32PingStatusReturnCode = "Destination Host Unreachable";
			break;
		case 11004:
			l_Process_win32PingStatusReturnCode = "Destination Protocol Unreachable";
			break;
		case 11005:
			l_Process_win32PingStatusReturnCode = "Destination Port Unreachable";
			break;
		case 11006:
			l_Process_win32PingStatusReturnCode = "No Resources";
			break;
		case 11007:
			l_Process_win32PingStatusReturnCode = "Bad Option";
			break;
		case 11008:
			l_Process_win32PingStatusReturnCode = "Hardware Error";
			break;
		case 11009:
			l_Process_win32PingStatusReturnCode = "Packet Too Big";
			break;
		case 11010:
			l_Process_win32PingStatusReturnCode = "Request Timed Out";
			break;
		case 11011:
			l_Process_win32PingStatusReturnCode = "Bad Request";
			break;
		case 11012:
			l_Process_win32PingStatusReturnCode = "Bad Route";
			break;
		case 11013:
			l_Process_win32PingStatusReturnCode = "TimeToLive Expired Transit";
			break;
		case 11014:
			l_Process_win32PingStatusReturnCode = "TimeToLive Expired Reassembly";
			break;
		case 11015:
			l_Process_win32PingStatusReturnCode = "Parameter Problem";
			break;
		case 11016:
			l_Process_win32PingStatusReturnCode = "Source Quench";
			break;
		case 11017:
			l_Process_win32PingStatusReturnCode = "Option Too Big";
			break;
		case 11018:
			l_Process_win32PingStatusReturnCode = "Bad Destination";
			break;
		case 11032:
			l_Process_win32PingStatusReturnCode = "Negotiating IPSEC";
			break;
		case 11050:
			l_Process_win32PingStatusReturnCode = "General Failure";
			break;
		default:
			l_Process_win32PingStatusReturnCode = "Bad Win32_PingStatus code";
			break;
		}
		return l_Process_win32PingStatusReturnCode;
	}
	catch(ll_Error) {
		errorHandler(ll_Error, true, "f_Process_win32PingStatusReturnCode")
		return 0;
	}
}
function f_ping_ex (In_ServerNameList, In_ReturnResult, In_MaxSimultaneousPings, In_ListDelimeter) {
	//Required Windows XP+
	//In_ServerNameList - lets accept in several cases. Simple list and prepared string array...
	//In_ReturnResult - 0 - just return first successfull address...
	//In_ReturnResult - 1 - just return all successfull in list
	//In_ReturnResult - 2 - just return all successfull in array
	//In_ReturnResult - 3 - just return all in list with statuses
	//In_ReturnResult - 4 - just return all in array with statuses
	//In_ReturnResult - 5 - just return all in array with statuses and string status designations
	//In_ReturnResult - 6 - just return all in array with statuses and string status designations
	//In_ReturnResult - 7 - just return all in array with string status designations
	//In_ReturnResult - 8 - just return all in array with string status designations
	//In_MaxSimultaneousPings - Maximal number simultaneos pings (default 50)
	try {
		l_ReturnResult = clng(f_validate_str_param (In_ReturnResult, "0"));
		l_MaxSimultaneousPings = clng(f_validate_str_param (In_MaxSimultaneousPings, "50"));
		l_ListDelimeter = f_validate_str_param (In_ListDelimeter, ";");
		l_ping_ex = 0;		//false
		if (trim(g_OSVersion) == "") {
			g_OSVersion  = f_GetOSInfo(".", 14);
		}
		if (g_OSVersion >= "05.01.02600") {
			var l_Result= "";
			var l_SrvCount = 0;
			var l_PingQuery = "";
		if (In_ServerNameList instanceof Array) {
			lSrvList = In_ServerNameList;
		} else {
			lSrvList = In_ServerNameList.replace(/,/g,";").split(";");
		}
		if (lSrvList instanceof Array) {
			for (var i_local = 0; i_local < lSrvList.length; ++ i_local) {
				if (trim(lSrvList[i_local]) != "") {
					l_SrvCount = l_SrvCount + 1;
					if (trim(l_PingQuery) != "") {
						l_PingQuery = l_PingQuery + " or Address='" + trim(lSrvList[i_local]) + "'";
					} else {
						l_PingQuery = l_PingQuery + "Address='" + trim(lSrvList[i_local]) + "'";
					}
				}
				if (l_SrvCount >= l_MaxSimultaneousPings | (i_local == (lSrvList.length - 1) & l_SrvCount > 0)) {
					//Lets ping our targets...
					var cPingResults = GetObject("winmgmts://./root/cimv2").ExecQuery("Select * from win32_PingStatus where " + l_PingQuery, "WQL", wbemFlagReturnImmediately | wbemFlagForwardOnly);
					var pingResult = new Enumerator(cPingResults);
					for (; !pingResult.atEnd(); pingResult.moveNext()) {
						var oPing = pingResult.item();
						switch (clng(l_ReturnResult)) {
						case 0:
							if (oPing.StatusCode==0) {
								l_ping_ex = oPing.Address;
								return l_ping_ex;
							}
							break;
						case 1:
						case 2:
							if (oPing.StatusCode == 0) {
								l_Result = l_Result + oPing.Address + l_ListDelimeter;
							}
							break;
						case 3:
						case 4:
							l_Result = l_Result + oPing.Address + "-" + oPing.StatusCode + l_ListDelimeter;
							break;
						case 5:
						case 6:
							l_Result = l_Result + oPing.Address + "-" + oPing.StatusCode + "-" + f_Process_win32PingStatusReturnCode(oPing.StatusCode) + l_ListDelimeter;
							break;
						case 7:
						case 8:
							l_Result = l_Result + oPing.Address + "-" + f_Process_win32PingStatusReturnCode(oPing.StatusCode) + l_ListDelimeter;
						default:
							//like case 0
							if (oPing.StatusCode==0) {
								l_ping_ex = oPing.Address;
								return l_ping_ex;
							}
							break;
						}
					}
					l_SrvCount = 0;
                	                l_PingQuery = "";
				}
			}
		}
		switch (clng(l_ReturnResult)) {
		case 2:
		case 4:
		case 6:
		case 8:
			l_ping_ex = l_Result.split(l_ListDelimeter);
			break;
		default:
			l_ping_ex = l_Result;
			break;
		}
		return l_ping_ex;
	} else {
		//Windows 2000 case...
		//Windows 2k, Windows XP beta, Windows NT 4.0 case...
		var WshShell = WScript.CreateObject("WScript.Shell");
		var objFSO = WScript.CreateObject("Scripting.FileSystemObject");
		var lWshEnv = WshShell.Environment("Process");
		var l_temp_folder = lWshEnv("Temp");
		var l_Result= "";
		var l_SrvCount = 0;
		var l_PingQuery = "";
		if (In_ServerNameList instanceof Array) {
			lSrvList = In_ServerNameList;
		} else {
			lSrvList = In_ServerNameList.replace(/,/g,";").split(";");
		}
		var l_AlreadyTested="";
		//randomize timer
		var ll_UUID = parseInt(Math.random()*100, 10);
		for (i_local = 0; i_local < lSrvList.length; ++i_local) {
			if ((trim(lSrvList[i_local]) != "") & (l_AlreadyTested.indexOf(";" + trim(lSrvList[i_local].toLowerCase()) + ";" ) == -1)) {
				l_SrvCount = l_SrvCount + 1;
				l_AlreadyTested = l_AlreadyTested + ";" + trim(lSrvList[i_local].toLowerCase()) + ";";
				var objFile;
				if (l_temp_folder != "") {
					WshShell.Run ("cmd /c ping " + trim(lSrvList[i_local]) + " -n 2 -w 300 >\"" + l_temp_folder + "\\ping_status" + ll_UUID + ".dat\"", 0, true);
					objFile = objFSO.OpenTextFile(l_temp_folder + "\\ping_status" + ll_UUID + ".dat", 1);
				} else {
					WshShell.Run ("cmd /c ping " + trim(lSrvList[i_local]) + " -n 2 -w 300 >ping_status" + ll_UUID + ".dat", 0, true);
					objFile = objFSO.OpenTextFile("ping_status" + ll_UUID + ".dat", 1);
				}
				strContents = objFile.ReadAll();
				objFile.Close();
				if (strContents.toLowerCase().indexOf("ttl=") != -1) {
					switch (clng(l_ReturnResult)) {
					case 0:
						l_ping_ex = trim(lSrvList[i_local]);
						return l_ping_ex;
						break;
					case 1:
					case 2:
						l_Result = l_Result + trim(lSrvList[i_local]) + l_ListDelimeter;
						break;
					case 3:
					case 4:
						l_Result = l_Result + trim(lSrvList[i_local]) + "-" + 0 + l_ListDelimeter;
						break;
					case 5:
					case 6:
						l_Result = l_Result + trim(lSrvList[i_local]) + "-" + 0 + "-" + f_Process_win32PingStatusReturnCode(0) + l_ListDelimeter;
						break;
					case 7:
					case 8:
						l_Result = l_Result + trim(lSrvList[i_local]) + "-" + f_Process_win32PingStatusReturnCode(0) + l_ListDelimeter;
						break;
					default:
						//like case 0
						l_ping_ex = trim(lSrvList[i_local]);
						return l_ping_ex;
						break;
					}
				} else {
					switch (clng(l_ReturnResult)) {
					case 3:
					case 4:
						l_Result = l_Result + trim(lSrvList[i_local]) + "-" + 11010 + l_ListDelimeter;
						break;
					case 5:
					case 6:
						l_Result = l_Result + trim(lSrvList[i_local]) + "-" + 11010 + "-" + f_Process_win32PingStatusReturnCode(11010) + l_ListDelimeter;
						break;
					case 7:
					case 8:
						l_Result = l_Result + trim(lSrvList[i_local]) + "-" + f_Process_win32PingStatusReturnCode(11010) + l_ListDelimeter;
						break;
					}
				}
			}
		}
		var lFile;
		if (l_temp_folder != "") {
	  		lFile = objFSO.GetFile(l_temp_folder + "\\ping_status" + ll_UUID + ".dat");
		} else {
	  		lFile = objFSO.GetFile("ping_status" + ll_UUID + ".dat");
		}
	  	// Delete file
  		lFile.Delete (true); 		//Lets force file deletion
  		lFile = null;
		WshShell = null;
		objFSO = null;
		objFile = null;
		lWshEnv = null;
		switch (clng(l_ReturnResult)) {
		case 2:
		case 4:
		case 6:
		case 8:
			l_ping_ex = l_Result.split(l_ListDelimeter);
			break;
		default:
			l_ping_ex = l_Result;	
			break;
		}
		return l_ping_ex;
	}
	}
	catch(ll_Error) {
		errorHandler(ll_Error, true, "f_ping_ex")
		return 0;
	}
}
function f_TestSQLDBConnection(In_DB_Srv) {
	try {
		var sConn;
		if (f_ping (In_DB_Srv) != 0) {
			sConn = "Provider=sqloledb;connect timeout=4;server=" + In_DB_Srv + ";database=master;uid=sql_ping;pwd=sql_ping;";
			var lConnection = WScript.CreateObject("ADODB.Connection");
			lConnection.ConnectionTimeout = 4;
			lConnection.CommandTimeout = 0;
		var lConn = lConnection.Open (sConn);
		l_Result = true;
		lConnection.Close();
		lConnection = null;
		return l_Result;
		} else {
			return false;
		}
	}
	catch (ll_error) {
		if (ll_error.number == -2147217843) {
			//login failed... Ok, our server is online...
			l_Result = true;
		} else {
			l_Result = false;
		}
		return l_Result;
	}
}
function f_GetLastOSBootDateTime() {
	try {
		var objWin32OS = GetObject("winmgmts://./root/cimv2").ExecQuery("Select LastBootUpTime from win32_OperatingSystem", "WQL", wbemFlagReturnImmediately | wbemFlagForwardOnly);
		var lWin32OS = new Enumerator(objWin32OS);
		var lLastBootUpDate = "";
		for (; !lWin32OS.atEnd(); lWin32OS.moveNext()) {
			lLastBootUpDate = lWin32OS.item().LastbootUpTime;
		}
		return lLastBootUpDate;
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "f_GetLastOSBootDateTime")
		return 0;
	}
}
function f_GetLastLogonDateTime(In_UserName, In_UserDomain) {
	try {
		var objWin32OS = GetObject("winmgmts://./root/cimv2").ExecQuery("Select * from Win32_LogonSession Where LogonType = 2", "WQL", wbemFlagReturnImmediately | wbemFlagForwardOnly);
		var lWin32OS = new Enumerator(objWin32OS);
		var lLastBootUpDate = "20000101000000.000000+000";
		var l_UserName = f_validate_str_param (In_UserName, trim(GetUserName().toLowerCase()));
		var l_UserDomain = f_validate_str_param (In_UserDomain, trim(GetDomainName().toLowerCase()));
		for (; !lWin32OS.atEnd(); lWin32OS.moveNext()) {
			lLastBootUpDate = lWin32OS.item().StartTime;
			var objLoggedOnUser = GetObject("winmgmts://./root/cimv2").ExecQuery("Associators of {Win32_LogonSession.LogonId=" + lWin32OS.item().LogonId + "} Where AssocClass=Win32_LoggedOnUser Role=Dependent", "WQL", wbemFlagReturnImmediately | wbemFlagForwardOnly);
			var lLoggedUser = new Enumerator(objLoggedOnUser);
			for (; !lLoggedUser.atEnd(); lLoggedUser.moveNext()) {
				if ((trim(lLoggedUser.item().Name.toString().toLowerCase()) == trim(l_UserName.toString().toLowerCase())) && (trim(lLoggedUser.item().Domain.toString().toLowerCase()) == trim(l_UserDomain.toString().toLowerCase()))) {
					//WScript.Echo (lLoggedUser.item().Name + "@" + lLoggedUser.item().Domain + "; " + lLastBootUpDate);
					break;
				}
			}
			var objLoggedOnUser = null;
			var lLoggedUser = null;
		}
		var objWin32OS = null;
		var lWin32OS = null;
		return lLastBootUpDate;
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "f_GetLastLogonDateTime")
		return 0;
	}
}
function f_ParseWMIDateToUlian(In_WMIDate) {
	try {
		if (In_WMIDate.toString().length > 14) {
			var ll_Year = parseInt(In_WMIDate.toString().substring(0, 4), 10);
			var ll_Month = parseInt(In_WMIDate.toString().substring(4, 6), 10);
			var ll_Day = parseInt(In_WMIDate.toString().substring(6, 8), 10);
			var ll_Hour = parseInt(In_WMIDate.toString().substring(8, 10), 10);
			var ll_Min = parseInt(In_WMIDate.toString().substring(10, 12), 10);
			var ll_Sec = parseInt(In_WMIDate.toString().substring(12, 14), 10);
			return f_calendar_to_julian (ll_Year, ll_Month, ll_Day) + f_ConvertHMSToJulianDate (ll_Hour, ll_Min, ll_Sec);
		}
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "f_ParseWMIDateToUlian")
		return 0;
	}
}
function f_GetEnvironment (in_Env_Name) {
	try {
		var lWshShell = WScript.CreateObject("WScript.Shell");
		var lWshEnv = lWshShell.Environment("Process");
		return lWshEnv(in_Env_Name.toString().replace(/%/g,""));
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "f_GetEnvironment")
		return 0;
	}
}
function getXmlHttp(){
	try {
		return new ActiveXObject("Msxml2.XMLHTTP");
	} catch (e) {
		try {
			return new ActiveXObject("Microsoft.XMLHTTP");
		} catch (ee) {
		}
	}
	if (typeof XMLHttpRequest != 'undefined') {
		return new XMLHttpRequest();
	}
}
function getUrl(url, cb) {
	var xmlhttp = getXmlHttp();
	xmlhttp.open("GET", url+'?r='+Math.random(), true);
	xmlhttp.onreadystatechange = function() {
		if (xmlhttp.readyState == 4) {
			cb(
			xmlhttp.status, 
			xmlhttp.getAllResponseHeaders(), 
			xmlhttp.responseText
			);
		}
	}
	try {
		xmlhttp.send(null);
	} catch (e) {
		try {
			WScript.sleep(2000);
			xmlhttp.send(null);
		} catch (ee) {
		}
	}
}
function getUrlSync(url, NotUseIECache) {
	var xmlhttp = getXmlHttp();
	var l_UseIECache = NotUseIECache;
	xmlhttp.open("GET", url, false);
	if (!l_UseIECache) {
		xmlhttp.setRequestHeader ("If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT");
	}
	try {
		xmlhttp.send(null);
	} catch (e) {
		try {
			WScript.sleep(2000);
			xmlhttp.send(null);
		} catch (ee) {
		}
	}
	if (xmlhttp.status == 200) {
		return xmlhttp.responseText;
	}
}
function getUrlSizeSync(url, NotUseIECache) {
	var xmlhttp = getXmlHttp();
	var l_UseIECache = NotUseIECache;
	xmlhttp.open("HEAD", url, false);
	if (!l_UseIECache) {
		xmlhttp.setRequestHeader ("If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT");
	}
	try {
		xmlhttp.send(null);
	} catch (e) {
		try {
			WScript.sleep(2000);
			xmlhttp.send(null);
		} catch (ee) {
		}
	}
	if (xmlhttp.status == 200) {
		return parseFloat(xmlhttp.getResponseHeader("Content-Length"));
	}
}
function f_OverWriteFileSimple(In_FileName, In_Data) {
	var fso;
	var ll_FileName;
	var file_handle;
	try
	{
		fso = new ActiveXObject("Scripting.FileSystemObject");
		if (f_IsFileExist(In_FileName)) {
			//'File already exist... So we must open it for writing...
			if (In_FileName.toString().indexOf("~") >= 0) {
				//'Too bad case - short name...
				var aFileObj = fso.GetFile(In_FileName);
				ll_FileName = aFileObj.ShortPath;
				aFileObj = null;
			} else {
				ll_FileName = In_FileName;
			}
			//set file_handle = fso.OpenTextFile(In_LogName, OpenFileForWriting, True, TristateFalse)
			file_handle = fso.OpenTextFile(ll_FileName, OpenFileForWriting, true, TristateFalse);
		} else {
			//'File not exist... So we must open it for writing...
			file_handle = fso.CreateTextFile(In_FileName, true);
		}
		file_handle.write (In_Data.toString());
		file_handle.close();
		file_handle = null;
		fso = null;
		return In_Data.toString().length;
	}
	catch (ll_Error) {
		errorHandler(ll_Error, true, "f_OverWriteFileSimple")
		return 0;
	}
}
function f_CreateDirectory (In_FolderName) {
	var l_FolderName = f_validate_str_param (In_FolderName, "");
	if (trim(l_FolderName) == "") {
		return 1;	//Empty folder name...
	}
	if (l_FolderName.toString().substring(l_FolderName.toString().length-1, l_FolderName.toString().length) != "\\") {
		l_FolderName = l_FolderName + "\\";
	}
	//Ok at this point I must split full folder path
	var l_FilePathBuffer = l_FolderName.toString().replace(/\\\\/g, "??UNC?");
	var objFSO, objFolder;
	l_PathArray = l_FilePathBuffer.toString().split("\\");
	l_FolderBuffer = l_PathArray[0];
	for (var i_local = 1; i_local < l_PathArray.length; ++i_local) {
		if (trim(l_PathArray[i_local]) != "") {
			l_FolderBuffer = l_FolderBuffer + "\\" + l_PathArray[i_local];
			if (! f_IsFolderExist(l_FolderBuffer.toString().replace(/\?\?UNC\?/g,"\\\\"))) {
				objFSO = new ActiveXObject("Scripting.FileSystemObject");
				objFolder = objFSO.CreateFolder(l_FolderBuffer.toString().replace("??UNC?","\\"));
				objFSO = null;
				objFolder = null;
			} else {
				if (clng(i_local) == (l_PathArray.length - 2)) {
					return 2;	//Folder already exists...
				}
			}
		}
	}
	return 0;	//All ok...
}
function f_DeleteFile(In_File_Name) {
	var l_FSO, lFile;
	l_FSO = new ActiveXObject("Scripting.FileSystemObject");
	if (f_IsFileExist(In_File_Name)) {
		//' Get a handle to the file.
  		var lFile = l_FSO.GetFile(In_File_Name)
	  	//' Delete file
  		lFile.Delete (true) //'Lets force file deletion
	}
	l_FSO = null;
}
function f_DeleteFolder(In_Folder_Name) {
	var l_FSO, lFile, l_FolderName
	l_FSO = new ActiveXObject("Scripting.FileSystemObject");
	if (f_IsFolderExist(In_Folder_Name)) {
		if (In_Folder_Name.toString().substring(In_Folder_Name.toString().length-1, In_Folder_Name.toString().length) == "\\") {
			l_FSO.DeleteFolder (In_Folder_Name.toString().substring(0, In_Folder_Name.toString().length-1, true));
		} else {
			l_FSO.DeleteFolder (In_Folder_Name, true);
		}
	}
}
function f_ReadAllTextFile(In_TextFile) {
	try {
		var file_handle;
		var fso = new ActiveXObject("Scripting.FileSystemObject");
		if (f_IsFileExist(In_TextFile)) {
			//'File already exist... So we must open it for reading, and read all text
			file_handle = fso.OpenTextFile(In_TextFile, OpenFileForReading);
		} else {
			return "";
		}
		if (file_handle.AtEndOfStream) {
	       		return "";
		} else {
        		return file_handle.ReadAll();
		}
		file_handle.close;
		file_handle = null;
		fso = null;
	}
	catch (ll_Error) {
		errorHandler(ll_Error, false, "f_ReadAllTextFile")
		return "";
	}
}
function testfunctions(){
	var iError;
	//f_GetWin32Processes(In_PCName, In_ShowPID, In_ShowProcessName, In_ShowCommandLine, In_GetProcOwner, In_GetProcOwnerSID, In_ProcElemDelim, In_RowDelim)
//	WScript.Echo (f_GetWin32Processes("", true, true, true, false, false, "|||", "\n")); //- Ok
//	WScript.Echo (f_IsFileExist("c:\\windows"));
//	WScript.Echo (f_IsFolderExist("c:\\windows"));
//	WScript.Echo (f_Read_Registry_Ex(".", HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", "DigitalProductId", "REG_BINARY", iError));
//	WScript.Echo (f_Read_Registry_Ex(".", HKEY_LOCAL_MACHINE, "SYSTEM\\CurrentControlSet\\Control\\Session Manager", "ObjectDirectories", "REG_MULTI_SZ", iError));
//	WScript.Echo (f_Read_Registry_Ex(".", HKEY_LOCAL_MACHINE, "SYSTEM\\CurrentControlSet\\Control\\Session Manager", "LicensedProcessors", 4, iError));
	WScript.Echo (GetScriptEngineInfo());
/*
	var registry = new Registry();
	var lTestArray = registry.read (HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", "DigitalProductId");
	WScript.Echo ("lTestArray.length = " + lTestArray.length);
	WScript.Echo (f_string(50, "+"));
	WScript.Echo ("!" + trim("  123  543  ") + "!");
//	WScript.Echo ("now = " + now()+ "; " + year(now()) + "; " + month(now()) + "; " + day(now()) );
	WScript.Echo ("f_calendar_to_julian (year(now()), month(now()), day(now())) = " + f_calendar_to_julian (year(now()), month(now()), day(now())));
	WScript.Echo ("f_ConvertHMSToJulianDate (hour(now()), minute(now()), second(now()) = " + f_ConvertHMSToJulianDate (hour(now()), minute(now()), second(now())));
	WScript.Echo ("f_SetHeartBeats = " + f_SetHeartBeats());
	printf ("GetComputerName() = " + GetComputerName());
	printf("f_GetPCDomainName(In_PCName) = " + f_GetPCDomainName(""));
	for (iError=0; iError < 9; ++iError) {
		printf("GetFullUserName_Ex(" + iError + ") = " + trim(GetFullUserName_Ex(iError)));
		sleep(0.1);
	}
	printf("GetUserName() = " + GetUserName());
	printf("GetDomainName() = " + GetDomainName());
	printf("GetFullUserName() = " + GetFullUserName());
	printf("f_GetUserSID (In_PCName, In_UserName) = " + f_GetUserSID ("", ""));
	printf("f_ping (172.17.0.1) = " + f_ping ("172.17.0.1"));
	for (var io=0; io<9; ++io) {
		printf("f_ping_ex (\"172.17.0.1;172.17.0.5;172.17.0.101;172.17.2.1\", " + io + ", 5, \";\") = " + f_ping_ex ("172.17.0.1;172.17.0.5;172.17.0.101;172.17.2.1", io, 5, ";"));
	}
	WScript.Echo (f_string(50, "-"));
	for (var i=0; i<19; ++i)
	{
		WScript.Echo ("Processing index: [" + i + "]: => " + f_GetOSInfo("", i));
	}
*/
	printf("f_GetUserSID (In_PCName, In_UserName) = " + f_GetUserSID ("", ""));
	WScript.Echo ("f_ConvertHMSToJulianDate (hour(now()), minute(now()), second(now()) = " + f_ConvertHMSToJulianDate (hour(now()), minute(now()), second(now())));
	WScript.Echo ("f_SetHeartBeats = " + f_SetHeartBeats());
	WScript.Echo ("f_GetFileDirectory(WScript.ScriptFullName) = [" + f_GetFileDirectory(WScript.ScriptFullName) + "]");
	var l_DB_Srv = "172.17.0.1";
	WScript.Echo ("f_TestSQLDBConnection(\""+ l_DB_Srv +"\") = " + f_TestSQLDBConnection(l_DB_Srv));
	WScript.Echo ("f_GetLastOSBootDateTime() = " + f_GetLastOSBootDateTime());
	WScript.Echo ("f_ParseWMIDateToUlian(f_GetLastOSBootDateTime()) = " + f_ParseWMIDateToUlian(f_GetLastOSBootDateTime()));
	printf("f_GetEnvironment(\"%SystemRoot%\") = " + f_GetEnvironment("%SYSTEMROOT%"));
	printf("f_GetLastLogonDateTime() = " + f_GetLastLogonDateTime());
//	printf("f_ReadAllTextFile(In_TextFile) = " + f_ReadAllTextFile(gUserSysLibLocation));
//	printf("f_CreateDirectory (\"\") = " + f_CreateDirectory ("C:\\Documents and Settings\\Administrator\\SysLib1\\ chser$"));
//	printf(f_DeleteFolder("C:\\Documents and Settings\\Administrator\\SysLib1\\ chser$\\"));
	/*
	printf("getUrlSync(\"https://sirius.space//\") = " + getUrlSync("https://sirius.space/", true));
	printf("getUrlSizeSync = " + getUrlSizeSync("https://sirius.space/syslib.vbs", true));
	printf("f_OverWriteFileSimple(In_FileName, In_Data) = " + f_OverWriteFileSimple("$$sl_mon.out$$", getUrlSync("https://sirius.space/syslib.vbs", true)));
	*/
	/*
	f_StartProcessMonitor();
	do 
	{
		WScript.sleep (200);
	} while (true);
	*/
	return;
}
function main ()
{
	f_LowLevelInit();
	f_StartupLogo();
	/////////////////////////////////////////////
	//Special test code section...		/////
	/////////////////////////////////////////////
	//testfunctions();
	/////////////////////////////////////////////
	//End of special test code section...   /////
	/////////////////////////////////////////////
	var l_RunCycle = true;
	var l_IsSystem = false;
	var l_CurrentPhase = 0;
	var l_UserSID = f_GetUserSID ("", "");
	var ll_UserName = trim(GetUserName().toLowerCase());
	var lLastOSBootupTime = parseFloat(f_ParseWMIDateToUlian(f_GetLastOSBootDateTime()));
	var lLastLogonTime = parseFloat(f_ParseWMIDateToUlian(f_GetLastLogonDateTime("", "")));
	var lLastCmpTime = 0;
	var lDBMSSrvArray = gSysLibDBMSSrv.split(";");
	var lNetworkOk = false;
	var lStartupWait = parseInt(gStartupWait, 10);
	var lScriptStartupWait = parseInt(gScriptStartupWait, 10);
	var lScriptSLStartupWait = parseInt(gScriptSLStartupWait, 10);
	var lScriptSLSysLibProcessWait = parseInt(gScriptSLSysLibProcessWait, 10);
	var llSysLibFile = "";
	var llSysLibArg = "";
	var llSysLibArgStd = "";
	var llSystemRoot = f_GetEnvironment("%SYSTEMROOT%");
	var llTempRoot =  f_GetEnvironment("%Temp%");
	var llUserProfile =  f_GetEnvironment("%UserProfile%");
	var llSystem32Root = trim(llSystemRoot) + "\\System32\\";
	var llCommandLine = "";
	var llUseAlternateLocation = false;
	var llUseAlternateMethod = false;
	var llExecCount = 0;
	var ll_PCName = trim(GetComputerName().toLowerCase());
	var ll_UserDomain = trim(GetDomainName().toLowerCase());
	if (ll_PCName == ll_UserDomain) {
		//User logged on into local account. Lets start script at this time...
		lStartupWait = 3;
		lScriptStartupWait = 3;
	}
	var llCurrentTime = f_calendar_to_julian (year(now()), month(now()), day(now())) + f_ConvertHMSToJulianDate (hour(now()), minute(now()), second(now()));
	do {
//		WScript.Echo ("Network presence: " + lNetworkOk + "; lScriptStartupWait = " + lScriptStartupWait + "; lStartupWait = " + lStartupWait);
		if (lStartupWait == 0) {
			//Network presense test section
			if (lScriptStartupWait == 0) {
				if (lDBMSSrvArray instanceof Array) {
					for (var i_local = 0; i_local < lDBMSSrvArray.length; ++ i_local) {
//						if (f_TestSQLDBConnection(lDBMSSrvArray[i_local])) {
						if (f_ping(lDBMSSrvArray[i_local])) {
							lNetworkOk = true;
							break;
						}
					}
				}
				if (lNetworkOk) {
//					WScript.Echo("Entered into network section!");
					//Network present now lets check SysLib completition.
					//Location dependent on username...
					if (ll_UserName == "system" || ll_UserName == "ñèñòåìà" || l_UserSID == "S-1-5-18" || l_UserSID == "S-1-5-19" || l_UserSID == "S-1-5-20") {
						l_IsSystem = true;
					} else {
						l_IsSystem = false;
					}
					if (l_IsSystem) {
						//Lets read last successfull SysLib startup date
						ll_LastSuccessfullOperation = parseFloat(ReadRegistry ("HKEY_LOCAL_MACHINE\\Software\\SysLib\\LastSuccessfullStartup"));
						llCurrentTime = f_calendar_to_julian (year(now()), month(now()), day(now())) + f_ConvertHMSToJulianDate (hour(now()), minute(now()), second(now()));;
						if (isNaN(ll_LastSuccessfullOperation)) {
							ll_LastSuccessfullOperation = 0;
						}
						if (isNaN(llCurrentTime)) {
							llCurrentTime = 0;
						}
						lLastCmpTime = lLastOSBootupTime;
					} else {
						//Lets read last successfull SysLib logon date
						ll_LastSuccessfullOperation = parseFloat(ReadRegistry ("HKEY_CURRENT_USER\\Software\\SysLib\\LastSuccessfullLogon"));
						llCurrentTime = f_calendar_to_julian (year(now()), month(now()), day(now())) + f_ConvertHMSToJulianDate (hour(now()), minute(now()), second(now()));;
						if (isNaN(llCurrentTime)) {
							llCurrentTime = 0;
						}
						if (isNaN(ll_LastSuccessfullOperation)) {
							ll_LastSuccessfullOperation = 0;
						}
						lLastCmpTime = lLastLogonTime;
					}
//					WScript.Echo("ll_LastSuccessfullOperation = " + ll_LastSuccessfullOperation + "; lLastCmpTime = " + lLastCmpTime);
					if ((ll_LastSuccessfullOperation >= lLastCmpTime) && (ll_LastSuccessfullOperation <= llCurrentTime)) { 
						//All ok. Syslib already complete it's operations. Lets simple exit
						//Is Complete
						printf ("ll_LastSuccessfullOperation = " + ll_LastSuccessfullOperation + "; Syslib engine already processed this session...");
						if (llUseAlternateLocation && gRemoveSysLibDownloadedFromAlternateLoc) {
							var ll_ret = f_OverWriteFileSimple(llUserProfile + "\\SysLib\\" + gSysLibSrcName, "");
							ll_ret = f_DeleteFile(llUserProfile + "\\SysLib\\" + gSysLibSrcName);
						}
						l_RunCycle = false;
					} else {
						//Smthing wrong... Lets wait for specified timeout and try to run SysLib script...
						if (l_IsSystem) {
							llSysLibFile = gSystemSysLibLocation;
							llSysLibArg = "startup-slm";
							llSysLibArgStd = "startup";
						} else {
							llSysLibFile = gUserSysLibLocation;
							llSysLibArg = "userlogon-slm";
							llSysLibArgStd = "userlogon";
						}
						if (f_FindProcessesByCmdLine(".", "syslib.vbs\" " + llSysLibArg) == 0) {
							if (f_FindProcessesByCmdLine(".", "syslib.vbs " + llSysLibArg) == 0) {
								if (f_FindProcessesByCmdLine(".", "syslib.vbs\" " + llSysLibArgStd) == 0) {
									if (f_FindProcessesByCmdLine(".", "syslib.vbs " + llSysLibArgStd) == 0) {
										//All too bad... SysLib process not running, also it may be terminated with some error until it's completition...
										printf("Unable to detect running SysLib.vbs processes...");
										if (lScriptSLSysLibProcessWait > 5) {
											lScriptSLSysLibProcessWait = 5;	//Lets give OS time to flush written data to registry...
										}
									} else {
										//printf("Detected SysLib.vbs with [" + llSysLibArgStd + "] parameter...");
										lScriptSLSysLibProcessWait = parseInt(gScriptSLSysLibProcessWait, 10);
									}
								} else {
									//printf("Detected \"SysLib.vbs\" with [" + llSysLibArgStd + "] parameter...");
									lScriptSLSysLibProcessWait = parseInt(gScriptSLSysLibProcessWait, 10);
								}
							} else {
								lScriptSLSysLibProcessWait = parseInt(gScriptSLSysLibProcessWait, 10);
								//printf("Detected SysLib.vbs with [" + llSysLibArg + "] parameter...");
							}
						} else {
							//printf("Detected \"SysLib.vbs\" with [" + llSysLibArg + "] parameter...");
							lScriptSLSysLibProcessWait = parseInt(gScriptSLSysLibProcessWait, 10);
						}
						if (lScriptSLSysLibProcessWait == 0) {
							if (f_IsFileExist(llSysLibFile)) {
								//Our script file exists.. Lets try to read it... In may be inaccessible...
								if (f_ReadAllTextFile(llSysLibFile).toString().length > 0) {
									//All ok, lets try  to execute our script engine and wait for syslib for it's completition...
									//Lets assembly required command...
									llCommandLine = llSystem32Root + gScriptEngine + " \"" + llSysLibFile + "\" " + llSysLibArg;
									printf("Ready to execute our script!");
									printf("llCommandLine = " + llCommandLine);
									if (llExecCount < gSysLibExecTryCount) {
										var WshShell = WScript.CreateObject("WScript.Shell");
										llExecCount++;
										WshShell.run(llCommandLine, 1, false);
										//var oExec = WshShell.Exec(llCommandLine);
										//var ll_ProcessID = oExec.ProcessID;
										//printf("ll_ProcessID = " + ll_ProcessID);
										//oExec = null;
										//Lets wait until our script complete...
										lScriptStartupWait = parseInt(gScriptStartupWait, 10);
										lScriptSLSysLibProcessWait = parseInt(gScriptSLSysLibProcessWait, 10);
										WshShell = null;
									} else {
										printf("Script run's too many times ("+ gSysLibExecTryCount + ")... It's possible, that some critical system error occured.\nExiting...");
										l_RunCycle = false;
									}
								} else {
									//Our file unreadable... Lets try to use alternate method...
									llUseAlternateMethod = true;
								}
							} else {
								llUseAlternateMethod = true;
							}
							if (llUseAlternateMethod) {
								llUseAlternateMethod = false;
								//Lets try to use alternative source, and if we are fail - lets wait for good moment...
								if (getUrlSizeSync(gSysLibAltrnativeLocation, true) > 0) {
									//Ok, now trying to download our script to temporary location...
									if (! f_IsFolderExist(llUserProfile + "\\SysLib")) {
										f_CreateDirectory (llUserProfile + "\\SysLib");
									}
									var ll_ret = f_OverWriteFileSimple(llUserProfile + "\\SysLib\\" + gSysLibSrcName, getUrlSync(gSysLibAltrnativeLocation, true));
									llCommandLine = llSystem32Root + gScriptEngine + " \"" + llUserProfile + "\\SysLib\\" + gSysLibSrcName + "\" " + llSysLibArg;
									llUseAlternateLocation = true;
									printf("Ready to execute our script, downloaded via alternate path!");
									printf("llCommandLine = " + llCommandLine);
									if (llExecCount < gSysLibExecTryCount) {
										var WshShell = WScript.CreateObject("WScript.Shell");
										llExecCount++;
										WshShell.run(llCommandLine, 1, false);
										lScriptStartupWait = parseInt(gScriptStartupWait, 10);
										lScriptSLSysLibProcessWait = parseInt(gScriptSLSysLibProcessWait, 10);
										WshShell = null;
									} else {
										printf("Script run's too many times ("+ gSysLibExecTryCount + ")... It's possible, that some critical system error occured.\nExiting...");
										l_RunCycle = false;
									}
								} else {
									//Lets try to wait once more...
									lStartupWait = parseInt(gStartupWait, 10);
								}
							}
						}
					}
				} else {
					//Lets reset startup counter in case of network failure...
					lStartupWait = parseInt(gStartupWait, 10);
				}
			}
		}
		//Sleep section
		WScript.sleep (1000);
		if (lStartupWait > 0) {
			lStartupWait = lStartupWait - 1;
		}
		if ((lScriptStartupWait > 0) && (lStartupWait == 0)) {
			lScriptStartupWait = lScriptStartupWait - 1;
		}
		if ((lScriptSLSysLibProcessWait > 0) && (lScriptStartupWait == 0) && (lStartupWait == 0)) {
			lScriptSLSysLibProcessWait = lScriptSLSysLibProcessWait - 1;
		}
	} while (l_RunCycle);
	return 0;
}
main ();

