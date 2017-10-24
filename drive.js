var wbemFlagReturnImmediately = 0x10; 
var wbemFlagForwardOnly = 0x20; 

var arrComputers = new Array("."); 
for (i = 0; i < arrComputers.length; i++) { 
WScript.Echo(); 
WScript.Echo("=========================================="); 
WScript.Echo("Computer: " + arrComputers[i]); 
WScript.Echo("=========================================="); 

var objWMIService = GetObject("winmgmts:\\\\" + arrComputers[i] + "\\root\\CIMV2"); 
var colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PhysicalMedia", "WQL", 
wbemFlagReturnImmediately | wbemFlagForwardOnly); 

var enumItems = new Enumerator(colItems); 
for (; !enumItems.atEnd(); enumItems.moveNext()) { 
var objItem = enumItems.item(); 

WScript.Echo("Capacity: " + objItem.Capacity); 
WScript.Echo("Caption: " + objItem.Caption); 
WScript.Echo("CleanerMedia: " + objItem.CleanerMedia); 
WScript.Echo("CreationClassName: " + objItem.CreationClassName); 
WScript.Echo("Description: " + objItem.Description); 
WScript.Echo("HotSwappable: " + objItem.HotSwappable); 
WScript.Echo("InstallDate: " + WMIDateStringToDate(objItem.InstallDate)); 
WScript.Echo("Manufacturer: " + objItem.Manufacturer); 
WScript.Echo("MediaDescription: " + objItem.MediaDescription); 
WScript.Echo("MediaType: " + objItem.MediaType); 
WScript.Echo("Model: " + objItem.Model); 
WScript.Echo("Name: " + objItem.Name); 
WScript.Echo("OtherIdentifyingInfo: " + objItem.OtherIdentifyingInfo); 
WScript.Echo("PartNumber: " + objItem.PartNumber); 
WScript.Echo("PoweredOn: " + objItem.PoweredOn); 
WScript.Echo("Removable: " + objItem.Removable); 
WScript.Echo("Replaceable: " + objItem.Replaceable); 
WScript.Echo("SerialNumber: " + objItem.SerialNumber); 
WScript.Echo("SKU: " + objItem.SKU); 
WScript.Echo("Status: " + objItem.Status); 
WScript.Echo("Tag: " + objItem.Tag); 
WScript.Echo("Version: " + objItem.Version); 
WScript.Echo("WriteProtectOn: " + objItem.WriteProtectOn); 
} 
} 

function WMIDateStringToDate(dtmDate) 
{ 
if (dtmDate == null) 
{ 
return "null date"; 
} 
var strDateTime; 
if (dtmDate.substr(4, 1) == 0) 
{ 
strDateTime = dtmDate.substr(5, 1) + "/"; 
} 
else 
{ 
strDateTime = dtmDate.substr(4, 2) + "/"; 
} 
if (dtmDate.substr(6, 1) == 0) 
{ 
strDateTime = strDateTime + dtmDate.substr(7, 1) + "/"; 
} 
else 
{ 
strDateTime = strDateTime + dtmDate.substr(6, 2) + "/"; 
} 
strDateTime = strDateTime + dtmDate.substr(0, 4) + " " + 
dtmDate.substr(8, 2) + ":" + 
dtmDate.substr(10, 2) + ":" + 
dtmDate.substr(12, 2); 
return(strDateTime); 
} 