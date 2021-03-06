/* Globals */
var objDrives = {};
var isAdmin = Elevated();
var isWinPE;

/**********************************************************************************************************************
*						        Document ready functions
**********************************************************************************************************************/
$(document).ready(function () {
    loadPage("landing");

    //Initialize replacement scrollbars
    $(".pages").mCustomScrollbar({
        theme: "minimal",
        live: "once"
    });
    $("#general-output-scroll").mCustomScrollbar({
        theme: "minimal",
        live: "once"
    });

    //Initialize dialog box for warning and error messages
    $("#dialog").dialog({
        autoOpen: false
    });

    $('#general-output h3').click(function () {
        $(this).next().toggle('fast');
        $(this).children().toggleClass("ui-icon-triangle-1-e ui-icon-triangle-1-s");
        return false;
    }).next().hide();
    $('#general-output .ui-widget-content').show('fast');
    $('#general-output .ui-widget-content').prev().children().removeClass("ui-icon-triangle-1-e ui-icon-triangle-1-s").addClass("ui-icon-triangle-1-s");

    isWinPE = "true" //$("#input-isWinPE").val();
    if (isWinPE == "true") {
        //Booting from flash drive. Don't show Loadstate, since it can only run from from within Windows
        $('.hideInPE').css('display', 'none');
    }
    else{
        //Booting from Windows, don't show OSD, FnF, DISM, since those need to run from WinPE
        $('.hideInWindows').css('display', 'none');
    }
});

$(".on-load").delay(200).animate({ "opacity": "1" }, 700);

/**********************************************************************************************************************
*						        Navigation Functions
**********************************************************************************************************************/
$('#button-ShowAll').on('click', function () {
    $('.accordion .ui-widget-content').show('fast');
    $('.accordion .ui-widget-content').prev().children().removeClass("ui-icon-triangle-1-e ui-icon-triangle-1-s").addClass("ui-icon-triangle-1-s");
    
});

$('#button-HideAll').on('click', function () {
    $('.accordion .ui-widget-content').hide('fast');
    $('.accordion .ui-widget-content').prev().children().removeClass("ui-icon-triangle-1-e ui-icon-triangle-1-s").addClass("ui-icon-triangle-1-e");
});

$('.button-nav, #button-landing').on('click', function (event) {
    var targetID = $(this).attr("id");      //Get the id of the button that was clicked
    var targetOperation = targetID.replace("button-", "");  //Strip away button-, leaving us with osd, flushfill, dism, usmt, software, or tools
    loadPage(targetOperation, event);
});

$('.button-toolbar').hover(
    function () {
        $(this).children().addClass("button-background-hover");
        $(this).children().removeClass("button-screen");
    }, function () {
        $(this).children().removeClass("button-background-hover");
        $(this).children().addClass("button-screen");
    }
);

$('#button-exit').hover(
    function () {
        $("#portal-exit").animate({
            top: "20px",
            queue: false
        }, 300, "easeInOutElastic");
    }, function () {
        $("#portal-exit").animate({
            top: "3px",
            queue: false
        }, 300, "easeInBounce", function(){$(this).clearQueue();});
    }
);

$("#screen-test").hover(
    function () {
        $("#screen-test").children().css("background-image", "");
    }, function () {
        $("#screen-test").children().css("background-image", "url(images/screen-only-45.svg)");
    }
);

// $('#button-ScanState').on('click', function () {
//     $('#page-usmt .scanState').css('display', 'list-item');
//     $('#page-usmt #run-usmt').css('display', 'inline-block');
//     $('#page-usmt #run-LoadState').css('display', 'none');
// });

// $('#button-LoadState').on('click', function () {
//     $('#page-usmt .scanState').css('display', 'none');
//     $('#page-usmt #run-LoadState').css('display', 'inline-block');
//     $('#page-usmt #run-usmt').css('display', 'none');
// });

function loadPage(targetOperation, event) {
    //Rerun listDrives function if returning to landing page
    if (targetOperation == "landing") {
        //Find and list drives attached to this machine
        //If no windows directory was found and an encrypted drive WAS found, recolor list items that require a windows directory
        WMIListDrives();
        dimElements();
        event = null;
        try {
            //Check for a TPM chip
            TPMCheck();
        }
        catch (err) {
            $('#page-landing').append(err.message);
        }
    }

    //Prevent recreating the page if we are already on it. i.e. clicking dism button when we are already on dism page.
    var isSamePage = $("#ul-" + targetOperation).children().length;
    if (isSamePage == 0) {
        //If navigating to a new page, clear out the list elements so we can reuse list items on the new page.
        $(".ul-elements").empty();
    }
    else return;

    //Grab the appropriate templates for the page we are navigating to and join them together in a variable named "template"
    var templates = $("." + targetOperation);
    if (templates.length > 0) {
        var template = "";
        
        templates.each(function () {
            if (isWinPE != "true" && $(this).hasClass("hideInWindows")) {
                //If we are in Windows, don't show template items that should be hidden (ie Scanstate)
                return;
            }
            if (($(this).attr("id") == "temp-bitlocker-unlock" && objDrives["Encrypted Drive Found"] == false)){
                //Only show the Bitlocker Unlock template if an encrypted drive is found and locked
                return;
            }

            //Append the current template to the template list
            template += $(this).html();
        });

        //Insert the joined templates into the page we are nevigating to
        $("#ul-" + targetOperation).append(template);
    }

    initAccordion();
    $('.pages').removeClass("active-page");
    $('.header-text').removeClass("active-header");
    $('.button-nav').removeClass("active");
    try { $(event.currentTarget).addClass("active"); }
    catch (e) { }
    $('#page-' + targetOperation).addClass("active-page");
    $('#header-' + targetOperation).addClass("active-header");
    //Reset Scanstate button, since the template for USMT always starts on Scanstate
    // $('#page-usmt #run-usmt').css('display', 'inline-block');
    // $('#page-usmt #run-LoadState').css('display', 'none');

    if(targetOperation == "osd")
    {
        $('#input-osd-checkbox').prop("checked", true);
        $('#MBAM').prop("checked", true);
        $('#page-osd .loadState').css('display', 'none');
    }

    dimElements();
}

/**********************************************************************************************************************
*						        Misc Functions
**********************************************************************************************************************/
function getInputValues() {
    var inputValues = {};
    $(".accordion .input-objects").each(function () {
        //If this input is a checkbox, get the value of prop("checked") instead of val()
        if ($(this).attr("type") == "checkbox") {
            inputValues[$(this).attr("name")] = $(this).prop("checked");
        }
        //If this input is a text input field, get val()
        else {
            inputValues[$(this).attr("name")] = $(this).val();
        }
    });
    return inputValues;
}

function osdCheckboxChange(e) {
    if (e.checked == true) {
        $('#MBAM').prop("checked", true);
    }
    else {
        $('#MBAM').prop("checked", false);
    }
}

function loadStateCheckboxChange(e) {
    if (e.checked == true) {
        $('#page-osd .loadState').css('display', 'list-item');
    }
    else {
        $('#page-osd .loadState').css('display', 'none');
    }
}

function initAccordion() {
    //Initialize accordion for displaying list itmes
    $('.accordion h3').click(function () {
        $(this).next().toggle('fast');
        $(this).children().toggleClass("ui-icon-triangle-1-e ui-icon-triangle-1-s");
        return false;
    }).next().hide();
    $('.accordion .ui-widget-content').not($(".li-win-not-found .ui-widget-content")).show('fast');
    $('.accordion .ui-widget-content').not($(".li-win-not-found .ui-widget-content")).prev().children().removeClass("ui-icon-triangle-1-e ui-icon-triangle-1-s").addClass("ui-icon-triangle-1-s");
}

function WMIListDrives() {
    //Get Drive Letter, Label, Capacity, and Encryption status for each drive that has a Letter mapped to it
    //drives should contain: Drive Letter, Protection Status, Encryption Method, Lock Status, Key Type, Key ID, Label, Capacity 
    writeToLog("***** Begin Sub listDrives *****");
    objDrives = { "Physical Drives": {} };
    var outputDiv = $("#bl-info-output");
    // var arConversionStatus = ["Fully Decrypted", "Fully Encrypted", "Encryption In Progress", "Decryption In Progress", "Encryption Paused", "Decryption Paused"];
    
    var loc = new ActiveXObject("WbemScripting.SWbemLocator");
    writeToLog("Executing command: ConnectServer(\".\", \"root\\cimv2\\Security\\MicrosoftVolumeEncryption\")");
    var svc = loc.ConnectServer(".", "root\\cimv2");
    writeToLog("Executing command: ExecQuery(\"SELECT Caption, DeviceID FROM Win32_DiskDrive\")");
    var wmiDiskDrives = svc.ExecQuery("SELECT Caption, DeviceID FROM Win32_DiskDrive");
    var enumDrives = new Enumerator(wmiDiskDrives);

    //Get Physical Drives
    for (; !enumDrives.atEnd(); enumDrives.moveNext()) {
        var objDrive = enumDrives.item();
        var driveID = objDrive.DeviceID;
        query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + driveID + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
        writeToLog("Executing command: ExecQuery(\"" + query + "\")");
        var wmiPartitions = svc.ExecQuery(query);
        var enumPartitions = new Enumerator(wmiPartitions);
        objDrives["Physical Drives"][driveID] = {};
        writeToLog("objDrives[\"Physical Drives\"]["+driveID+"][\"Model\"] = " + objDrive.Caption);
        objDrives["Physical Drives"][driveID]["Model"] = objDrive.Caption;

        writeToLog("Getting partitions for drive " + objDrive.Caption);
        //Get Partitions for each Physical Drive
        for (; !enumPartitions.atEnd(); enumPartitions.moveNext()) {
            var objPartition = enumPartitions.item();
            var partitionID = objPartition.DeviceID;
            writeToLog("partitionID = " + partitionID);
            writeToLog("Executing command: ExecQuery(\"ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + partitionID + "'} WHERE AssocClass = Win32_LogicalDiskToPartition\")");
            var wmiLogicalDisks = svc.ExecQuery("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + partitionID + "'} WHERE AssocClass = Win32_LogicalDiskToPartition");
            var enumLogicalDisks = new Enumerator(wmiLogicalDisks);

            writeToLog("Getting drive letter for partition " + partitionID);
            //Get Drive Letter for each Partition
            for (; !enumLogicalDisks.atEnd(); enumLogicalDisks.moveNext()) {
                try{
                    var objLogicalDisk = enumLogicalDisks.item();
                    var DriveLetter = objLogicalDisk.DeviceID.replace(":", "");
                    writeToLog("DriveLetter = " + DriveLetter);
                    //If a Physical Disk has more than one Volume, only create the Volumes object the first time we loop through the volumes
                    //Otherwise, the second volume will overwrite the first
                    if (!(objDrives["Physical Drives"][driveID].hasOwnProperty("Volumes"))){
                        objDrives["Physical Drives"][driveID]["Volumes"] = {};
                    }
                
                    objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter] = {"DriveLetter" : DriveLetter};
                    objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["Partition"] = partitionID;
                    writeToLog("objDrives[\"Physical Drives\"]["+driveID+"] = " + JSON.stringify(objDrives["Physical Drives"][driveID], null, 1));
                    WMIEncryptableVolumes(driveID, DriveLetter);
                    WMIVolumes(driveID, DriveLetter);
                }
                catch(err){
                    writeToLog("Error in WMIListDrives\n" + err.message);
                }
            }
        }
        //$("#tpm-check-output").append("<pre>" + JSON.stringify(objDrives["Physical Drives"], null, 3) + "</pre><br><br>");
    }
    
    writeToLog("Finished gathering drive info\n");
    writeToLog("objDrives[\"Physical Drives\"] = " + JSON.stringify(objDrives["Physical Drives"], null, 1));

    /*  ********* Update Windows Drive Found and Encrypted Drive Found status **********/
    objDrives["Encrypted Drive Found"] = false;
    objDrives["Windows Drive Found"] = false;
    objDrives["Encrypted Drives"] = [];
    objDrives["Windows Drives"] = [];

    /*  ********* Sort the objDrives object by physical drive ID **********/
    writeToLog("Prepare Physical Drive objects for sorting drives by Drive Letter");
    var pKeys = [], pKey, pLength, drivesSorted = [];
    try {
        //Loop through each Physical Drive
        for (var k in objDrives["Physical Drives"]) {
            writeToLog("Physical Drive: " + k);
            var vKeys = [];
            //Loop through each Volume in the current Physical Drive
            for (var j in objDrives["Physical Drives"][k]["Volumes"]) {
                //Push the drive letter of each Volume in this Physical Drive into an array
                writeToLog("Pushing key " + j + " onto vKeys array");
                vKeys.push(j);
            }
            //Sort the Volumes array so the Drive Letters are in alphabetical order
            writeToLog("Sorting Drive Letters for Physical Drive: " + k);
            vKeys.sort();
            //Set the Sort property of the current Physical Drive to the lowest Drive Letter found in its Volumes array
            writeToLog("Adding property \"Sort\": " + vKeys[0] + " to JSON object Physical Drive: " + k);
            objDrives["Physical Drives"][k]["Sort"] = vKeys[0];
            writeToLog("Adding property \"DriveID\": " + k + " to JSON object Physical Drive: " + k);
            objDrives["Physical Drives"][k]["DriveID"] = k;
            //Push the current Physical Drive onto an array to be sorted later by Drive Letter
            writeToLog("Push the current Physical Drive object onto array drivesSorted");
            drivesSorted.push(objDrives["Physical Drives"][k]);
        }

        writeToLog("Sort drivesSorted array by Physical Drives with lowest drive letter");
        //Sort Physical Drives according to their lowest Drive Letters
        drivesSorted.sort(function (a, b) {
            //Sort Drive Letters under objDrives["Physical Drives"][]["Volumes"] first
            var driveLetter1 = a["Volumes"][Object.keys(a["Volumes"])[0]]["DriveLetter"].toUpperCase();
            var driveLetter2 = b["Volumes"][Object.keys(b["Volumes"])[0]]["DriveLetter"].toUpperCase();
            if (driveLetter1 < driveLetter2) {
                return -1;
            }
            if (driveLetter1 > driveLetter2) {
                return 1;
            }
            return 0;
        }).sort(function (a, b) {
            //Sort Physical Drives by their lowest Drive Letter now
            var driveLetter1 = a["Sort"].toUpperCase();
            var driveLetter2 = b["Sort"].toUpperCase();
            if (driveLetter1 < driveLetter2) {
                return -1;
            }
            if (driveLetter1 > driveLetter2) {
                return 1;
            }
            return 0;
        });
        writeToLog("Finished sorting drives");
        writeToLog("drivesSorted = " + JSON.stringify(drivesSorted, null, "\t"));
    }
    catch(e){
        writeToLog("Error sorting drives");
        writeToLog("drivesSorted = " + JSON.stringify(drivesSorted, null, "\t"));
    }

    outputDiv.empty();

        /*******************************************************************************
        *********** Output drive info to #bl-info-output div in sorted order ***********
        *******************************************************************************/
    try{
        //Loop through each physical drive and enumerate volume info
        for (var i = 0; i < drivesSorted.length; i++)
        {
            //Store the Physical Drive ID
            var pDrive = drivesSorted[i];

            //Physical Drive ID
            outputDiv.append("<span class='drivePhysicalSpan'>Physical Drive ID:</span> " + pDrive["DriveID"]);
            //Drive Model
            outputDiv.append("<br><span class='drives'>&nbsp&nbspDrive Model:</span> " + pDrive["Model"]);

            //Loop through each volume of the current physical drive to output the volume information
            for (var j in pDrive["Volumes"]) {
                //Volumes
                drivesSorted[i]["Volumes"][j]["Partition"] ? outputDiv.append("<br><span class='drives'>&nbsp&nbspPartition:</span> " + drivesSorted[i]["Volumes"][j]["Partition"]) : "";
                //Drive Letter
                outputDiv.append("<br><span class='driveLetterSpan drives'>&nbsp&nbspDrive Letter:</span> <span class='driveLetterValue'>" + j + "</span>");
                //Drive Type
                drivesSorted[i]["Volumes"][j]["Drive Type"] ? outputDiv.append("<br><span class='drives'>&nbsp&nbsp&nbsp&nbspDrive Type:</span> " + drivesSorted[i]["Volumes"][j]["Drive Type"]) : "";
                //Lock Status
                drivesSorted[i]["Volumes"][j]["Lock Status"] == "Locked" ? outputDiv.append(" <span class='driveLocked'>(" + drivesSorted[i]["Volumes"][j]["Lock Status"] + ")</span> ") : "";
                //Volume Label
                drivesSorted[i]["Volumes"][j]["Label"] ? outputDiv.append(" | Label: " + drivesSorted[i]["Volumes"][j]["Label"]) : "";
                //Capacity
                drivesSorted[i]["Volumes"][j]["Capacity"] ? outputDiv.append("<br><span class='drives'>&nbsp&nbsp&nbsp&nbspCapacity:</span> " + drivesSorted[i]["Volumes"][j]["Capacity"]) : "";
                //Free Space
                drivesSorted[i]["Volumes"][j]["Free Space"] ? outputDiv.append(" | Free Space: " + drivesSorted[i]["Volumes"][j]["Free Space"]) : "";
                //Encryption Method
                drivesSorted[i]["Volumes"][j]["Encryption Method"] ? outputDiv.append("<br><span class='drives'>&nbsp&nbsp&nbsp&nbspEncryption:</span> " + drivesSorted[i]["Volumes"][j]["Encryption Method"]) : "";
                //Recovery Key ID
                drivesSorted[i]["Volumes"][j]["Numerical Password"] ? outputDiv.append("<br><span class='drives'>&nbsp&nbsp&nbsp&nbspRecovery Key ID:</span> " + drivesSorted[i]["Volumes"][j]["Numerical Password"]) : "";
                drivesSorted[j] = drivesSorted[i]["Volumes"][j];
            }
            outputDiv.append("<br><br>");

            //Check for encrypted drives and Windows partitions
            for (var drive in objDrives["Physical Drives"][pDrive["DriveID"]]["Volumes"]) {
                if (objDrives["Physical Drives"][pDrive["DriveID"]]["Volumes"][drive]["isEncrypted"] == true) {
                    objDrives["Encrypted Drives"].push(drive);
                    objDrives["Encrypted Drive Found"] = true;
                }
                if (objDrives["Physical Drives"][pDrive["DriveID"]]["Volumes"][drive]["isWindowsFound"] == true) {
                    objDrives["Windows Drives"].push(drive);
                    objDrives["Windows Drive Found"] = true;
                }
            }
        }
    }
    catch (err) {
       outputDiv.append(err.message);
    }

    if (objDrives["Encrypted Drive Found"] == true && objDrives["Windows Drive Found"] == false) {
        objDrives["Windows Drive Probably Locked"] = true;
    }

    return drivesSorted;
}

function WMIEncryptableVolumes(driveID, DriveLetter){
    /*  ********* Get Encryption information **********/
    if (isAdmin == false){return;}
    writeToLog(">>>> Executing function WMIEncryptableVolumes(" + driveID + ", " + DriveLetter + ") >>>>");
    // objDrives["Physical Drives"][driveID]["Volumes"] = {};
    var arEncryptionMethod = [null, "AES 128 With Diffuser", "AES 256 With Diffuser", "AES 128", "AES 256", "Hardware Encryption", "XTS AES 128", "XTS AES 256", "Unknown"];
    var arProtectionStatus = ["Protection Off", "Protection On", "Protection Unknown"];
    var arLockStatus = ["Unlocked", "Locked"];
    var arKeyType = ["Unknown or other protector type", "Trusted Platform Module (TPM)", "External key", "Numerical password", "TPM And PIN", "TPM And Startup Key", "TPM And PIN And Startup Key", "Public Key", "Passphrase", "TPM Certificate", "CryptoAPI Next Generation (CNG) Protector"];
    try {
        var loc = new ActiveXObject("WbemScripting.SWbemLocator");
        writeToLog("Executing command: ConnectServer(\".\", \"root\\cimv2\\Security\\MicrosoftVolumeEncryption\")");
        var svc = loc.ConnectServer(".", "root\\cimv2\\Security\\MicrosoftVolumeEncryption");
        writeToLog("Executing command: ExecQuery(\"SELECT * FROM Win32_EncryptableVolume WHERE DriveLetter='" + DriveLetter + ":'\")");
        var colItems = svc.ExecQuery("SELECT * FROM Win32_EncryptableVolume WHERE DriveLetter='" + DriveLetter + ":'");
        var enumItems = new Enumerator(colItems);

        var objItem = enumItems.item();
        var VolumeKeyID;
        writeToLog("WMIEncryptableVolumes - Drive Letter: " + DriveLetter);
        writeToLog("Setting Encryption Method for objDrives[\"Physical Drives\"][" + driveID + "][\"Volumes\"][" + DriveLetter + "]");
        objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["Encryption Method"] = arEncryptionMethod[objItem.ExecMethod_(objItem.Methods_.Item("GetEncryptionMethod").Name).EncryptionMethod];
        writeToLog("Setting Protection Status for objDrives[\"Physical Drives\"][" + driveID + "][\"Volumes\"][" + DriveLetter + "]");
        objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["Protection Status"] = arProtectionStatus[objItem.ExecMethod_(objItem.Methods_.Item("GetProtectionStatus").Name).ProtectionStatus];
        writeToLog("Setting Lock Status for objDrives[\"Physical Drives\"][" + driveID + "][\"Volumes\"][" + DriveLetter + "]");
        objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["Lock Status"] = arLockStatus[objItem.ExecMethod_(objItem.Methods_.Item("GetLockStatus").Name).LockStatus];
        objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["Lock Status"] == "Locked" ? objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["isEncrypted"] = true : objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["isEncrypted"] = false;

        /********** Check for Encryption Keys **********/
        writeToLog("Getting Encryption Keys for objDrives[\"Physical Drives\"][" + driveID + "][\"Volumes\"][" + DriveLetter + "]");
        var method = objItem.Methods_.Item("GetKeyProtectors");
        var inparams = method.InParameters.SpawnInstance_();

        inparams.KeyProtectorType = 3;
        VolumeKeyID = objItem.ExecMethod_(method.Name, inparams).VolumeKeyProtectorID(0);
        method = objItem.Methods_.Item("GetKeyProtectorType");
        inparams = method.InParameters.SpawnInstance_();
        inparams.VolumeKeyProtectorID = VolumeKeyID;

        var VolumeKeyProtectorType = objItem.ExecMethod_(method.Name, inparams).KeyProtectorType;
        if (VolumeKeyProtectorType != "") {
            objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["Numerical Password"] = VolumeKeyID;
        }
    }
    catch (err) {
        if(err.number == -2146823281){
            writeToLog("This drive is not encryptable - " + err.message);
        }
        else if (err.number == -2147352565){
            writeToLog("No protectors found. This drive is not encrypted - " + err.message);
        }
        else{
            writeToLog("Error in WMIEncryptableVolumes function\n Error # " + err.number + " - " + err.message);
        }
     }
    writeToLog("<<<< Exiting function WMIEncryptableVolumes(" + driveID + ", " + DriveLetter + ") <<<<\n");
}

function WMIVolumes(driveID, DriveLetter){
    /*  ********* Get Volume information **********/
    writeToLog(">>>> Executing function WMIVolumes(" + driveID + ", " + DriveLetter + ") >>>>");
    var arDriveTypes = ["Unknown", "No Root Directory", "Removable Disk", "Local Disk", "Network Drive", "Compact Disk", "RAM"];
    var loc = new ActiveXObject("WbemScripting.SWbemLocator");
    writeToLog("Executing command: ConnectServer(\".\", \"root\\cimv2\")");
    var svc = loc.ConnectServer(".", "root\\cimv2");
    writeToLog("Executing command: ExecQuery(\"SELECT * FROM Win32_Volume WHERE DriveLetter = '" + DriveLetter + ":'\")");

    colItems = svc.ExecQuery("SELECT * FROM Win32_Volume WHERE DriveLetter = '" + DriveLetter + ":'");
    enumItems = new Enumerator(colItems);
    var objItem = enumItems.item();

    writeToLog("Adding Drive " + DriveLetter + " to objDrives[\"Physical Drives\"]["+driveID+"][\"Volumes\"]");
    //If an object for this Drive Letter hasn't already been added under objDrives["Physical Drives"][driveID]["Volumes"], create one
    if (!objDrives["Physical Drives"][driveID]["Volumes"].hasOwnProperty(DriveLetter)) { objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter] = { "Drive Letter": DriveLetter }; }
    writeToLog("Adding property \"Label\": " + objItem.Label);
    objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["Label"] = objItem.Label;
    writeToLog("Adding property \"Free Space\": " + ConvertSize(objItem.Freespace));
    objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["Free Space"] = ConvertSize(objItem.Freespace);
    writeToLog("Adding property \"Capacity\": " + ConvertSize(objItem.Capacity));
    objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["Capacity"] = ConvertSize(objItem.Capacity);
    writeToLog("Adding property \"Drive Type\": " + arDriveTypes[objItem.DriveType]);
    objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["Drive Type"] = arDriveTypes[objItem.DriveType];

    //Check for a Windows path
    DriveLetter != "X" ? objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["isWindowsFound"] = ReportFolderStatus(DriveLetter + ":\\Windows") : objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["isWindowsFound"] = false;

    //Set $("#windows-drive-letter").val = true if a windows drive was found, otherwise null
    objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["isWindowsFound"] == true ? $("#windows-drive-letter").val(DriveLetter) : null;
    objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["isWindowsFound"] == true ? objDrives["Physical Drives"][driveID]["Volumes"][DriveLetter]["Label"] = "<span class='windowsSpan'>Windows</span>" : null;

    writeToLog("<<<< Exiting function WMIVolumes(" + driveID + ", " + DriveLetter + ") <<<<\n");
}

function ConvertSize(Size) {
    //Convert Bytes to KB, MB, GB, TB
    if (Size >= 1099511627776) { return Math.round(Size / 1099511627776) + " TB" }
    else if (Size >= 1073741824) { return Math.round(Size / 1073741824) + " GB" }
    else if (Size >= 1048576) { return Math.round(Size / 1048576) + " MB" }
    else if (Size >= 1024) { return Math.round(Size / 1024) + " KB" }
    else if (Size > 0) { return Size + " Bytes" }
    else { return null }
}

/*
function jsListDrives(){
     //Get Drive Letter, Label, Capacity, and Encryption status for each drive that has a Letter mapped to it
     //drives should contain: Drive Letter, Protection Status, Encryption Method, Lock Status, Key Type, Key ID, Label, Capacity 
    var drives = {}, drivesSorted = {};
    try {
        drives = listDrives();
        delete drives["setProp"];   //Remove the setProp function from the object before looping through the drive letters

        var keys = [], k, len;

        for (k in drives) {
            if (drives.hasOwnProperty(k)){
                keys.push(k);
            }
        }
        keys.sort();
        len = keys.length;

        // var landingPageDiv = $("#drive-list-output");
        var landingPageDiv = $("#bl-info-output");
        landingPageDiv.empty();
        var arKeyType = ["Unknown or other protector type", "Trusted Platform Module (TPM)", "External key", "Numerical password", "TPM And PIN", "TPM And Startup Key", "TPM And PIN And Startup Key", "Public Key", "Passphrase", "TPM Certificate", "CryptoAPI Next Generation (CNG) Protector"];
        for (var i = 0; i < len; i++){
            k = keys[i];
            landingPageDiv.append("<span class='driveLetterSpan'>Drive Letter:</span> " + k);
            drives[k]["Lock Status"] == "Locked" ? landingPageDiv.append(" <span class=\"driveLocked\">(" + drives[k]["Lock Status"] + ")</span> ") : "";
            drives[k]["Drive Type"] ? landingPageDiv.append("<br>Drive Type: " + drives[k]["Drive Type"]) : "";
            drives[k]["Label"] ? landingPageDiv.append(" | Label: " + drives[k]["Label"]) : "";
            drives[k]["Capacity"] ? landingPageDiv.append("<br>Capacity: " + drives[k]["Capacity"]) : "";
            drives[k]["Free Space"] ? landingPageDiv.append(" | Free Space: " + drives[k]["Free Space"]) : "";
            drives[k]["Encryption Method"] ? landingPageDiv.append("<br>Encryption: " + drives[k]["Encryption Method"]) : "";
            arKeyType.forEach(function (element){
                drives[k][element] && element == "Numerical password" ? landingPageDiv.append("<br>Recovery Key ID: " + drives[k][element]) : "";
            });

            landingPageDiv.append("<br><br>");
            drivesSorted[k] = drives[k];
        }
    }
    catch (err) {
        $('#page-landing').append(err.message);
    }

    return drivesSorted;
}
*/

/**********************************************************************************************************************
*						        Launch Button Functions
**********************************************************************************************************************/
function launchOSD() {
    var capturedVars = getInputValues();
    var tpmDeferred = $.Deferred();
    var osdDeferred = $.Deferred();
    var cnDeferred = $.Deferred();

    //Check if TPM is enabled (TPM checkbox checked)
    //If TPMCheckBox is false, show warning popup, else resolve the deferred object immediately
    var message = "TPM is not functioning or not present!";
    capturedVars.TPMCheckBox != "true" ? warnUser(tpmDeferred, message, true) : tpmDeferred.resolve(true);

    //Check if OSD checkbox is checked
    //If OSDCheckBox is false, show warning popup, else resolve the deferred object immediately
    tpmDeferred.done(function (tpmResult) {
        if (tpmResult == false) { return };
        message = "OSD Checkbox is not checked!";
        capturedVars.OSDCheckBox != true ? warnUser(osdDeferred, message, false) : osdDeferred.resolve(true);
    });

    //Check if computer name is filled in
    //If computer name is not filled in, show warning popup, else resolve the deferred object immediately
    osdDeferred.done(function (osdResult) {
        if (osdResult == false) { return };
        message = "Computer name is not filled in!";
        capturedVars.compName == "" ? warnUser(cnDeferred, message, false) : cnDeferred.resolve(true);
    });

    cnDeferred.done(function (cnResult) {
        if (cnResult == false) { return };
        $(".pages").mCustomScrollbar("disable");
        $("#general-output-scroll").mCustomScrollbar("disable");
        try {
            ButtonFinishClick();
        }
        catch (err) {
            document.getElementById("general-output").innerHTML = err.message;
        }
    });
}

function launchFNF() {
    var capturedVars = getInputValues();
    var selectedUsers = {};
    $('#input-usmt-usernames :selected').each(function () {
        selectedUsers[$(this).text()] = $(this).val();
    });

    for (user in selectedUsers) {
        capturedVars[user] = selectedUsers[user];
    }
    var capturedVarsString = JSON.stringify(capturedVars);

    var tpmDeferred = $.Deferred();     //TPM checkbox
    var osdDeferred = $.Deferred();     //OSD checkbox
    var cnDeferred = $.Deferred();      //Computer Name field
    var wdDeferred = $.Deferred();      //Windows Drive field
    var puDeferred = $.Deferred();      //Primary Username field
    var edDeferred = $.Deferred();      //External Drive field
    var usmtDeferred = $.Deferred();      //USMT usernames select

    //Check if TPM is enabled (TPM checkbox checked)
    //If TPMCheckBox is false, show warning popup, else resolve the deferred object immediately
    var message = "TPM is not functioning or not present!";
    capturedVars.TPMCheckBox != "true" ? warnUser(tpmDeferred, message, true) : tpmDeferred.resolve(true);

    //Check if OSD checkbox is checked
    //If OSDCheckBox is false, show warning popup, else resolve the deferred object immediately
    tpmDeferred.done(function (tpmResult) {
        if (tpmResult == false) { return };
        message = "OSD Checkbox is not checked!";
        capturedVars.fnfOsdCheckBox != true ? warnUser(osdDeferred, message, false) : osdDeferred.resolve(true);
    });

    //Check if Computer Name is filled in
    //If computer name is not filled in, show warning popup, else resolve the deferred object immediately
    osdDeferred.done(function (osdResult) {
        if (osdResult == false) { return };
        message = "Computer name is not filled in!";
        capturedVars.compName == "" ? warnUser(cnDeferred, message, false) : cnDeferred.resolve(true);
    });

    //Check if Windows Drive is filled in
    //If Windows Drive is not filled in, show warning popup, else resolve the deferred object immediately
    cnDeferred.done(function (cnResult) {
        if (cnResult == false) { return };
        message = "Windows Drive Letter is not filled in!";
        capturedVars.windowsDrive == "" ? warnUser(wdDeferred, message, false) : wdDeferred.resolve(true);
    });

    //Check if Primary Username is filled in
    //If Primary Username is not filled in, show warning popup, else resolve the deferred object immediately
    wdDeferred.done(function (wdResult) {
        if (wdResult == false) { return };
        message = "Primary Username is not filled in!";
        capturedVars.primaryUsername == "" ? warnUser(puDeferred, message, false) : puDeferred.resolve(true);
    });

    //Check if External Drive is filled in
    //If External Drive is not filled in, show warning popup, else resolve the deferred object immediately
    puDeferred.done(function (puResult) {
        if (puResult == false) { return };
        message = "External Drive Letter is not filled in!";
        capturedVars.externalDrive == "" ? warnUser(edDeferred, message, false) : edDeferred.resolve(true);
    });

    //Check if USMT Usernames are selected
    //If USMT Usernames are not selected, show warning popup, else resolve the deferred object immediately
    edDeferred.done(function (edResult) {
        if (edResult == false) { return };
        message = "No USMT Usernames are selected!";
        Object.keys(selectedUsers).length > 0 ? usmtDeferred.resolve(true) : warnUser(usmtDeferred, message, false);
    });

    //If we make it past all warning checks, run flushfill function
    usmtDeferred.done(function (usmtResult) {
        if (capturedVars.scanStateCheckBox == false) { return };
        if (usmtResult == false) { return };
        $(".pages").mCustomScrollbar("destroy");
        $("#general-output-scroll").mCustomScrollbar("destroy");
        try {
            runFlushFill();
            //runFlushFill2(capturedVarsString);
        }
        catch (err) {
            document.getElementById("general-output").innerHTML = err.message;
            //Re-enable scroll bar replacement
            setTimeout(function () {
                $(".pages").mCustomScrollbar("update");
                $("#general-output-scroll").mCustomScrollbar("update");
            }, 2000);
        }
    });
}

function launchDISM() {
    var capturedVars = getInputValues();
    var wdDeferred = $.Deferred();      //Windows Drive field
    var puDeferred = $.Deferred();      //Primary Username field
    var edDeferred = $.Deferred();      //External Drive field

    var message = "Windows Drive is not filled in!";
    capturedVars.windowsDrive == "" ? warnUser(wdDeferred, message, false) : wdDeferred.resolve(true);

    wdDeferred.done(function (wdResult) {
        if (wdResult == false) { return };
        message = "Primary Username is not filled in!";
        capturedVars.primaryUsername == "" ? warnUser(puDeferred, message, false) : puDeferred.resolve(true);
    });

    puDeferred.done(function (puResult) {
        if (puResult == false) { return };
        message = "External Drive Letter is not filled in!";
        capturedVars.externalDrive == "" ? warnUser(edDeferred, message, false) : edDeferred.resolve(true);
    });

    edDeferred.done(function (edResult) {
        if (edResult == false) { return };

        //Disable scroll bar replacement because it causes a slow script warning after dism finishes
        $(".pages").mCustomScrollbar("disable");
        $("#general-output-scroll").mCustomScrollbar("disable");

        var dismReturn = dismCapture();

        //Check if DISM was successful
        var dismDeferred = $.Deferred();
        var message = "DISM operation failed!";
        dismReturn != 0 ? warnUser(dismDeferred, message, false) : dismDeferred.resolve(true);
    });

    //Re-enable scroll bar replacement
    setTimeout(function () {
        $(".pages").mCustomScrollbar("update");
        $("#general-output-scroll").mCustomScrollbar("update");
    }, 2000);
}

function launchScanstate() {
    var selectedUsers = {};
    var scanReturn = 1;     //scanReturn should be set by the vbscript scanState function. It should set to 0 if scanstate completed successfully. Initialize to something other than 0 here so the program will ONLY continue if scanstate runs ans is successful
    $('#input-usmt-usernames :selected').each(function () {
        selectedUsers[$(this).text()] = $(this).val();
    });

    var capturedVars = getInputValues();
    for (user in selectedUsers) {
        capturedVars[user] = selectedUsers[user];
    }
    var wdDeferred = $.Deferred();      //Windows Drive field
    var puDeferred = $.Deferred();      //Primary Username field
    var edDeferred = $.Deferred();      //External Drive field
    var usmtDeferred = $.Deferred();    //USMT usernames select

    var message = "Windows Drive is not filled in!";
    capturedVars.windowsDrive == "" ? warnUser(wdDeferred, message, false) : wdDeferred.resolve(true);

    wdDeferred.done(function (wdResult) {
        if (wdResult == false) { return };
        message = "Primary Username is not filled in!";
        capturedVars.primaryUsername == "" ? warnUser(puDeferred, message, false) : puDeferred.resolve(true);
    });

    puDeferred.done(function (puResult) {
        if (puResult == false) { return };
        message = "External Drive Letter is not filled in!";
        capturedVars.externalDrive == "" ? warnUser(edDeferred, message, false) : edDeferred.resolve(true);
    });

    edDeferred.done(function (edResult) {
        if (edResult == false) { return };
        message = "No USMT Usernames are selected!";
        Object.keys(selectedUsers).length > 0 ? usmtDeferred.resolve(true) : warnUser(usmtDeferred, message, false);
    });

    usmtDeferred.done(function (usmtResult) {
        if (usmtResult == false) { return };

        $(".pages").mCustomScrollbar("disable");
        $("#general-output-scroll").mCustomScrollbar("disable");

        scanReturn = usmtScanstate(true);
    });
    
    setTimeout(function () {
        $(".pages").mCustomScrollbar("update");
        $("#general-output-scroll").mCustomScrollbar("update");
    }, 2000);
}

function launchLoadstate() {
    var capturedVars = getInputValues();

    var puDeferred = $.Deferred();      //Primary Username field
    var edDeferred = $.Deferred();      //External Drive field

    var message = "Primary Username is not filled in!";
    capturedVars.primaryUsername == "" ? warnUser(puDeferred, message, false) : puDeferred.resolve(true);

    puDeferred.done(function (puResult) {
        if (puResult == false) { return };
        message = "External Drive Letter is not filled in!";
        capturedVars.externalDrive == "" ? warnUser(edDeferred, message, false) : edDeferred.resolve(true);
    });

    edDeferred.done(function (edResult) {
        if (edResult == false) { return };

        $(".pages").mCustomScrollbar("disable");
        $("#general-output-scroll").mCustomScrollbar("disable");

        usmtLoadstate();

        setTimeout(function () {
            $(".pages").mCustomScrollbar("update");
            $("#general-output-scroll").mCustomScrollbar("update");
        }, 2000);
    });
}

function launchSoftwareInstall() {
    var softwareInstall = false;
    $(".accordion .software-install").each(function () {
        //If any software install boxes are checked, continue with software installation
        if($(this).prop("checked") == true){
            softwareInstall = true;
        }
    });
    if(softwareInstall == false){return;}

    try {
        ButtonFinishClick();
    }
    catch (err) {
        document.getElementById("general-output").innerHTML = err.message;
    }
}

/**********************************************************************************************************************
*						        Buttons Inside List Elements
**********************************************************************************************************************/
function blInfoClick() {
    var target = document.getElementById('general-output');
    var opts = {
        color: '#FFF'
    }
    var spinner = new Spinner(opts).spin(target);
    WMIListDrives();
    //BitlockerInfo();
    spinner.stop(target);

}

function blUnlockClick() {
    var target = document.getElementById('general-output');
    var opts = {
        color: '#FFF'
    }
    var spinner = new Spinner(opts).spin(target);
    BitlockerUnlock();
    // BitlockerInfo();
    WMIListDrives();
    dimElements();
    spinner.stop(target);
}

function showUsersClick() {
    var dlDeferred = $.Deferred();
    var inputWindowsDrive = $("#input-windows-drive").val().toUpperCase();          //Windows drive letter input by user
    var determinedWindowsDrive = $("#windows-drive-letter").val().toUpperCase();    //Windows drive letter determined by listDrives vbscript function
    var winDirFound = ReportFolderStatus(inputWindowsDrive + ":\\Windows");

    //This seems flimsy. Maybe try checking for C:\Windows or inputWindowsDrive\Windows instead?
    if (inputWindowsDrive == "") {
        //Check if windows drive letter is filled in
        message = "Windows drive letter is not filled in!";
        warnUser(dlDeferred, message, false);
    }
    else if (determinedWindowsDrive == "") {
        //Check if windows drive encryption is unlocked
        message = "Windows drive not found! It may still be encrypted.";
        warnUser(dlDeferred, message, true);
    }
    else if (inputWindowsDrive != determinedWindowsDrive) {
        //Check if windows drive letter entered by user matches what was found in listDrives function
        message = "Windows drive letter does not match volume labeled 'Windows' on this machine. Are you sure?";
        warnUser(dlDeferred, message, true);
    }
    else { dlDeferred.resolve(true); }
    dlDeferred.done(function (continueFunction) {
        if (continueFunction == false) { return };
        if (winDirFound == false) {
            //Check if windows directory exists
            message = "Windows directory not found at " + inputWindowsDrive + ":\\Windows";
            warnUser(dlDeferred, message, false);
            return;
        }
        //If we made it past all of the error checks, try to enumerate users for the given Windows drive
        try {
            enumUsers();
        }
        catch (err) {
            document.getElementById("general-output").innerHTML = err.message;
        }
    });
}

function jsMapNetworkDrive() {
    var sharePath = $("#input-share-path").val();
    $("#dialog").dialog({
        resizable: false,
        dialogClass: "no-close",
        autoOpen: false,
        height: "auto",
        width: 400,
        modal: true,
        title: "Map Network Drive",
        hide: { effect: "explode", duration: 200 }
    });
    $("#dialog").text("Mapping " + sharePath + " as drive N:");
    $("#dialog").dialog("open");
    setTimeout(function () {
        try {
            mapNetDrive();
        }
        catch (e) { };
        $("#dialog").dialog("close");
    }, 1000);

}

/**********************************************************************************************************************
*						        Verification/Warning Functions
**********************************************************************************************************************/
function warnUser(objDeferred, message, continueButton) {
    $("#dialog").dialog({
        resizable: false,
        dialogClass: "no-close",
        autoOpen: false,
        height: "auto",
        width: 400,
        modal: true,
        hide: { effect: "explode", duration: 200 }
    });

    if (continueButton == true) {
        $("#dialog").dialog({
            buttons: {
                "Continue anyway": function () {
                    $(this).dialog("close");
                    objDeferred.resolve(true);
                },
                Cancel: function () {
                    $(this).dialog("close");
                    objDeferred.resolve(false);
                }
            }
        });
    }
    else {
        $("#dialog").dialog({
            buttons: {
                Cancel: function () {
                    $(this).dialog("close");
                    objDeferred.resolve(false);
                }
            }
        });
    }
    $("#dialog").text(message);
    $("#dialog").dialog("open");
}

function ReportFolderStatus(fldr) {
    var fso;
    try {
        fso = new ActiveXObject("Scripting.FileSystemObject");
        // fldr = "blah";
        //If windows directory doesn't exist, add win-not-found class to USMT, DISM, FnF, and Software Install buttons
        if (fso.FolderExists(fldr)) {
            return true;
        }
        else {
            return false;
        }
    }
    catch (err) {
        document.getElementById("general-output").innerHTML = err.message;
    }
}

function dimElements() {
    if (objDrives["Windows Drive Found"] == false) {
        //No Windows drive found
        //Dim elements that require a Windows directory as a warning
        if ($(".li-win-not-found").length == 0){
            $(".li-win-required").addClass("li-win-not-found");
            $(".li-win-required .ui-widget-content").addClass("content-win-not-found");
            $(".li-win-required .ui-accordion-header").addClass("h3-win-not-found");
            $(".require-win-dir").addClass("win-not-found");
            $(".require-win-dir").append("<img src=\"images/alert-25.svg\" class=\"warning-image\" title=\"Encrypted drive needs to be unlocked\">"); 
        }

    }
    else {
        //Windows drive found. Remove dimming effect and show all list items
        $(".li-win-required").removeClass("li-win-not-found");
        $(".li-win-required .ui-widget-content").removeClass("content-win-not-found");
        $(".li-win-required .ui-accordion-header").removeClass("h3-win-not-found");
        $(".require-win-dir").removeClass("win-not-found");
        $(".accordion .ui-widget-content").show("fast");
        $(".accordion .ui-widget-content").prev().children().removeClass("ui-icon-triangle-1-e ui-icon-triangle-1-s").addClass("ui-icon-triangle-1-s");
        $(".warning-image").remove();
    }
}

function writeToLog(text){
    if (isAdmin == false){return;}
    try{
        var stamp = new Date();
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        var ForReading = 1, ForAppending = 8;
        var strCurrentPath = window.location.pathname;
        var strLogDir = strCurrentPath.substring(0, strCurrentPath.lastIndexOf('\\'));
        var htaLog = fso.OpenTextFile(strLogDir + "\\HTALOG.txt", ForAppending, true);
        htaLog.writeLine("<<" + stamp + ">> " + text);
        htaLog.close();
    }
    catch(err){
    }
}