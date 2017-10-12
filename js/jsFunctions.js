/************************************ JavaScript functions ************************************/
$(document).ready(function () {
    $('.pages').css('display', 'none');
    $('.header-text').css('display', 'none');
    $('#page-landing').css('display', 'inline-block');
    $('#header-landing').css('display', 'inline');
    $('#general-output-scroll').css('visibility', 'hidden');

    //Find and list drives attached to this machine
    jsListDrives();
    try {
        //Check for a TPM chip
        TPMCheck();
    }
    catch (err) {
        $('#page-landing').append(err.message);
    }

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
});

// ******************* Placeholder Text ***********************
$('[placeholder]').focus(function () {
    var input = $(this);
    if (input.val() == input.attr('placeholder')) {
        input.val('');
        input.removeClass('placeholder');
    }
}).blur(function () {
    var input = $(this);
    if (input.val() == '' || input.val() == input.attr('placeholder')) {
        input.addClass('placeholder');
        input.val(input.attr('placeholder'));
    }
}).blur().parents('form').submit(function () {
    $(this).find('[placeholder]').each(function () {
        var input = $(this);
        if (input.val() == input.attr('placeholder')) {
            input.val('');
        }
    })
});

// ******************** Navigation Buttons ********************
$('#button-ShowAll').on('click', function () {
    $('ul .ui-widget-content').show('fast');
    $('ul .ui-widget-content').prev().children().removeClass("ui-icon-triangle-1-e ui-icon-triangle-1-s").addClass("ui-icon-triangle-1-s");
});

$('#button-HideAll').on('click', function () {
    $('ul .ui-widget-content').hide('fast');
    $('ul .ui-widget-content').prev().children().removeClass("ui-icon-triangle-1-e ui-icon-triangle-1-s").addClass("ui-icon-triangle-1-e");
});

$('.button-nav').on('click', function () {
    var targetID = $(this).attr("id");      //Get the id of the button that was clicked
    var targetOperation = targetID.replace("button-", "");  //Strip away button-, leaving us with osd, flushfill, dism, usmt, software, or tools
    
    //if (windowsFound == false && !(targetID == "button-osd" || targetID == "button-tools" || targetID == "button-home" || targetID == "button-exit")){return}

    //Prevent recreating the page if we are already on it. i.e. clicking dism button when we are already on dism page.
    var isSamePage = $("#ul-" + targetOperation).children().length; 
    if (isSamePage == 0) {
        //If navigating to a new page, clear out the list elements so we can reuse list items on the new page.
        $(".ul-elements").empty();
    }
    else return;

    //Grab the appropriate templates for the page we are navigating to and join them together in a variable named "template"
    var templates = $("." + targetOperation);
    var template = "";

    templates.each(function () {
        template += $(this).html();
    });

    //Insert the joined templates into the page we are nevigating to
    $("#ul-" + targetOperation).append(template);
    //If no windows directory was found and an encrypted drive WAS found, recolor list items that require a windows directory
    var winFoundAndEncrypted = isWindowsFoundIsEncrypted();

    initAccordion();
    $('.pages').css('display', 'none');
    $('.header-text').css('display', 'none');
    $('.button-nav').removeClass('active');
    $(this).addClass('active');
    $('#page-' + targetOperation).css('display', 'inline-block');
    $('#header-' + targetOperation).css('display', 'inline');
    $("#general-output-scroll").css('visibility', 'hidden');
});

$(document).on('click', '.show-output', function () {
    $("#general-output-scroll").css('visibility', 'visible');
});

$('#button-home').on('click', function () {
    jsListDrives();
    $('#page-landing').css('display', 'inline-block');
    $('#header-landing').css('display', 'inline');
});
$('.button-exithome').hover(
    function () {
        $(this).children().addClass("button-background-hover");
    }, function () {
        $(this).children().removeClass("button-background-hover");
    }
);

// ******************** Misc functions ********************
function getInputValues() {
    var inputValues = {};
    $("ul .input-objects").each(function () {
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

function initAccordion() {
    //Initialize accordion for displaying list itmes
    $('.accordion h3').click(function () {
        $(this).next().toggle('fast');
        $(this).children().toggleClass("ui-icon-triangle-1-e ui-icon-triangle-1-s");
        return false;
    }).next().hide();
    $('ul .ui-widget-content').not($(".li-win-not-found .ui-widget-content")).show('fast');
    $('ul .ui-widget-content').not($(".li-win-not-found .ui-widget-content")).prev().children().removeClass("ui-icon-triangle-1-e ui-icon-triangle-1-s").addClass("ui-icon-triangle-1-s");
}

function jsListDrives() {
    //Get Drive Letter, Label, Capacity, and Encryption status for each drive that has a Letter mapped to it
    //drives should contain: Drive Letter, Protection Status, Encryption Method, Lock Status, Key Type, Key ID, Label, Capacity 
    var drives = {}, drivesSorted = {};
    try {
        drives = listDrives();
        delete drives["setProp"];   //Remove the setProp function from the object before looping through the drive letters
        
        var keys = [], k, len;

        for (k in drives) {
            if (drives.hasOwnProperty(k)) {
                keys.push(k);
            }
        }
        keys.sort();
        len = keys.length;
        //alert(JSON.stringify(drives));
        var landingPageDiv = $("#drive-list-output");
        landingPageDiv.empty();
        for (var i = 0; i < len; i++) {
            k = keys[i];
            landingPageDiv.append("<span class='driveLetterSpan'>Drive Letter:</span> " + k);
            drives[k]["Lock Status"] == "Locked" ? landingPageDiv.append(" <span class=\"driveLocked\">(" + drives[k]["Lock Status"] + ")</span> ") : "";
            drives[k]["Label"] ? landingPageDiv.append(" | Label: " + drives[k]["Label"]) : "";
            drives[k]["Capacity"] ? landingPageDiv.append("<br>Capacity: " + drives[k]["Capacity"]) : "";
            drives[k]["Encryption Method"] ? landingPageDiv.append(" | Encryption: " + drives[k]["Encryption Method"]) : "";
            drives[k]["Key ID"] ? landingPageDiv.append("<br>Key ID: " + drives[k]["Key ID"]) : "";
            landingPageDiv.append("<br><br>");
            drivesSorted[k] = drives[k];
        }
    }
    catch (err) {
        document.getElementById("general-output").innerHTML = err.message;
    }

    isWindowsFound();
    return drivesSorted;
}

// ******************** Verification/Error Checking Functions ********************
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


/*******************************
************* OSD **************
*******************************/
function launchOSD() {
    var capturedVars = getInputValues();
    var tpmDeferred = $.Deferred();
    var osdDeferred = $.Deferred();
    var cnDeferred = $.Deferred();

    //Check if TPM is enabled (TPM checkbox checked)
    //If TPMCheckBox is false, show warning popup, else resolve the deferred object immediately
    var message = "TPM is not functioning or not present!";
    capturedVars.TPMCheckBox != true ? warnUser(tpmDeferred, message, true) : tpmDeferred.resolve(true);

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

/*******************************
************* FnF **************
*******************************/
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

    var tpmDeferred = $.Deferred();
    var osdDeferred = $.Deferred();
    var cnDeferred = $.Deferred();

    //Check if TPM is enabled (TPM checkbox checked)
    //If TPMCheckBox is false, show warning popup, else resolve the deferred object immediately
    var message = "TPM is not functioning or not present!";
    capturedVars.TPMCheckBox != true ? warnUser(tpmDeferred, message, true) : tpmDeferred.resolve(true);

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
            runFlushFill();
            //runFlushFill2(capturedVarsString);
        }
        catch (err) {
            document.getElementById("general-output").innerHTML = err.message;
        }
    });
}

/*******************************
************* DISM *************
*******************************/
function launchDISM() {
    var capturedVars = getInputValues();

    //Disable scroll bar replacement because it causes a slow script warning after dism finishes
    $(".pages").mCustomScrollbar("disable");
    $("#general-output-scroll").mCustomScrollbar("disable");

    var dismReturn = dismCapture();

    //Re-enable scroll bar replacement
    setTimeout(function () {
        $(".pages").mCustomScrollbar("update");
        $("#general-output-scroll").mCustomScrollbar("update");
    }, 2000);

    //Check if DISM was successful
    var dismDeferred = $.Deferred();
    var message = "DISM operation failed!";
    dismReturn != 0 ? warnUser(dismDeferred, message, false) : dismDeferred.resolve(true);
}

/*******************************
************* USMT *************
*******************************/
function launchScanstate() {
    var selectedUsers = {};
    $('#input-usmt-usernames :selected').each(function () {
        selectedUsers[$(this).text()] = $(this).val();
    });

    var capturedVars = getInputValues();
    for (user in selectedUsers) {
        capturedVars[user] = selectedUsers[user];
    }

    $(".pages").mCustomScrollbar("disable");
    $("#general-output-scroll").mCustomScrollbar("disable");

    usmtScanstate("true");

    setTimeout(function () {
        $(".pages").mCustomScrollbar("update");
        $("#general-output-scroll").mCustomScrollbar("update");
    }, 2000);
}

function launchLoadstate() {
    var capturedVars = getInputValues();

    $(".pages").mCustomScrollbar("disable");
    $("#general-output-scroll").mCustomScrollbar("disable");

    usmtLoadstate();

    setTimeout(function () {
        $(".pages").mCustomScrollbar("update");
        $("#general-output-scroll").mCustomScrollbar("update");
    }, 2000);
}

$('#button-ScanState').on('click', function () {
    $('.scanState').css('display', 'list-item');
    $('#run-usmt').css('display', 'inline-block');
    $('#run-LoadState').css('display', 'none');
});

$('#button-LoadState').on('click', function () {
    $('.scanState').css('display', 'none');
    $('#run-LoadState').css('display', 'inline-block');
    $('#run-usmt').css('display', 'none');
});

//*******************Bitlocker loading spinner*****************
function blInfoClick() {
    $("#general-output-scroll").css('visibility', 'visible');
    var target = document.getElementById('general-output');
    var opts = {
        color: '#FFF'
    }
    var spinner = new Spinner(opts).spin(target);
    BitlockerInfo();
    spinner.stop(target);

}
function blUnlockClick() {
    $("#general-output-scroll").css('visibility', 'visible');
    var target = document.getElementById('general-output');
    var opts = {
        color: '#FFF'
    }
    var spinner = new Spinner(opts).spin(target);
    BitlockerUnlock();
    BitlockerInfo();
    isWindowsFoundIsEncrypted();
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
    dlDeferred.done(function (continueFunction) {
        if (continueFunction == false) { return };
        if (winDirFound == false) {
            //Check if windows directory exists
            message = "Windows directory not found at " + inputWindowsDrive + ":\\Windows";
            warnUser(dlDeferred, message, false);
            return;
        }
    });

    //If we made it past all of the error checks, try to enumerate users for the given Windows drive
    try {
        enumUsers();
    }
    catch(err) {
        document.getElementById("general-output").innerHTML = err.message;
    }
}

function ReportFolderStatus(fldr) {
    var fso;
    var navButtons = $(".require-win-dir");
    try{
        fso = new ActiveXObject("Scripting.FileSystemObject");
fldr = "blah";
        //If windows directory doesn't exist, add win-not-found class to USMT, DISM, FnF, and Software Install buttons
        if (fso.FolderExists(fldr)) {
            navButtons.removeClass("win-not-found");
            return true;
        }
        else {
            navButtons.addClass("win-not-found");
            return false;
        }
    }
    catch(err){
        navButtons.addClass("win-not-found");
        document.getElementById("general-output").innerHTML = err.message;
    }
        
}

function isWindowsFound(){
    var determinedWindowsDrive = $("#windows-drive-letter").val();
    var winDirFound = ReportFolderStatus(determinedWindowsDrive + ":\\Windows");
    var winDeferred = $.Deferred();
    if(winDirFound == false) {
        return false;
    }
    else{
        return true;
    }
}

function isEncrypted(){
    var driveList = {}, k, encryptedDrives = [];
    driveList = jsListDrives();
    for (k in driveList) {
        if (driveList[k]["Encryption"] == "Encrypted") {
            encryptedDrives.push(k);
        }
    }

    if(encryptedDrives.length > 0){
        return true;
    }
    else return false;
}

function isWindowsFoundIsEncrypted(){
    //Check for a Windows directory. Returns true or false
    var windowsFound = isWindowsFound();
    //Check for any encrypted drives. Returns true if any are found, false otherwise
    var encryptedDriveFound = isEncrypted();
    var results = [];

    results.WinFound = windowsFound;
    results.Encrypted = encryptedDriveFound;

    if (results["WinFound"] == false && results["Encrypted"] == true) {
        $(".li-win-required").addClass("li-win-not-found");
        $(".li-win-required .ui-widget-content").addClass("content-win-not-found");
        $(".li-win-required .ui-accordion-header").addClass("h3-win-not-found");
    }
    else{
        $(".li-win-required").removeClass("li-win-not-found");
        $(".li-win-required .ui-widget-content").removeClass("content-win-not-found");
        $(".li-win-required .ui-accordion-header").removeClass("h3-win-not-found");
    }
    return results;

}
// $("#input-windows-drive, #input-primary-username").keyup(function () {
//     usmtScanstate("false");
// });

$(".on-load").delay(200).animate({ "opacity": "1" }, 700);
