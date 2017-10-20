/* Globals */
var winFoundAndEncrypted = [];

/************************************ JavaScript functions ************************************/
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
)

function loadPage(targetOperation, event) {
    //Rerun listDrives function if returning to landing page
    if (targetOperation == "landing") {
        //Find and list drives attached to this machine
        //If no windows directory was found and an encrypted drive WAS found, recolor list items that require a windows directory
        isWindowsFoundIsEncrypted();
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
            //Only show the Bitlocker Unlock template if an encrypted drive is found and locked
            if (!($(this).attr("id") == "temp-bitlocker-unlock" && winFoundAndEncrypted["Encrypted"] == false)){
                template += $(this).html();
            }
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
    //$("#general-output-scroll").css('visibility', 'hidden');

    dimElements();
}

// ******************** Misc functions ********************
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

function jsMapNetworkDrive(){
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
    setTimeout(function() {
        try {
            mapNetDrive();
        }
        catch (e) { };
        $("#dialog").dialog("close");
    }, 1000);

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

        //Re-enable scroll bar replacement
        setTimeout(function () {
            $(".pages").mCustomScrollbar("update");
            $("#general-output-scroll").mCustomScrollbar("update");
        }, 2000);

        //Check if DISM was successful
        var dismDeferred = $.Deferred();
        var message = "DISM operation failed!";
        dismReturn != 0 ? warnUser(dismDeferred, message, false) : dismDeferred.resolve(true);
    });
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

        usmtScanstate(true);

        setTimeout(function () {
            $(".pages").mCustomScrollbar("update");
            $("#general-output-scroll").mCustomScrollbar("update");
        }, 2000);
    });
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
    // BitlockerInfo();
    jsListDrives();
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

function isWindowsFound() {
    var determinedWindowsDrive = $("#windows-drive-letter").val();
    var winDirFound = ReportFolderStatus(determinedWindowsDrive + ":\\Windows");
    var winDeferred = $.Deferred();

    if (winDirFound == false) {
        return false;
    }
    else {
        return true;
    }
}

function isEncrypted(driveList) {
    var k, encryptedDrives = [];
    for (k in driveList) {
        if (driveList[k]["Lock Status"] == "Locked") {
            encryptedDrives.push(k);
        }
    }

    if (encryptedDrives.length > 0) {
        return true;
    }
    else return false;
}

function isWindowsFoundIsEncrypted() {
    var driveList = {};
    driveList = jsListDrives();
    //Check for a Windows directory. Returns true or false
    var windowsFound = isWindowsFound();
    //Check for any encrypted drives. Returns true if any are found, false otherwise
    var encryptedDriveFound = isEncrypted(driveList);
    var results = [];

    winFoundAndEncrypted["WinFound"] = windowsFound;
    winFoundAndEncrypted["Encrypted"] = encryptedDriveFound;

    dimElements();

    return results;
}

function dimElements() {
    if (winFoundAndEncrypted["WinFound"] == false && winFoundAndEncrypted["Encrypted"] == true) {
        //No Windows drive found and an encrypted drive WAS found
        //Dim elements that require a Windows directory as a warning
        if ($(".li-win-not-found").length == 0){
            $(".li-win-required").addClass("li-win-not-found");
            $(".li-win-required .ui-widget-content").addClass("content-win-not-found");
            $(".li-win-required .ui-accordion-header").addClass("h3-win-not-found");
            $(".require-win-dir").addClass("win-not-found");
            $(".require-win-dir").append("<img src=\"images/alert-25.svg\" class=\"warning-image\">"); 
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

$(".on-load").delay(200).animate({ "opacity": "1" }, 700);
