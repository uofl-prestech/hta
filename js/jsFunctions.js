/************************************ JavaScript functions ************************************/
$(document).ready(function () {
    $('.pages').css('display', 'none');
    $('.header-text').css('display', 'none');
    $('#page-landing').css('display', 'inline-block');
    $('#header-landing').css('display', 'inline');
    $('#general-output-scroll').css('visibility', 'hidden');
    try {
        jsListDrives();
        TPMCheck();
    }
    catch (err) {
        $('#page-landing').append(err.message);
    }

    $(".pages").mCustomScrollbar({
        theme: "minimal",
        live: "once"
    });
    $("#general-output-scroll").mCustomScrollbar({
        theme: "minimal",
        live: "once"
    });
    $("#dialog").dialog({
        autoOpen: false
    });
});

function initAccordion() {
    $('.accordion h3').click(function () {
        $(this).next().toggle('fast');
        $(this).children().toggleClass("ui-icon-triangle-1-e ui-icon-triangle-1-s");
        return false;
    }).next().hide();
    $('ul .ui-widget-content').show('fast');
    $('ul .ui-widget-content').prev().children().removeClass("ui-icon-triangle-1-e ui-icon-triangle-1-s").addClass("ui-icon-triangle-1-s");
}

//Get Drive Letter, Label, Capacity, and Encryption status for each drive that has a Letter mapped to it
function jsListDrives() {
    var drives = {};
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

    var landingPageDiv = $("#drive-list-output");
    for (var i = 0; i < len; i++) {
        k = keys[i];
        landingPageDiv.append("<span class='driveLetterSpan'>Drive Letter:</span> " + drives[k]["Drive Letter"]);
        drives[k]["Label"] ? landingPageDiv.append(" | Label: " + drives[k]["Label"]) : "";
        drives[k]["Capacity"] ? landingPageDiv.append("<br>Capacity: " + drives[k]["Capacity"]) : "";
        drives[k]["Encryption"] ? landingPageDiv.append(" | Encryption: " + drives[k]["Encryption"]) : "";
        landingPageDiv.append("<br><br>");
    }
}
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
// ******#general-output-scroll hidden/visible functions*******
$('.button-nav').on('click', function () {
    $('.pages').css('display', 'none');
    $('.header-text').css('display', 'none');
    $('.button-nav').removeClass('active');
    $(this).addClass('active');
    $("#general-output-scroll").css('visibility', 'hidden');
});
$(document).on('click', '.show-output', function () {
    $("#general-output-scroll").css('visibility', 'visible');
});

$('#button-osd').on('click', function () {
    $('#page-osd').css('display', 'inline-block');
    $('#header-osd').css('display', 'inline');

    if ($("#ul-osd").children().length == 0) {
        $(".ul-elements").empty();
        var template = $("#temp-osd-options").html();
        template += $("#temp-computer-name").html();
        template += $("#temp-software-install").html();
        $("#ul-osd").append(template);
        initAccordion();
    }

    try {
        TPMCheck();
    }
    catch (err) {
        document.getElementById("general-output").innerHTML = err.message;
    }
    $('#input-osd-checkbox').change(function () {
        if (this.checked == true) {
            $('#MBAM').prop("checked", true);
        }
        else {
            $('#MBAM').prop("checked", false);
        }
    });
});

$('#button-flushfill').on('click', function () {
    $('#page-flushfill').css('display', 'inline-block');
    $('#header-flushfill').css('display', 'inline');

    if ($("#ul-fnf").children().length == 0) {
        $(".ul-elements").empty();
        var template = $("#temp-fnf-options").html();
        template += $("#temp-bitlocker-info").html();
        template += $("#temp-windows-drive").html();
        template += $("#temp-bitlocker-unlock").html();
        template += $("#temp-select-users").html();
        template += $("#temp-primary-username").html();
        template += $("#temp-external-drive").html();
        template += $("#temp-computer-name").html();
        template += $("#temp-software-install").html();
        $("#ul-fnf").append(template);
        initAccordion();
    }
    // if ($("#input-usmt-usernames").length == 0) {
    //     try {
    //         enumUsers();
    //     }
    //     catch (err) {
    //         document.getElementById("general-output").innerHTML = err.message;
    //     }
    // }
});

$('#button-dism').on('click', function () {
    $('#page-dism').css('display', 'inline-block');
    $('#header-dism').css('display', 'inline');

    if ($("#ul-dism").children().length == 0) {
        $(".ul-elements").empty();
        var template = $("#temp-bitlocker-info").html();
        template += $("#temp-windows-drive").html();
        template += $("#temp-bitlocker-unlock").html();
        template += $("#temp-primary-username").html();
        template += $("#temp-external-drive").html();
        $("#ul-dism").append(template);
        initAccordion();
    }
});

$('#button-usmt').on('click', function () {
    $('#page-usmt').css('display', 'inline-block');
    $('#header-usmt').css('display', 'inline');

    if ($("#ul-usmt").children().length == 0) {
        $(".ul-elements").empty();
        var template = $("#temp-bitlocker-info").html();
        template += $("#temp-windows-drive").html();
        template += $("#temp-bitlocker-unlock").html();
        template += $("#temp-select-users").html();
        template += $("#temp-primary-username").html();
        template += $("#temp-external-drive").html();
        $("#ul-usmt").append(template);
        initAccordion();
    }

    //if ($("#input-usmt-usernames").length == 0) {
    //enumUsers();
    //}
});

$('#button-software').on('click', function () {
    $('#page-software-install').css('display', 'inline-block');
    $('#header-software').css('display', 'inline');

    if ($("#ul-software-install").children().length == 0) {
        $(".ul-elements").empty();
        var template = $("#temp-software-install").html();
        $("#ul-software-install").append(template);
        initAccordion();
    }
});

$('#button-tools').on('click', function () {
    $('#page-tools').css('display', 'inline-block');
    $('#header-tools').css('display', 'inline');
});

$('#button-home').on('click', function () {
    $('#page-landing').css('display', 'inline-block');
    $('#header-landing').css('display', 'inline');

    // if ($("#ul-landing").children().length ==0) {
    //     $(".ul-elements").empty();
    //     var template = $("#temp-landing").html();
    //     $("#ul-landing").append(template);
    //     initAccordion();
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
    catch (err) {
        document.getElementById("general-output").innerHTML = err.message;
    }
}

function ReportFolderStatus(fldr) {
    var fso;
    fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.FolderExists(fldr))
        return true;
    else
        return false;
}
// $("#input-windows-drive, #input-primary-username").keyup(function () {
//     usmtScanstate("false");
// });

$(".on-load").delay(200).animate({ "opacity": "1" }, 700);
