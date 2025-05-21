// JSON parse fallback for older versions of ExtendScript
if (typeof JSON === 'undefined') {
    JSON = {};
}
if (typeof JSON.parse !== 'function') {
    JSON.parse = function (s) {
        return eval('(' + s + ')');
    };
}

// Polyfill for Array.prototype.indexOf in older ExtendScript
if (!Array.prototype.indexOf) {
    Array.prototype.indexOf = function(searchElement) {
        for (var i = 0; i < this.length; i++) {
            if (this[i] === searchElement) {
                return i;
            }
        }
        return -1;
    };
}

function loadEmailList() {
    var emailFile = new File("/Users/knippingt/Library/CloudStorage/OneDrive-SharedLibraries-YOOXNET-A-PORTERGROUP/O365G-Ecommerce-Studio - US Files/Retouch/Actions & Scripts/Photoshop Scripts/Extra Scripts/email_list.json");

    if (!emailFile.exists) {
        alert("email_list.json not found at /Users/knippingt/Library/CloudStorage/OneDrive-SharedLibraries-YOOXNET-A-PORTERGROUP/O365G-Ecommerce-Studio - US Files/Retouch/Actions & Scripts/Photoshop Scripts/Extra Scripts/email_list.json!");
        return null;
    }

    emailFile.open("r");
    var raw = emailFile.read();
    emailFile.close();

    try {
        return JSON.parse(raw);
    } catch (e) {
        alert("Error parsing email_list.json:\n" + e.message);
        return null;
    }
}

function showEmailForm(emailData) {
    var dialog = new Window('dialog', 'Easy Emails™');

    // Email category panel
    var emailPanel = dialog.add('panel', undefined, 'Select Recipients');
    emailPanel.orientation = 'row';
    emailPanel.alignChildren = 'top';

    var checkboxes = {};
    var categoryOrder = ["Photographers", "Stylists", "Managers", "Groups"];
    for (var i = 0; i < categoryOrder.length; i++) {
        var category = categoryOrder[i];
        if (!emailData[category]) continue;

        var column = emailPanel.add('panel', undefined, category);
        column.orientation = 'column';
        column.alignChildren = 'left';
        checkboxes[category] = [];

        var entries = emailData[category];
        for (var j = 0; j < entries.length; j++) {
            var cbGroup = column.add('group');
            cbGroup.orientation = 'row';
            var cb = cbGroup.add('checkbox', undefined, entries[j].name);
            cb.email = entries[j].email;
            cb.brandTypeTags = entries[j].brandTypeTags || [];

            var ccCheckbox = cbGroup.add('checkbox', undefined, 'CC');
            ccCheckbox.email = entries[j].email;

            // If CC is checked, force main checkbox checked
            ccCheckbox.onClick = (function(cbRef, ccRef) {
                return function () {
                    if (ccRef.value && !cbRef.value) {
                        cbRef.value = true;
                    }
                };
            })(cb, ccCheckbox);

            checkboxes[category].push({ checkbox: cb, ccCheckbox: ccCheckbox });
        }
    }

    // Brand, Type, List, and Subject on the same line
    var brandListSubjectGroup = dialog.add('group');
    brandListSubjectGroup.orientation = 'row';
    brandListSubjectGroup.alignChildren = 'top';

    // Brand radio buttons
    brandListSubjectGroup.add('statictext', undefined, 'Brand:');
    var brandGroup = brandListSubjectGroup.add('group');
    brandGroup.orientation = 'column';
    var brandNAP = brandGroup.add('radiobutton', undefined, 'NAP');
    var brandMRP = brandGroup.add('radiobutton', undefined, 'MRP');
    brandNAP.value = true;

    // Type dropdown
    brandListSubjectGroup.add('statictext', undefined, 'Type:');
    var typeDropdown = brandListSubjectGroup.add(
        'dropdownlist',
        undefined,
        ['IN', 'ACC', 'OM', 'Home', 'FJ', 'FW']
    );

    // List dropdown
    brandListSubjectGroup.add('statictext', undefined, 'List:');
    var listDropdown = brandListSubjectGroup.add(
        'dropdownlist',
        undefined,
        ['A List', 'B List', 'C List', 'D List', 'E List']
    );
    listDropdown.selection = 0;

    // Subject field
    brandListSubjectGroup.add('statictext', undefined, 'Subject:');
    var docName = app.documents.length > 0 ? app.activeDocument.name.split('.')[0] : '';
    var subjectField = brandListSubjectGroup.add(
        'edittext',
        undefined,
        '',
        { name: 'subjectField' }
    );
    subjectField.preferredSize = [300, 20];

    // Body field
    var bodyField = dialog.add(
        'edittext',
        undefined,
        '',
        { multiline: true, scrollable: true }
    );
    bodyField.size = [400, 150];

    // Buttons
    var btnGroup = dialog.add('group');
    var createBtn = btnGroup.add('button', undefined, 'Create Email');
    var cancelBtn = btnGroup.add('button', undefined, 'Cancel');

    // Helpers
    function updateSubject() {
        var brandText = brandMRP.value ? 'MRP' : 'NAP';
        var typeText  = typeDropdown.selection ? typeDropdown.selection.text : '';
        var listText  = listDropdown.selection ? listDropdown.selection.text : '';
        subjectField.text =
            brandText +
            (typeText ? ' ' + typeText : '') +
            (listText ? ' ' + listText : '') +
            (docName ? ' - ' + docName : '') +
            ' - ';
    }

    function clearAllCheckboxes() {
        for (var cat in checkboxes) {
            for (var i = 0; i < checkboxes[cat].length; i++) {
                checkboxes[cat][i].checkbox.value = false;
                checkboxes[cat][i].ccCheckbox.value = false;
            }
        }
    }

    function autoSelectRecipients(brand, type) {
        clearAllCheckboxes();
        var searchTag = brand + '_' + type;
        for (var cat in checkboxes) {
            for (var k = 0; k < checkboxes[cat].length; k++) {
                var cb = checkboxes[cat][k].checkbox;
                var ccCb = checkboxes[cat][k].ccCheckbox;
                if (cb.brandTypeTags.indexOf(searchTag) !== -1) {
                    cb.value = true;
                    ccCb.value = true;
                }
            }
        }
    }

    // Events
    brandMRP.onClick = function() {
        updateSubject();
        var t = typeDropdown.selection ? typeDropdown.selection.text : null;
        if (t) autoSelectRecipients('MRP', t);
    };
    brandNAP.onClick = function() {
        updateSubject();
        var t = typeDropdown.selection ? typeDropdown.selection.text : null;
        if (t) autoSelectRecipients('NAP', t);
    };
    typeDropdown.onChange = function() {
        updateSubject();
        var brand = brandMRP.value ? 'MRP' : 'NAP';
        var t = typeDropdown.selection ? typeDropdown.selection.text : null;
        if (t) autoSelectRecipients(brand, t);
    };
    listDropdown.onChange = updateSubject;

    createBtn.onClick = function () {
        var toRecipients = [];
        var ccRecipients = [];

        for (var cat in checkboxes) {
            for (var k = 0; k < checkboxes[cat].length; k++) {
                var cb = checkboxes[cat][k].checkbox;
                var ccCb = checkboxes[cat][k].ccCheckbox;
                if (cb.value) {
                    var emails = cb.email;
                    if (typeof emails === "string") emails = [emails];

                    for (var i = 0; i < emails.length; i++) {
                        if (ccCb.value) {
                            ccRecipients.push(emails[i]);
                        } else {
                            toRecipients.push(emails[i]);
                        }
                    }
                }
            }
        }

        dialog.close(1);
        createEmailInOutlook(
            toRecipients,
            ccRecipients,
            subjectField.text,
            bodyField.text
        );
    };

    cancelBtn.onClick = function () {
        dialog.close();
    };

    // Init
    updateSubject();
    dialog.show();
}

function createEmailInOutlook(toRecipients, ccRecipients, subject, body) {
    function formatRecipients(recipientList, type) {
        var lines = "";
        for (var i = 0; i < recipientList.length; i++) {
            lines +=
                '        make new recipient with properties ' +
                '{type:' + type + ' recipient type, email address:{address:"' +
                recipientList[i] + '"}}\n';
        }
        return lines;
    }

    var appleScript =
        'tell application "Microsoft Outlook"\n' +
        '    activate\n' +
        '    set newMessage to make new outgoing message at mail folder "Drafts" ' +
        'with properties {subject:"' + subject.replace(/"/g,'\\"') + '", ' +
        'content:"' + body.replace(/"/g,'\\"') + '"}\n' +
        '    tell newMessage\n' +
        formatRecipients(toRecipients, "to") +
        formatRecipients(ccRecipients, "cc") +
        '    end tell\n' +
        '    open newMessage\n' +
        'end tell\n';

    var osaFile = new File("~/temp_outlook_email.scpt");
    osaFile.open("w");
    osaFile.write(appleScript);
    osaFile.close();
    app.system('osascript "' + osaFile.fsName + '"');
    osaFile.remove();
}

// Run it
var emailList = loadEmailList();
if (emailList) {
    showEmailForm(emailList);
}
