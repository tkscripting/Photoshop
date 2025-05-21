// Load the save_path.jsx file to get the save path
var savePathFile = new File("/Users/Shared/Extra Scripts/save_path.jsx"); // Path to your save_path.jsx file
if (savePathFile.exists) {
    $.evalFile(savePathFile); // This will load the savePath variable from save_path.jsx
} else {
    alert("Save path file not found!");
    exit(); // Use exit() to stop the script execution if the file is not found
}

// Ensure the savePath variable is defined in the loaded file
if (!savePath) {
    alert("No savePath variable found in save_path.jsx");
    exit(); // Use exit() to stop the script execution if savePath is not defined
}

if (app.documents.length > 0) {
    // Use the dynamically loaded savePath for the export folder
    var exportFolder = new Folder(savePath); // Use the value of savePath from save_path.jsx

    if (!exportFolder.exists) {
        exportFolder.create(); // Create the folder if it doesn't exist
    }

    var docs = [];
    for (var i = 0; i < app.documents.length; i++) {
        docs.push(app.documents[i]);
    }

    var savedFiles = {}; // { base: [suffix1, suffix2, ...] }
    var skippedFiles = []; // [{ name, issues[] }]

    for (var j = 0; j < docs.length; j++) {
        var doc = docs[j];
        app.activeDocument = doc;

        try {
            var width = doc.width.as("px");
            var height = doc.height.as("px");
            var aspectRatio = width / height;

            var issues = [];
            if (width < 2000 || height < 3000) {
                issues.push("too small (" + width + "×" + height + ")");
            }
            if (Math.abs(aspectRatio - 0.75) >= 0.01) {
                issues.push("wrong aspect ratio (" + aspectRatio.toFixed(2) + ")");
            }

            if (issues.length === 0) {
                doc.save(); // Save original

                var fileName = doc.name.replace(/\.[^\.]+$/, ''); // Strip extension
                var saveFile = new File(exportFolder + '/' + fileName + '.tif');

                var tiffOptions = new TiffSaveOptions();
                tiffOptions.imageCompression = TIFFEncoding.NONE;
                tiffOptions.layers = true;
                tiffOptions.embedColorProfile = true;
                tiffOptions.byteOrder = ByteOrder.IBM;

                doc.saveAs(saveFile, tiffOptions, true, Extension.LOWERCASE);

                // Track grouped filenames like 4637..._bk
                var match = fileName.match(/^(\d+)_([a-z]{2})$/i);
                if (match) {
                    var base = match[1];
                    var suffix = match[2].toUpperCase();
                    if (!savedFiles[base]) savedFiles[base] = [];
                    savedFiles[base].push(suffix);
                }

                // 🆕 Extra "pr" version if file ends in "in.tiff"
                if (/in\.tiff?$/i.test(doc.name)) {
                    var prFile = new File(exportFolder + '/' + fileName.replace(/in$/i, 'pr') + '.tif');

                    var plateLayer = null;
                    for (var k = 0; k < doc.layers.length; k++) {
                        if (doc.layers[k].name.toLowerCase() === "plate") {
                            plateLayer = doc.layers[k];
                            break;
                        }
                    }

                    var originalVisibility = null;
                    if (plateLayer) {
                        originalVisibility = plateLayer.visible;
                        plateLayer.visible = false;
                    }

                    doc.saveAs(prFile, tiffOptions, true, Extension.LOWERCASE);

                    if (plateLayer && originalVisibility !== null) {
                        plateLayer.visible = originalVisibility; // Restore visibility
                    }

                    // Track the "PR" suffix for notification
                    if (match) {
                        var base = match[1];
                        if (!savedFiles[base]) savedFiles[base] = [];
                        savedFiles[base].push("PR");
                    }
                }

                doc.close(SaveOptions.DONOTSAVECHANGES);

            } else {
                skippedFiles.push({ name: doc.name, issues: issues });
            }

        } catch (err) {
            alert("Export Error in document: " + doc.name + "\n\n" + err);
            try { doc.close(SaveOptions.DONOTSAVECHANGES); } catch (e) {}
        }
    }

    // 🔔 Send one big notification
    var messages = [];
    for (var base in savedFiles) {
        var parts = savedFiles[base];
        messages.push("Uploaded " + base + " " + parts.join(", ") + "\n");
    }

    if (messages.length > 0) {
        var finalMessage = messages.join("");
        app.system("osascript -e 'display notification \"" + finalMessage + "\" with title \"Madame Export\"'");
    }

    // 🚨 Show skipped files in formatted alert
    if (skippedFiles.length > 0) {
        var msg = "The following files were not exported:\nMinimum Size - 2000×3000\nRequired Ratio - 3/4 - .75\n\n";
        for (var i = 0; i < skippedFiles.length; i++) {
            var f = skippedFiles[i];
            var cleanName = f.name.replace(/\.[^\.]+$/, ''); // Remove .psd/.tif

            var match = cleanName.match(/^(\d+)_([a-z]{2})$/i);
            if (match) {
                cleanName = match[1] + "_" + match[2].toUpperCase();
            }

            msg += cleanName + ":\n";

            for (var j = 0; j < f.issues.length; j++) {
                var issue = f.issues[j];

                if (issue.indexOf("too small") === 0) {
                    var sizeMatch = issue.match(/\((\d+)×(\d+)\)/);
                    if (sizeMatch) {
                        var w = sizeMatch[1];
                        var h = sizeMatch[2];
                        msg += "Too small - " + w + "×" + h + "\n";
                    }
                }

                if (issue.indexOf("wrong aspect ratio") === 0) {
                    var ratioMatch = issue.match(/\(([\d.]+)\)/);
                    if (ratioMatch) {
                        var ratio = ratioMatch[1];
                        msg += "Wrong aspect ratio - " + ratio + "\n";
                    }
                }
            }

            msg += "\n";
        }

        alert(msg);
    }
}
