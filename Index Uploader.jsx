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

    var doc = app.activeDocument;

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
            doc.close(SaveOptions.DONOTSAVECHANGES);

            // Format file name for the notification
            var match = fileName.match(/^(\d+)_([a-z]{2})$/i);
            var formattedFileName = "";
            if (match) {
                formattedFileName = match[1] + " " + match[2].toUpperCase();
            }

            // Send macOS notification
            app.system("osascript -e 'display notification \"Uploaded " + formattedFileName + "\" with title \"Madame Export\"'");

        } else {
            var msg = "The active file was not exported:\nMinimum Size - 2000×3000\nRequired Ratio - 3/4 - .75\n\n";
            var cleanName = doc.name.replace(/\.[^\.]+$/, '');
            var match = cleanName.match(/^(\d+)_([a-z]{2})$/i);
            if (match) cleanName = match[1] + "_" + match[2].toUpperCase();

            msg += cleanName + ":\n";

            for (var j = 0; j < issues.length; j++) {
                var issue = issues[j];
                if (issue.indexOf("too small") === 0) {
                    var sizeMatch = issue.match(/\((\d+)×(\d+)\)/);
                    if (sizeMatch) {
                        msg += "Too small - " + sizeMatch[1] + "×" + sizeMatch[2] + "\n";
                    }
                }
                if (issue.indexOf("wrong aspect ratio") === 0) {
                    var ratioMatch = issue.match(/\(([\d.]+)\)/);
                    if (ratioMatch) {
                        msg += "Wrong aspect ratio - " + ratioMatch[1] + "\n";
                    }
                }
            }

            alert(msg);
        }

    } catch (err) {
        alert("Export Error in document: " + doc.name + "\n\n" + err);
        try { doc.close(SaveOptions.DONOTSAVECHANGES); } catch (e) {}
    }
}
