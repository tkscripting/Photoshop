// Define thumbnail size
var thumbnailWidth = 75;
var thumbnailHeight = 100;

// Define the folder path using the absolute path
var folderPath = new Folder("/Users/knippingt/Library/CloudStorage/OneDrive-SharedLibraries-YOOXNET-A-PORTERGROUP/O365G-Ecommerce-Studio - US Files/Retouch/Actions & Scripts/Photoshop Scripts/Extra Scripts/Templates");

// Check if the folder exists
if (folderPath.exists) {
    // Define the names for the thumbnails in all rows
    var templateNamesRow1 = ["Index", "Accessory", "Fine Jewelry", "Candle"];
    var templateNamesRow2 = ["White Shirt", "Tom Ford", "Blue Shirt", "Cufflink"]; // Added Tom Ford
    var templateNamesRow3 = ["Fine Watch IN", "Fine Watch E2", "Plate"]; // New third row with watches and "Plate"

    // Create a dialog to display the files
    var dlg = new Window("dialog", "Templater");

    // Create a group for displaying the first row of thumbnails
    var thumbnailGroupRow1 = dlg.add("group");
    thumbnailGroupRow1.orientation = "row";  // Arrange thumbnails horizontally
    thumbnailGroupRow1.alignChildren = 'top'; // Align children to the top

    // Create a group for displaying the second row of thumbnails
    var thumbnailGroupRow2 = dlg.add("group");
    thumbnailGroupRow2.orientation = "row";  // Arrange thumbnails horizontally
    thumbnailGroupRow2.alignChildren = 'top'; // Align children to the top

    // Create a group for displaying the third row of thumbnails
    var thumbnailGroupRow3 = dlg.add("group");
    thumbnailGroupRow3.orientation = "row";  // Arrange thumbnails horizontally
    thumbnailGroupRow3.alignChildren = 'top'; // Align children to the top

    // Helper function to create thumbnails with a label under it
    function createThumbnailGroup(names, group) {
        for (var i = 0; i < names.length; i++) {
            var name = names[i];
            var imgPath = new File(folderPath.fsName + "/" + name + ".jpg");

            // Create a group for each thumbnail and label
            var imageGroup = group.add("group");
            imageGroup.orientation = "column";  // Stack image and label vertically
            imageGroup.alignChildren = 'center'; // Center the image and text

            if (imgPath.exists) {
                try {
                    // Create an image element to display the image
                    var img = imageGroup.add("image", undefined, imgPath.fsName);
                    img.size = [thumbnailWidth, thumbnailHeight];  // Set the size of the thumbnail

                    // Create a statictext element for the filename below the thumbnail
                    var fileNameText = imageGroup.add("statictext", undefined, name);
                    fileNameText.justify = "center";  // Center the filename text

                    // Add event listener for image click
                    (function(i) {
                        img.onClick = function() {
                            // Get the active document
                            var doc = app.activeDocument;
                            var originalDocName = doc.name.replace(/\.[^\.]+$/, ''); // Get the active document name without extension
                            var originalFile = new File(doc.fullName); // Get the full path of the original document

                            // Check the aspect ratio before proceeding
                            if (!isAspectRatio3to4(doc)) {
                                var result = confirm("The file is not in a 3:4 aspect ratio, would you like to crop it?", true, "File Aspect Ratio Issue");
                                if (result) {
                                    cropTo3to4(doc);
                                } else {
                                    dlg.close();
                                    return; // Close the dialog if the user chooses to not proceed
                                }
                            }

                            if (names[i] === "White Shirt" || names[i] === "Blue Shirt" || names[i] === "Tom Ford") { // Added Tom Ford
                                // Check if a selection exists
                                if (!hasSelection()) {
                                    // No selection, alert the user
                                    alert("Please make a selection before running");
                                    dlg.close(); // Close the dialog after alerting the user
                                } else {
                                    // A selection exists, copy it and open the .tif file
                                    var selection = doc.selection;
                                    selection.copy();

                                    // Rename and save the original document with '_old' suffix
                                    var originalOldFile = new File(originalFile.path + "/" + originalDocName + "_old.tif");
                                    originalFile.rename(originalOldFile.name); // Rename the original file

                                    // Open the corresponding White Shirt, Blue Shirt, or Tom Ford .tif file
                                    var templateFile = new File(folderPath.fsName + "/" + names[i] + ".tif");

                                    // Check if the template file exists
                                    if (templateFile.exists) {
                                        // Open the template file (White Shirt, Blue Shirt, or Tom Ford .tif)
                                        var templateDoc = open(templateFile);

                                        // Paste the copied selection into the template document
                                        templateDoc.paste();

                                        // Save the template as the original document name
                                        var savePath = new File(originalFile.path + "/" + originalDocName + ".tif");

                                        var saveOptions = new TiffSaveOptions();
                                        templateDoc.saveAs(savePath, saveOptions, true); // Save and overwrite if necessary

                                        // Close all files and reopen the original renamed file
                                        closeAllDocuments();
                                        open(savePath); // Open the newly saved file (the original name)

                                        // Close the dialog after saving and reopening
                                        dlg.close();
                                    } else {
                                        alert(names[i] + " template not found.");
                                    }
                                }
                            } else if (names[i] === "Fine Watch IN" || names[i] === "Fine Watch E2") {
                                // For watches: open the TIF file, duplicate the smart object layer, close the file, and then scale the layer to fit the canvas
                                var templateFile = new File(folderPath.fsName + "/" + names[i] + ".tif");
                                if (templateFile.exists) {
                                    var templateDoc = open(templateFile);
                                    // Duplicate the active layer (assumed smart object) into the current active document
                                    var newLayer = templateDoc.activeLayer.duplicate(doc, ElementPlacement.PLACEATBEGINNING);
                                    templateDoc.close(SaveOptions.DONOTSAVECHANGES);
                                    FitLayerToCanvas(false);
                                    dlg.close();
                                } else {
                                    alert(names[i] + " template not found.");
                                }
                            } else if (names[i] === "Plate") {
                                // For Plate: simply place the corresponding PNG file and scale it to fit the canvas
                                var fileToPlace = new File(folderPath.fsName + "/" + names[i] + ".png");
                                if (fileToPlace.exists) {
                                    var placedLayer = doc.artLayers.add();
                                    doc.activeLayer = placedLayer;
                                    var tempDoc = open(fileToPlace);
                                    tempDoc.activeLayer.copy();
                                    tempDoc.close(SaveOptions.DONOTSAVECHANGES);
                                    doc.paste();
                                    FitLayerToCanvas(false);
                                }
                                dlg.close();
                            } else {
                                // For other templates, place the corresponding PNG image
                                var fileToPlace = new File(folderPath.fsName + "/" + names[i] + ".png");
                                if (fileToPlace.exists) {
                                    var placedLayer = doc.artLayers.add();
                                    doc.activeLayer = placedLayer;
                                    var tempDoc = open(fileToPlace);
                                    tempDoc.activeLayer.copy();
                                    tempDoc.close(SaveOptions.DONOTSAVECHANGES);
                                    doc.paste();
                                    FitLayerToCanvas(false);
                                }
                                dlg.close();
                            }
                        };

                        fileNameText.onClick = img.onClick; // Same functionality for clicking text
                    })(i); // Pass the index of the file
                } catch (e) {
                    // Handle invalid image data error
                    continue; // Skip to the next file if the image loading fails
                }
            }
        }
    }

    // Function to check if the aspect ratio is 3:4
    function isAspectRatio3to4(doc) {
        var width = doc.width.as('px');
        var height = doc.height.as('px');
        var aspectRatio = width / height;
        return Math.abs(aspectRatio - (3 / 4)) < 0.01; // Allowing a small margin of error
    }

    // Function to crop the document to a 3:4 ratio
    function cropTo3to4(doc) {
        var width = doc.width.as('px');
        var height = doc.height.as('px');
        
        // Calculate the new dimensions for 3:4 aspect ratio
        var newWidth = width;
        var newHeight = height;

        if (width / height > 3 / 4) {
            // Crop width if width is too large
            newWidth = height * (3 / 4);
        } else {
            // Crop height if height is too large
            newHeight = width * (4 / 3);
        }

        // Calculate the center position for cropping
        var cropLeft = (width - newWidth) / 2;
        var cropTop = (height - newHeight) / 2;

        // Perform the cropping
        doc.crop([cropLeft, cropTop, cropLeft + newWidth, cropTop + newHeight]);
    }

    // Function to check if there is an active selection in the document
    function hasSelection() {
        try {
            var doc = app.activeDocument;
            var selection = doc.selection;
            // Check if the selection has non-zero width/height
            return selection.bounds[0] !== selection.bounds[2] && selection.bounds[1] !== selection.bounds[3];
        } catch (e) {
            return false; // In case of any error, treat as no selection
        }
    }

    // Function to fit the layer to the canvas (optional)
    function FitLayerToCanvas(keepAspect) {
        var doc = app.activeDocument;
        var layer = doc.activeLayer;

        // Do nothing if the layer is background or locked
        if (layer.isBackgroundLayer || layer.allLocked || layer.pixelsLocked || layer.positionLocked || layer.transparentPixelsLocked) return;

        // Do nothing if the layer is not a normal artLayer or Smart Object
        if (layer.kind != LayerKind.NORMAL && layer.kind != LayerKind.SMARTOBJECT) return;

        // Store the ruler units
        var defaultRulerUnits = app.preferences.rulerUnits;
        app.preferences.rulerUnits = Units.PIXELS;

        var width = doc.width.as('px');
        var height = doc.height.as('px');
        var bounds = layer.bounds;
        var layerWidth = bounds[2].as('px') - bounds[0].as('px');
        var layerHeight = bounds[3].as('px') - bounds[1].as('px');

        // Move the layer so the top-left corner matches the canvas top-left corner
        layer.translate(new UnitValue(0 - layer.bounds[0].as('px'), 'px'), new UnitValue(0 - layer.bounds[1].as('px'), 'px'));

        if (!keepAspect) {
            // Scale the layer to match the canvas
            layer.resize((width / layerWidth) * 100, (height / layerHeight) * 100, AnchorPosition.TOPLEFT);
        } else {
            // Maintain the aspect ratio
            var layerRatio = layerWidth / layerHeight;
            var newWidth = width;
            var newHeight = ((1.0 * width) / layerRatio);
            if (newHeight >= height) {
                newWidth = layerRatio * height;
                newHeight = height;
            }
            var resizePercent = newWidth / layerWidth * 100;
            layer.resize(resizePercent, resizePercent, AnchorPosition.TOPLEFT);
        }

        // Restore the ruler units
        app.preferences.rulerUnits = defaultRulerUnits;
    }

    // Function to close all open documents without saving changes
    function closeAllDocuments() {
        var docs = app.documents;
        for (var i = docs.length - 1; i >= 0; i--) {
            var doc = docs[i];
            doc.close(SaveOptions.DONOTSAVECHANGES); // Close without saving any documents
        }
    }

    // Create thumbnail groups for the predefined names in all rows
    createThumbnailGroup(templateNamesRow1, thumbnailGroupRow1);
    createThumbnailGroup(templateNamesRow2, thumbnailGroupRow2);
    createThumbnailGroup(templateNamesRow3, thumbnailGroupRow3); // Add the new third row with watches and Plate

    // Add a Cancel button to the dialog
    var cancelButton = dlg.add("button", undefined, "Cancel");

    // Add the cancel button functionality (close the dialog without doing anything)
    cancelButton.onClick = function() {
        dlg.close();
    };

    // Show the dialog box
    dlg.show();

} else {
    alert("The specified folder does not exist: " + folderPath.fsName);
}
