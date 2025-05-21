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

function mergeVisible() {
  var idMrgV = charIDToTypeID("MrgV");
  var desc6 = new ActionDescriptor();
  var idDplc = charIDToTypeID("Dplc");
  desc6.putBoolean(idDplc, true);
  executeAction(idMrgV, desc6, DialogModes.NO);
}

(function () {
  var originalDialogSetting = app.displayDialogs;
  app.displayDialogs = DialogModes.NO; // Disable dialogs temporarily

  try {
    if (!documents.length) {
      alert('Swatcher Error\nNo images open!');
      return;
    }

    var validFileName = /^(\d+)(_(mrp))?((_[a-z]{2})|(_[a-z]{1}[0-9]{1,2})|(_[a-z]{2}[0-9]{1})).tif$/;

    if (!validFileName.test(activeDocument.name)) {
      alert('Swatcher Error\nInvalid filename!');
      return;
    }
    
    var originalDocPath = app.activeDocument.path; // Capture original file path before we switch documents    

    try {
      activeDocument.selection.bounds;
    } catch (err) {
      alert('Swatcher Error\nPlease select an area.');
      return;
    }

    var startRulerUnits = app.preferences.rulerUnits;
    var startTypeUnits = app.preferences.typeUnits;
    app.preferences.rulerUnits = Units.PIXELS;
    app.preferences.typeUnits = TypeUnits.PIXELS;

    var selectionBounds = activeDocument.selection.bounds;
    var selectionHeight = selectionBounds[3] - selectionBounds[1];
    var selectionWidth = selectionBounds[2] - selectionBounds[0];

    if (selectionHeight < 72 || selectionWidth < 72) {
      alert('Swatcher Error\nPlease select a larger area. The minimum required height and width is 72px.');
      return;
    }
    if (selectionHeight !== selectionWidth) {
      alert('Swatcher Error\nPlease ensure your selection is a square.');
      return;
    }

    var shotType = activeDocument.name.split('_').pop();
    var swatchFileName = activeDocument.name.replace(shotType, 'sw.jpg');
    
    // Define the folder path using the `savePath` variable from the loaded file
    var saveFolder = new Folder(savePath);
    if (!saveFolder.exists) saveFolder.create(); // Create the folder if it doesn't exist
    var swatchPath = new File(saveFolder + '/' + swatchFileName);

    const copyMerge = activeDocument.artLayers.length > 1;
    activeDocument.artLayers.add();
    mergeVisible(); // Now it should work since mergeVisible is defined at the top

    activeDocument.selection.copy(copyMerge);
    activeDocument.selection.deselect();
    activeDocument.activeLayer.remove();

    documents.add(selectionHeight, selectionWidth);
    activeDocument.paste();
    activeDocument.resizeImage(72, 72);

    var jpgSaveOptions = new JPEGSaveOptions();
    jpgSaveOptions.embedColorProfile = true;
    jpgSaveOptions.formatOptions = FormatOptions.STANDARDBASELINE;
    jpgSaveOptions.matte = MatteType.NONE;
    jpgSaveOptions.quality = 10;

    try {
      // Save to main swatch path
      activeDocument.saveAs(swatchPath, jpgSaveOptions, true, Extension.LOWERCASE);
    
      // Also save to the original document folder
      var secondSavePath = new File(originalDocPath + '/' + swatchFileName);
      activeDocument.saveAs(secondSavePath, jpgSaveOptions, true, Extension.LOWERCASE);
    
      activeDocument.close(SaveOptions.DONOTSAVECHANGES);
    
      // Extract the numeric part from the filename for the notification
      var swatchNumber = swatchFileName.split('_')[0];
      app.system("osascript -e 'display notification \"Created swatch for " + swatchNumber + "\" with title \"Swatcher\"'");
    } catch (err) {
      alert('Swatcher Error\nUnable to create swatch:\n\n' + err);
    }
    

    app.preferences.rulerUnits = startRulerUnits;
    app.preferences.typeUnits = startTypeUnits;
  } finally {
    app.displayDialogs = originalDialogSetting; // Restore the dialog setting
  }
})();
