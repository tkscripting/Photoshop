// Load the save_path.jsx file to get the save path
var savePathFile = new File("/Users/Shared/Extra Scripts/save_path.jsx");
if (savePathFile.exists) {
  $.evalFile(savePathFile);
} else {
  alert("Save path file not found!");
  exit();
}

if (!savePath) {
  alert("No savePath variable found in save_path.jsx");
  exit();
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
  app.displayDialogs = DialogModes.NO;

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

    // ✅ Capture original document path before creating new documents
    var originalDocPath = app.activeDocument.path;

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

    if (selectionHeight < 1 || selectionWidth < 1) {
      alert('Swatcher Error\nInvalid selection. Please select a visible area.');
      return;
    }    

    var shotType = activeDocument.name.split('_').pop();
    var swatchFileName = activeDocument.name.replace(shotType, 'sw.jpg');

    var saveFolder = new Folder(savePath);
    if (!saveFolder.exists) saveFolder.create();
    var swatchPath = new File(saveFolder + '/' + swatchFileName);

    // ✅ New: define second output path in the same folder as original document
    var swatchPathOriginal = new File(originalDocPath + '/' + swatchFileName);

    const copyMerge = activeDocument.artLayers.length > 1;
    activeDocument.artLayers.add();
    mergeVisible();

    activeDocument.selection.copy(copyMerge);
    activeDocument.selection.deselect();
    activeDocument.activeLayer.remove();

    var tempDoc = app.documents.add(selectionWidth, selectionHeight, 72, "TempAverage", NewDocumentMode.RGB);
    tempDoc.paste();
    tempDoc.resizeImage(1, 1);
    tempDoc.flatten();

    var colorSample = tempDoc.colorSamplers.add([0.5, 0.5]);
    var avgColor = colorSample.color;
    colorSample.remove();
    tempDoc.close(SaveOptions.DONOTSAVECHANGES);

    try {
      app.runMenuItem(stringIDToTypeID('closeInfoPanel'));
    } catch (err) {}

    var swatchDoc = app.documents.add(72, 72, 72, "Swatch", NewDocumentMode.RGB);
    var solidLayer = swatchDoc.artLayers.add();
    swatchDoc.selection.selectAll();
    var fillColor = new SolidColor();
    fillColor.rgb.red = avgColor.rgb.red;
    fillColor.rgb.green = avgColor.rgb.green;
    fillColor.rgb.blue = avgColor.rgb.blue;
    swatchDoc.selection.fill(fillColor);
    swatchDoc.selection.deselect();

    var jpgSaveOptions = new JPEGSaveOptions();
    jpgSaveOptions.embedColorProfile = true;
    jpgSaveOptions.formatOptions = FormatOptions.STANDARDBASELINE;
    jpgSaveOptions.matte = MatteType.NONE;
    jpgSaveOptions.quality = 10;

    try {
      // ✅ Save to both locations
      swatchDoc.saveAs(swatchPath, jpgSaveOptions, true, Extension.LOWERCASE);
      swatchDoc.saveAs(swatchPathOriginal, jpgSaveOptions, true, Extension.LOWERCASE);
      swatchDoc.close(SaveOptions.DONOTSAVECHANGES);

      var swatchNumber = swatchFileName.split('_')[0];
      app.system("osascript -e 'display notification \"Created swatch for " + swatchNumber + "\" with title \"Averager\"'");
    } catch (err) {
      alert('Swatcher Error\nUnable to create swatch:\n\n' + err);
    }

    app.preferences.rulerUnits = startRulerUnits;
    app.preferences.typeUnits = startTypeUnits;
  } finally {
    app.displayDialogs = originalDialogSetting;
  }
})();
