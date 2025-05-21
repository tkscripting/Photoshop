if (app.documents.length > 0) {
    var doc = app.activeDocument;

    try {
        var file = doc.fullName;
        var fileName = doc.name.replace(/\.[^\.]+$/, ''); // Strip extension

        // Check if file name ends in "in.tif" or "in.tiff"
        if (!/in\.tiff?$/i.test(doc.name)) {
            alert("Please run this on an index file");
        } else {
            var prName = fileName.replace(/in$/i, 'pr') + '.tif';
            var prFile = new File(file.parent + '/' + prName);

            // Find "plate" layer
            var plateLayer = null;
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name.toLowerCase() === "plate") {
                    plateLayer = doc.layers[i];
                    break;
                }
            }

            var originalVisibility = null;
            if (plateLayer) {
                originalVisibility = plateLayer.visible;
                plateLayer.visible = false;
            }

            // Save as .tif
            var tiffOptions = new TiffSaveOptions();
            tiffOptions.imageCompression = TIFFEncoding.NONE;
            tiffOptions.layers = true;
            tiffOptions.embedColorProfile = true;
            tiffOptions.byteOrder = ByteOrder.IBM;

            doc.saveAs(prFile, tiffOptions, true, Extension.LOWERCASE);

            if (plateLayer && originalVisibility !== null) {
                plateLayer.visible = originalVisibility;
            }
        }

    } catch (err) {
        alert("Something went wrong while exporting the PR version:\n\n" + err);
    }
}
