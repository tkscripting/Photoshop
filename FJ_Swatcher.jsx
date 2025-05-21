#target photoshop

function showFJSwatchesDialog() {
    var dialog = new Window('dialog', 'FJ Swatches');
    var folderPath = "/Users/knippingt/Library/CloudStorage/OneDrive-SharedLibraries-YOOXNET-A-PORTERGROUP/O365G-Ecommerce-Studio - US Files/Retouch/Actions & Scripts/Photoshop Scripts/Extra Scripts/Swatches";
    var folder = new Folder(folderPath);
    var documentsFolder = Folder.myDocuments;

    // Load the save_path.jsx file to get the save path
    var savePathFile = new File("/Users/Shared/Extra Scripts/save_path.jsx");
    if (savePathFile.exists) {
        $.evalFile(savePathFile); // This will load the savePath variable from save_path.jsx
    } else {
        alert("Save path file not found!");
        return;
    }

    // Ensure the savePath variable is defined in the loaded file
    if (!savePath) {
        alert("No savePath variable found in save_path.jsx");
        return;
    }

    if (folder.exists) {
        var files = folder.getFiles("*.jpg");

        var imageGroup = dialog.add('group');
        imageGroup.orientation = 'row';

        for (var i = 0; i < files.length; i++) {
            (function(file) {
                var itemGroup = imageGroup.add('group');
                itemGroup.orientation = 'column';
                itemGroup.alignChildren = 'center';

                // Load the image as a thumbnail
                var thumbnailPath = file.fsName; // Absolute path to file
                var thumbnailImage = itemGroup.add('image', undefined, File(thumbnailPath));
                thumbnailImage.size = [100, 100];

                var baseName = file.name.replace(".jpg", "").replace(/%20/g, " ");
                itemGroup.add('statictext', undefined, baseName);

                thumbnailImage.onClick = function() {
                    if (app.documents.length === 0) {
                        alert("No document is currently open.");
                        return;
                    }

                    var activeDoc = app.activeDocument;

                    // Get the original document name (remove extension)
                    var originalName = activeDoc.name.replace(/\.[^\.]+$/, "");

                    // Remove last 2 characters and append sw.jpg
                    var newName = originalName.slice(0, -2) + "sw.jpg";

                    // Define the save folder using the savePath variable
                    var saveFolder = new Folder(savePath);
                    if (!saveFolder.exists) saveFolder.create(); // Create the folder if it doesn't exist
                    var swatchPath = new File(saveFolder + '/' + newName);

                    // Copy the clicked image to the new location with the desired name
                    file.copy(swatchPath);

                    dialog.close(); // Close the popup
                };
            })(files[i]);
        }

        dialog.show();
    } else {
        alert("The folder path does not exist.");
    }
}

showFJSwatchesDialog();
