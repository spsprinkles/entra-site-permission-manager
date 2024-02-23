const archiver = require("archiver");
const fs = require("fs");
const path = require("path");

// Read the flows directory
const srcDir = fs.readdirSync(path.join(__dirname, "flows"));
srcDir.forEach((dirName) => {
    // Determine the destination file
    const dstFile = path.join(__dirname, "dist/" + dirName + ".zip");
    const dir = srcDir[dirName];
    console.log(dstFile);

    // See if it exists
    if (fs.existsSync(dstFile)) {
        // Delete the file
        fs.unlinkSync(dstFile);

        // Log
        console.log("Deleted the flow package: " + dirName);
    }

    // Create the archive instance
    const archive = archiver("zip");
    archive.on("close", () => {
        console.log(archive.pointer() + " total bytes");
    });
    archive.on("error", (err) => {
        throw err;
    });

    // Create the zip file
    const oStream = fs.createWriteStream(dstFile);
    archive.pipe(oStream);

    // Archive the directory
    archive.directory(dirName + "/", false);
    archive.finalize();
});