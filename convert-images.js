const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

const assetsDir = path.join(__dirname, "assets");

const images = [
  "IMG4936.HEIC",
  "IMG4947.HEIC",
  "IMG4946.HEIC",
  "IMG4945.HEIC"
];

async function convertImages() {
  for (const file of images) {
    const input = path.join(assetsDir, file);
    const output = path.join(
      assetsDir,
      file.replace(".HEIC", ".png")
    );

    if (!fs.existsSync(input)) {
      throw new Error(`Image not found: ${file}`);
    }

    await sharp(input)
      .rotate()
      .png()
      .toFile(output);

    console.log(`Converted: ${file}`);
  }

  console.log("All images converted successfully.");
}

convertImages().catch((error) => {
  console.error(error);
  process.exit(1);
});
