const fs = require("fs");
const path = require("path");
const convert = require("heic-convert");

const assetsDir = path.join(__dirname, "assets");

const images = [
  "IMG_4936.HEIC",
  "IMG_4947.HEIC",
  "IMG_4946.HEIC",
  "IMG_4945.HEIC"
];

async function convertImages() {
  for (const file of images) {
    const inputPath = path.join(assetsDir, file);
    const outputPath = path.join(
      assetsDir,
      file.replace(/\.heic$/i, ".png")
    );

    if (!fs.existsSync(inputPath)) {
      throw new Error(`Image not found: ${inputPath}`);
    }

    const inputBuffer = fs.readFileSync(inputPath);

    const outputBuffer = await convert({
      buffer: inputBuffer,
      format: "PNG",
      quality: 1
    });

    fs.writeFileSync(outputPath, outputBuffer);
    console.log(`Converted: ${file}`);
  }

  console.log("All images converted successfully!");
}

convertImages().catch(error => {
  console.error("Image conversion failed:", error);
  process.exit(1);
});
