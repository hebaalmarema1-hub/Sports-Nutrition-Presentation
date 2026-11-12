const fs = require("fs");
const path = require("path");
const convert = require("heic-convert");

const assetsDir = path.join(__dirname, "assets");

const images = [
  "IMG4936.HEIC",
  "IMG4947.HEIC",
  "IMG4946.HEIC",
  "IMG4945.HEIC"
];

async function convertImages() {
  for (const file of images) {
    const inputPath = path.join(assetsDir, file);
    const outputPath = path.join(
      assetsDir,
      file.replace(/\.heic$/i, ".png")
    );

    if (!fs.existsSync(inputPath)) {
      throw new Error(`Image not found: ${file}`);
    }

    const inputBuffer = fs.readFileSync(inputPath);

    const outputBuffer = await convert({
      buffer: inputBuffer,
      format: "PNG",
      quality: 1
    });

    fs.writeFileSync(outputPath, outputBuffer);

    console.log(`Successfully converted: ${file}`);
  }

  console.log("All four images are ready!");
}

convertImages().catch((error) => {
  console.error("Image conversion failed:", error);
  process.exit(1);
});
