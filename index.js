const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const pptx = new pptxgen();

pptx.layout = "LAYOUT_WIDE";
pptx.author = "جامعة صبراتة";
pptx.subject = "التغذية الرياضية والأداء الرياضي";
pptx.title = "التغذية الرياضية والأداء الرياضي";
pptx.lang = "ar-LY";

pptx.theme = {
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  lang: "ar-LY"
};

const C = {
  navy: "10263F",
  blue: "087EA4",
  cyan: "35B6D6",
  white: "FFFFFF",
  pale: "F4FBFD",
  gold: "D6B46A",
  gray: "C7DCE5"
};

const SW = 13.333;
const SH = 7.5;
const assets = path.join(__dirname, "assets");

function addText(slide, text, x, y, w, h, options = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: "Arial",
    color: C.white,
    margin: 0.05,
    breakLine: false,
    valign: "mid",
    ...options
  });
}

function addImageIfExists(slide, filename, x, y, w, h) {
  const imagePath = path.join(assets, filename);

  if (fs.existsSync(imagePath)) {
    slide.addImage({
      path: imagePath,
      x, y, w, h
    });
  } else {
    console.warn("Image not found:", filename);
  }
}

// ==============================
// الشريحة الأولى: الغلاف الرئيسي
// ==============================

const slide = pptx.addSlide();

slide.background = { color: C.navy };

// خلفية زخرفية
slide.addShape(pptx.ShapeType.rect, {
  x: 0,
  y: 0,
  w: SW,
  h: 0.12,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

slide.addShape(pptx.ShapeType.ellipse, {
  x: 9.1,
  y: 0.5,
  w: 4.7,
  h: 6.5,
  line: { color: C.blue, transparency: 100 },
  fill: { color: C.blue, transparency: 75 }
});

slide.addShape(pptx.ShapeType.rect, {
  x: 0,
  y: 6.95,
  w: SW,
  h: 0.55,
  line: { color: C.blue, transparency: 100 },
  fill: { color: C.blue, transparency: 65 }
});

// الشعارات الرسمية
const logos = [
  ["IMG_4947.png", 0.8],
  ["IMG_4946.png", 2.25],
  ["IMG_4945.png", 3.7]
];

for (const [filename, x] of logos) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x,
    y: 0.45,
    w: 1.05,
    h: 1.05,
    line: { color: C.white, width: 1.5 },
    fill: { color: C.white }
  });

  addImageIfExists(
    slide,
    filename,
    x + 0.1,
    0.55,
    0.85,
    0.85
  );
}

// اسم الجامعة والجهة
addText(
  slide,
  "دولة ليبيا",
  0.6, 1.75, 6.8, 0.35,
  {
    fontSize: 17,
    bold: true,
    color: C.gold,
    align: "center"
  }
);

addText(
  slide,
  "جامعة صبراتة",
  0.6, 2.12, 6.8, 0.45,
  {
    fontSize: 23,
    bold: true,
    align: "center"
  }
);

addText(
  slide,
  "الدورة التدريبية لمعلمي ومعلمات التربية البدنية بالمنطقة الغربية",
  0.6, 2.65, 6.8, 0.65,
  {
    fontSize: 14,
    color: C.gray,
    align: "center"
  }
);

// الخط الزخرفي
slide.addShape(pptx.ShapeType.line, {
  x: 1.6,
  y: 3.45,
  w: 4.8,
  h: 0,
  line: { color: C.gold, width: 2 }
});

// عنوان الندوة
addText(
  slide,
  "التغذية الرياضية",
  0.5, 3.75, 7.1, 0.75,
  {
    fontSize: 35,
    bold: true,
    color: C.white,
    align: "center"
  }
);

addText(
  slide,
  "والأداء الرياضي",
  0.5, 4.5, 7.1, 0.7,
  {
    fontSize: 32,
    bold: true,
    color: C.cyan,
    align: "center"
  }
);

addText(
  slide,
  "Sports Nutrition and Athletic Performance",
  0.5, 5.25, 7.1, 0.4,
  {
    fontSize: 16,
    color: C.gray,
    align: "center"
  }
);
// صورة الدكتورة
slide.addShape(pptx.ShapeType.roundRect, {
  x: 8.45,
  y: 1.25,
  w: 3.85,
  h: 4.45,
  rectRadius: 0.12,
  line: { color: C.gold, width: 2.5 },
  fill: { color: C.white }
});

addImageIfExists(
  slide,
  "IMG_4936.png",
  8.55,
  1.35,
  3.65,
  4.25
);
  rectRadius: 0.12,
  line: { color: C.gold, width: 2.5 },
  fill: { color: C.white }
});

addImageIfExists(
  slide,
 "IMG_4936.png"
  8.55,
  1.35,
  3.65,
  4.25
);

// بيانات مقدمة الندوة
addText(
  slide,
  "إعداد وتقديم",
  7.9, 5.85, 4.9, 0.3,
  {
    fontSize: 14,
    color: C.gold,
    bold: true,
    align: "center"
  }
);

addText(
  slide,
  "الدكتورة مفيدة المختار راشد",
  7.5, 6.15, 5.6, 0.4,
  {
    fontSize: 19,
    bold: true,
    align: "center"
  }
);

addText(
  slide,
  "رئيسة قسم العلاج الطبيعي والتأهيل الحركي – الجميل",
  7.4, 6.55, 5.8, 0.25,
  {
    fontSize: 10,
    color: C.gray,
    align: "center"
  }
);

// سنة الندوة
addText(
  slide,
  "ندوة علمية | 2026",
  0.6, 7.08, 4, 0.22,
  {
    fontSize: 10,
    color: C.white,
    align: "left"
  }
);

// إنشاء ملف PowerPoint
async function buildPresentation() {
  await pptx.writeFile({
    fileName: "Sports-Nutrition.pptx"
  });

  console.log("Presentation created successfully!");
}

buildPresentation().catch(console.error);
