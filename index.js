const pptxgen = require("pptxgenjs");

const pptx = new pptxgen();

// إعدادات العرض
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Sports Nutrition Seminar";
pptx.subject = "Sports Nutrition and Athletic Performance";
pptx.title = "التغذية الرياضية والأداء الرياضي";
pptx.company = "جامعة صبراتة";
pptx.lang = "ar-LY";
pptx.theme = {
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  lang: "ar-LY"
};

// الهوية البصرية
const COLORS = {
  navy: "10263F",
  blue: "087EA4",
  cyan: "35B6D6",
  dark: "244B57",
  white: "FFFFFF",
  pale: "F4FBFD",
  light: "E5F5F9",
  gray: "647987",
  gold: "D6B46A"
};

// مقاسات الشريحة
const SW = 13.333;
const SH = 7.5;

// خلفية موحدة قابلة للتطوير
function addBackground(slide, color = COLORS.pale) {
  slide.background = { color };

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: SW,
    h: 0.12,
    line: { color: COLORS.cyan, transparency: 100 },
    fill: { color: COLORS.cyan }
  });
}

// عنوان الشريحة
function addHeading(slide, title, subtitle = "") {
  slide.addText(title, {
    x: 0.65,
    y: 0.45,
    w: 12,
    h: 0.65,
    fontFace: "Arial",
    fontSize: 27,
    bold: true,
    color: COLORS.navy,
    align: "right",
    margin: 0
  });

  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.7,
      y: 1.15,
      w: 11.9,
      h: 0.35,
      fontFace: "Arial",
      fontSize: 12,
      color: COLORS.gray,
      align: "right",
      margin: 0
    });
  }
}

// ترقيم الشرائح
function addFooter(slide, number) {
  slide.addShape(pptx.ShapeType.line, {
    x: 0.65,
    y: 7.08,
    w: 12,
    h: 0,
    line: { color: COLORS.light, width: 1.2 }
  });

  slide.addText("التغذية الرياضية والأداء الرياضي", {
    x: 0.7,
    y: 7.15,
    w: 8,
    h: 0.18,
    fontFace: "Arial",
    fontSize: 8,
    color: COLORS.gray,
    margin: 0
  });

  slide.addText(String(number), {
    x: 11.9,
    y: 7.12,
    w: 0.6,
    h: 0.22,
    fontFace: "Arial",
    fontSize: 10,
    bold: true,
    color: COLORS.blue,
    align: "center",
    margin: 0
  });
}

// اختبار أول شريحة
const slide = pptx.addSlide();

addBackground(slide);
addHeading(
  slide,
  "التغذية الرياضية والأداء الرياضي",
  "Sports Nutrition and Athletic Performance"
);

slide.addText("هوية بصرية أكاديمية", {
  x: 2,
  y: 3,
  w: 9.3,
  h: 0.8,
  fontFace: "Arial",
  fontSize: 30,
  bold: true,
  color: COLORS.blue,
  align: "center",
  margin: 0
});

addFooter(slide, 1);

// حفظ العرض
pptx.writeFile({ fileName: "Sports-Nutrition.pptx" });
