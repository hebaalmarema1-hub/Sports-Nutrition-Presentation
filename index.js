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

const SW = 13.333;
const SH = 7.5;
const assets = path.join(__dirname, "assets");

const C = {
  navy: "10263F",
  navy2: "173B55",
  blue: "087EA4",
  cyan: "35B6D6",
  aqua: "DDF4F8",
  pale: "F3F9FC",
  white: "FFFFFF",
  gold: "D6B46A",
  gray: "6E8492",
  light: "D7E6EC",
  green: "4B9B83",
  orange: "E5A34B",
  red: "C96C68",
  ink: "203747"
};

function txt(slide, text, x, y, w, h, opt = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: "Arial",
    fontSize: 18,
    color: C.ink,
    margin: 0.06,
    breakLine: false,
    valign: "mid",
    fit: "shrink",
    ...opt
  });
}

function shape(slide, type, x, y, w, h, fill, lineColor = fill, extra = {}) {
  slide.addShape(type, {
    x, y, w, h,
    fill: { color: fill },
    line: {
      color: lineColor,
      transparency: lineColor === fill ? 100 : 0,
      width: 1
    },
    ...extra
  });
}

function line(slide, x1, y1, x2, y2, color = C.light, width = 1.2, dash) {
  slide.addShape(pptx.ShapeType.line, {
    x: x1,
    y: y1,
    w: x2 - x1,
    h: y2 - y1,
    line: {
      color,
      width,
      ...(dash ? { dashType: dash } : {})
    }
  });
}

function imageIf(slide, name, x, y, w, h, sizing = "contain") {
  const p = path.join(assets, name);

  if (!fs.existsSync(p)) {
    console.warn("Image not found:", p);
    return;
  }

  slide.addImage({ path: p, x, y, w, h, sizing });
}

function topBand(slide, color = C.cyan) {
  shape(slide, pptx.ShapeType.rect, 0, 0, SW, 0.13, color);
}

function footer(slide, n, dark = false) {
  const color = dark ? "D9E8EF" : C.gray;

  line(slide, 0.55, 6.92, 12.78, 6.92,
    dark ? "426176" : C.light, 0.7);

  txt(slide,
    "التغذية الرياضية والأداء الرياضي  |  جامعة صبراتة  |  2026",
    0.65, 7.02, 8.4, 0.22,
    { fontSize: 9, color, align: "left" }
  );

  txt(slide, String(n).padStart(2, "0"),
    12.05, 6.99, 0.65, 0.28,
    {
      fontSize: 11,
      bold: true,
      color: dark ? C.gold : C.blue,
      align: "right"
    }
  );
}

function heading(slide, n, title, subtitle = "", dark = false) {
  const main = dark ? C.white : C.navy;
  const sub = dark ? "D6E5EC" : C.gray;

  txt(slide, String(n).padStart(2, "0"),
    0.62, 0.43, 0.75, 0.38,
    { fontSize: 15, bold: true, color: C.gold, align: "left" }
  );

  txt(slide, title, 1.55, 0.38, 11.1, 0.65,
    { fontSize: 28, bold: true, color: main, align: "right" }
  );

  shape(slide, pptx.ShapeType.rect,
    10.7, 1.15, 1.95, 0.045, C.gold);

  if (subtitle) {
    txt(slide, subtitle, 1.0, 1.35, 11.65, 0.55,
      { fontSize: 16, color: sub, align: "right" }
    );
  }
}

function dot(slide, x, y, r, color) {
  shape(slide, pptx.ShapeType.ellipse, x, y, r, r, color);
}

// ==============================
// 01 — الغلاف
// ==============================

{
  const s = pptx.addSlide();
  s.background = { color: C.navy };

  shape(s, pptx.ShapeType.rect, 0, 0, SW, 0.14, C.cyan);

  shape(s, pptx.ShapeType.ellipse,
    8.15, -0.65, 6.2, 7.7, C.navy2);

  shape(s, pptx.ShapeType.ellipse,
    -1.8, 5.7, 5.0, 3.0, C.blue, C.blue, {
      fill: { color: C.blue, transparency: 76 },
      line: { color: C.blue, transparency: 100 }
    }
  );

  const logos = [
    ["IMG_4947.png", 0.75],
    ["IMG_4946.png", 2.0],
    ["IMG_4945.png", 3.25]
  ];
  for (const [name, x] of logos) {
    shape(s, pptx.ShapeType.ellipse, x, 0.48, 0.9, 0.9, C.white, C.white);
    imageIf(s, name, x + 0.09, 0.57, 0.72, 0.72);
  }

  txt(s, "دولة ليبيا  |  جامعة صبراتة", 0.65, 1.58, 6.3, 0.38, {
    fontSize: 17, bold: true, color: C.gold, align: "center"
  });
  txt(s, "الدورة التدريبية لمعلمي ومعلمات التربية البدنية بالمنطقة الغربية",
    0.65, 2.02, 6.3, 0.58, { fontSize: 14, color: "D7E6EC", align: "center" });

  line(s, 1.25, 2.86, 6.35, 2.86, C.gold, 1.8);
  txt(s, "التغذية الرياضية", 0.55, 3.18, 6.7, 0.78, {
    fontSize: 35, bold: true, color: C.white, align: "center"
  });
  txt(s, "والأداء الرياضي", 0.55, 3.93, 6.7, 0.72, {
    fontSize: 32, bold: true, color: C.cyan, align: "center"
  });
  txt(s, "Sports Nutrition and Athletic Performance", 0.6, 4.72, 6.6, 0.38, {
    fontSize: 15, color: "D2E2E9", align: "center"
  });

  // صورة المحاضِرة في إطار أنيق
  shape(s, pptx.ShapeType.roundRect, 8.15, 0.82, 4.15, 4.78, C.white, C.gold, {
    line: { color: C.gold, width: 2.2 }, rectRadius: 0.16
  });
  imageIf(s, "IMG_4936.png", 8.27, 0.94, 3.91, 4.54);

  txt(s, "إعداد وتقديم", 7.7, 5.75, 5.05, 0.28, {
    fontSize: 13, bold: true, color: C.gold, align: "center"
  });
  txt(s, "الدكتورة مفيدة المختار راشد", 7.55, 6.02, 5.35, 0.38, {
    fontSize: 19, bold: true, color: C.white, align: "center"
  });
  txt(s, "رئيسة قسم العلاج الطبيعي والتأهيل الحركي – الجميل",
    7.45, 6.39, 5.55, 0.25, { fontSize: 10, color: "D5E3E9", align: "center" });
  txt(s, "ندوة علمية  |  2026", 0.65, 7.02, 3.2, 0.2, {
    fontSize: 10, color: C.white, align: "left"
  });
}

// 02 — مقدمة افتتاحية
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  topBand(s, C.gold);
  heading(s, 2, "مدخل إلى التغذية الرياضية",
    "الأداء الرياضي المتميز يبدأ من تكامل التدريب والتغذية والتعافي.", true);

  txt(s, "التدريب يطوّر القدرة… والتغذية تمدّ الجسم بما يحتاجه ليُظهر هذه القدرة.",
    1.0, 2.0, 11.4, 0.95, {
      fontSize: 25, bold: true, color: C.white, align: "center", margin: 0.12
    });

  const pillars = [
    { x: 1.05, num: "01", title: "الطاقة", desc: "وقود الحركة والجهد", col: C.cyan },
    { x: 4.9, num: "02", title: "الأداء", desc: "كفاءة وقوة وتحمل", col: C.gold },
    { x: 8.75, num: "03", title: "التعافي", desc: "إصلاح واستعادة", col: "75C5A8" }
  ];

  for (const p of pillars) {
    shape(s, pptx.ShapeType.roundRect, p.x, 3.65, 3.5, 1.65, C.navy2, "426176", {
      rectRadius: 0.15
    });
    dot(s, p.x + 1.42, 3.28, 0.66, p.col);
    txt(s, p.num, p.x + 1.45, 3.43, 0.6, 0.25, {
      fontSize: 13, bold: true, color: C.navy, align: "center"
    });
    txt(s, p.title, p.x + 0.2, 4.08, 3.1, 0.42, {
      fontSize: 22, bold: true, color: p.col, align: "center"
    });
    txt(s, p.desc, p.x + 0.2, 4.62, 3.1, 0.32, {
      fontSize: 14, color: C.white, align: "center"
    });
  }

  txt(s, "خلال هذه الندوة سنتعرّف كيف نختار الغذاء المناسب، ومتى نتناوله، وكيف ينعكس ذلك على الأداء.",
    1.0, 5.85, 11.4, 0.55, { fontSize: 17, color: "D7E6EC", align: "center" });
  footer(s, 2, true);
}

// 03 — مفهوم التغذية الرياضية
{
  const s = pptx.addSlide();
  s.background = { color: C.pale };
  topBand(s);
  heading(s, 3, "مفهوم التغذية الرياضية",
    "علم يهتم بتوفير احتياجات الجسم الغذائية لدعم النشاط البدني وتحسين الأداء.");

  // دائرة مركزية
  dot(s, 5.35, 2.35, 2.45, C.navy);
  txt(s, "التغذية", 5.65, 2.9, 1.85, 0.45, {
    fontSize: 27, bold: true, color: C.white, align: "center"
  });
  txt(s, "الرياضية", 5.65, 3.4, 1.85, 0.4, {
    fontSize: 22, color: C.cyan, align: "center"
  });

  const items = [
    { x: 0.9, y: 2.35, title: "توفير الطاقة", body: "تلبية احتياجات الجسم أثناء النشاط.", c: C.gold },
    { x: 9.0, y: 2.35, title: "دعم الأداء", body: "المساعدة على القوة والتحمل والتركيز.", c: C.cyan },
    { x: 0.9, y: 4.65, title: "تعزيز التعافي", body: "المساهمة في إصلاح الأنسجة واستعادة النشاط.", c: C.green },
    { x: 9.0, y: 4.65, title: "المحافظة على الصحة", body: "دعم وظائف الجسم والتوازن الغذائي.", c: C.blue }
  ];

  items.forEach(a => {
    shape(s, pptx.ShapeType.roundRect, a.x, a.y, 3.35, 1.25, C.white, C.light);
    shape(s, pptx.ShapeType.rect, a.x + 3.22, a.y + 0.16, 0.05, 0.92, a.c);
    txt(s, a.title, a.x + 0.18, a.y + 0.17, 2.9, 0.35, {
      fontSize: 18, bold: true, color: a.c, align: "right"
    });
    txt(s, a.body, a.x + 0.18, a.y + 0.62, 2.9, 0.42, {
      fontSize: 13, color: C.ink, align: "right"
    });
  });

  line(s, 4.25, 2.95, 5.4, 3.15, C.gold, 1.6);
  line(s, 7.75, 3.15, 9.0, 2.95, C.cyan, 1.6);
  line(s, 4.25, 5.05, 5.65, 4.45, C.green, 1.6);
  line(s, 7.55, 4.45, 9.0, 5.05, C.blue, 1.6);

  footer(s, 3);
}

// 04 — مصادر الطاقة
{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  topBand(s, C.gold);
  heading(s, 4, "مصادر الطاقة في الجسم",
    "يعتمد الجسم على المغذيات لإنتاج الطاقة اللازمة للحركة والجهد البدني.");

  const data = [
    { x: 0.9, title: "الكربوهيدرات", en: "Carbohydrates", body: "مصدر مهم للطاقة، خاصة أثناء الأنشطة متوسطة وعالية الشدة.", c: C.gold, pct: "طاقة سريعة ومتاحة" },
    { x: 4.95, title: "الدهون", en: "Fats", body: "مصدر مهم للطاقة، خاصة أثناء الراحة والأنشطة الطويلة منخفضة الشدة.", c: C.cyan, pct: "مخزون للطاقة" },
    { x: 9.0, title: "البروتين", en: "Protein", body: "يساهم في بناء الأنسجة وإصلاح العضلات، ويمكن استخدامه للطاقة عند الحاجة.", c: C.green, pct: "بناء وإصلاح" }
  ];

  data.forEach(a => {
    shape(s, pptx.ShapeType.roundRect, a.x, 2.15, 3.4, 3.65, C.pale, C.light);
    dot(s, a.x + 1.22, 2.48, 0.95, a.c);
    txt(s, a.title === "الكربوهيدرات" ? "C" : a.title === "الدهون" ? "F" : "P",
      a.x + 1.43, 2.72, 0.55, 0.4, {
        fontSize: 27, bold: true, color: C.white, align: "center"
      });
    txt(s, a.title, a.x + 0.15, 3.65, 3.1, 0.42, {
      fontSize: 21, bold: true, color: C.navy, align: "center"
    });
    txt(s, a.en, a.x + 0.15, 4.12, 3.1, 0.28, {
      fontSize: 12, color: C.gray, align: "center"
    });
    txt(s, a.body, a.x + 0.25, 4.58, 2.9, 0.62, {
      fontSize: 13, color: C.ink, align: "center", valign: "mid"
    });
    txt(s, a.pct, a.x + 0.2, 5.35, 3.0, 0.28, {
      fontSize: 13, bold: true, color: a.c, align: "center"
    });
  });

  txt(s, "تختلف احتياجات الجسم حسب نوع النشاط ومدته وشدته.",
    1.0, 6.15, 11.2, 0.35, {
      fontSize: 16, bold: true, color: C.blue, align: "center"
    });
  footer(s, 4);
}

// 05 — مكونات الوجبة المتوازنة
{
  const s = pptx.addSlide();
  s.background = { color: C.pale };
  topBand(s);
  heading(s, 5, "مكونات الوجبة الغذائية المتوازنة",
    "تنويع الطعام يساعد على توفير العناصر الغذائية التي يحتاجها الجسم.");

  // طبق غذائي توضيحي مقسم بصرياً
  dot(s, 1.0, 2.0, 3.9, C.light);
  dot(s, 1.14, 2.14, 3.62, C.white);
  shape(s, pptx.ShapeType.arc, 1.35, 2.35, 3.2, 3.2, C.green, C.green);
  shape(s, pptx.ShapeType.pie, 1.28, 2.28, 3.35, 3.35, C.green, C.green, {
    adjustPoint: 0.5
  });
  shape(s, pptx.ShapeType.rect, 2.96, 2.35, 0.04, 3.2, C.white);
  line(s, 1.4, 4.0, 4.5, 4.0, C.white, 4);

  txt(s, "½", 1.7, 3.05, 0.9, 0.6, {
    fontSize: 30, bold: true, color: C.white, align: "center"
  });
  txt(s, "¼", 3.25, 2.9, 0.75, 0.5, {
    fontSize: 25, bold: true, color: C.white, align: "center"
  });
  txt(s, "¼", 3.25, 4.35, 0.75, 0.5, {
    fontSize: 25, bold: true, color: C.white, align: "center"
  });

  const parts = [
    ["الخضروات والفواكه", "مصدر للفيتامينات والمعادن والألياف.", C.green],
    ["البروتين", "يدعم بناء العضلات وإصلاح الأنسجة.", C.cyan],
    ["الحبوب ومصادر النشويات", "تمد الجسم بالطاقة اللازمة للنشاط.", C.gold]
  ];

  parts.forEach((a, i) => {
    const y = 2.25 + i * 1.3;
    shape(s, pptx.ShapeType.roundRect, 5.65, y, 6.4, 0.95, C.white, C.light);
    dot(s, 11.55, y + 0.3, 0.28, a[2]);
    txt(s, a[0], 6.2, y + 0.12, 5.0, 0.3, {
      fontSize: 17, bold: true, color: a[2], align: "right"
    });
    txt(s, a[1], 6.0, y + 0.52, 5.2, 0.25, {
      fontSize: 12, color: C.ink, align: "right"
    });
  });

  txt(s, "هذا تقسيم إرشادي للطبق، وتختلف الكميات حسب احتياجات الفرد.",
    1.0, 6.25, 11.2, 0.3, {
      fontSize: 12, color: C.gray, align: "center"
    });
  footer(s, 5);
}

// 06 — أهمية الماء والترطيب
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  topBand(s, C.cyan);
  heading(s, 6, "أهمية الماء والترطيب",
    "الماء عنصر أساسي لتنظيم حرارة الجسم والمحافظة على كفاءة الأداء.", true);

  shape(s, pptx.ShapeType.teardrop, 1.2, 2.3, 2.6, 3.25, C.cyan, C.cyan, {
    rotate: 180
  });
  txt(s, "H₂O", 1.55, 3.48, 1.9, 0.65, {
    fontSize: 34, bold: true, color: C.white, align: "center"
  });

  const benefits = [
    ["تنظيم الحرارة", "يساعد على تبريد الجسم أثناء النشاط.", C.cyan],
    ["نقل العناصر", "يسهم في نقل المغذيات ودعم الدورة الدموية.", C.gold],
    ["كفاءة العضلات", "يدعم وظائف العضلات ويحافظ على الأداء.", "75C5A8"]
  ];

  benefits.forEach((a, i) => {
    const y = 2.25 + i * 1.25;
    shape(s, pptx.ShapeType.roundRect, 4.65, y, 7.45, 0.95, C.navy2, "426176");
    dot(s, 11.45, y + 0.28, 0.3, a[2]);
    txt(s, a[0], 7.0, y + 0.12, 4.1, 0.32, {
      fontSize: 19, bold: true, color: a[2], align: "right"
    });
    txt(s, a[1], 5.0, y + 0.52, 6.1, 0.25, {
      fontSize: 13, color: C.white, align: "right"
    });
  });

  txt(s, "لا تنتظر الشعور بالعطش حتى تبدأ بشرب الماء.",
    1.1, 6.05, 11.1, 0.45, {
      fontSize: 21, bold: true, color: C.gold, align: "center"
    });
  footer(s, 6, true);
}

// 07 — التوقيت الغذائي
{
  const s = pptx.addSlide();
  s.background = { color: C.pale };
  topBand(s);
  heading(s, 7, "التوقيت الغذائي حول التمرين",
    "اختيار الوقت المناسب للوجبات يساعد على الاستعداد والتعافي.");

  const stages = [
    {
      x: 0.9, title: "قبل التمرين", time: "قبل النشاط",
      body: "وجبة مناسبة تحتوي على الكربوهيدرات، مع وقت كافٍ للهضم.",
      c: C.gold, n: "01"
    },
    {
      x: 4.9, title: "أثناء التمرين", time: "خلال النشاط",
      body: "الماء بانتظام، وقد يحتاج النشاط الطويل إلى تعويض إضافي للطاقة.",
      c: C.cyan, n: "02"
    },
    {
      x: 8.9, title: "بعد التمرين", time: "مرحلة التعافي",
      body: "تعويض السوائل وتناول البروتين والكربوهيدرات حسب الحاجة.",
      c: C.green, n: "03"
    }
  ];

  line(s, 2.0, 3.0, 11.2, 3.0, C.light, 3);

  stages.forEach(a => {
    dot(s, a.x + 1.05, 2.52, 0.95, a.c);
    txt(s, a.n, a.x + 1.2, 2.78, 0.65, 0.3, {
      fontSize: 20, bold: true, color: C.white, align: "center"
    });
    shape(s, pptx.ShapeType.roundRect, a.x, 3.75, 3.45, 2.0, C.white, C.light);
    txt(s, a.title, a.x + 0.15, 4.0, 3.15, 0.42, {
      fontSize: 21, bold: true, color: a.c, align: "center"
    });
    txt(s, a.time, a.x + 0.15, 4.48, 3.15, 0.28, {
      fontSize: 13, bold: true, color: C.gray, align: "center"
    });
    txt(s, a.body, a.x + 0.25, 4.9, 2.95, 0.65, {
      fontSize: 14, color: C.ink, align: "center"
    });
  });

  footer(s, 7);
}

// 08 — التغذية والتعافي العضلي
{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  topBand(s, C.green);
  heading(s, 8, "التغذية والتعافي العضلي",
    "التعافي جزء من التدريب، وليس مرحلة منفصلة عنه.");

  const steps = [
    { x: 0.9, n: "1", title: "التمرين", body: "إجهاد العضلات واستهلاك الطاقة", c: C.blue },
    { x: 4.9, n: "2", title: "التغذية", body: "توفير البروتين والكربوهيدرات", c: C.cyan },
    { x: 8.9, n: "3", title: "الاستشفاء", body: "إصلاح الأنسجة واستعادة النشاط", c: C.green }
  ];

  steps.forEach((a, i) => {
    dot(s, a.x + 1.05, 2.35, 1.05, a.c);
    txt(s, a.n, a.x + 1.3, 2.62, 0.55, 0.4, {
      fontSize: 28, bold: true, color: C.white, align: "center"
    });
    shape(s, pptx.ShapeType.roundRect, a.x, 3.75, 3.45, 1.55, C.pale, C.light);
    txt(s, a.title, a.x + 0.15, 3.98, 3.15, 0.4, {
      fontSize: 21, bold: true, color: a.c, align: "center"
    });
    txt(s, a.body, a.x + 0.2, 4.55, 3.05, 0.45, {
      fontSize: 14, color: C.ink, align: "center"
    });
    if (i < 2) {
      txt(s, "←", a.x + 3.45, 2.7, 0.55, 0.5, {
        fontSize: 27, bold: true, color: C.gold, align: "center"
      });
    }
  });

  shape(s, pptx.ShapeType.roundRect, 2.0, 5.8, 9.35, 0.55, C.aqua, C.aqua);
  txt(s, "الغذاء المتوازن + السوائل + النوم الكافي = تعافٍ أفضل",
    2.2, 5.91, 8.95, 0.3, {
      fontSize: 17, bold: true, color: C.navy, align: "center"
    });

  footer(s, 8);
}
// 09 — الأخطاء الغذائية الشائعة
{
  const s = pptx.addSlide();
  s.background = { color: C.pale };
  topBand(s, C.orange);
  heading(s, 9, "أخطاء غذائية قد تؤثر في الأداء",
    "بعض العادات اليومية قد تقلل من كفاءة الجسم أثناء النشاط البدني.");

  const mistakes = [
    {
      n: "01",
      title: "تخطي الوجبات",
      body: "قد يؤدي إلى نقص الطاقة وصعوبة المحافظة على مستوى الأداء.",
      c: C.orange
    },
    {
      n: "02",
      title: "قلة شرب الماء",
      body: "قد تزيد من خطر الجفاف وتؤثر في التركيز والتحمل.",
      c: C.cyan
    },
    {
      n: "03",
      title: "الإفراط في المكملات",
      body: "استخدامها دون حاجة أو إرشاد مختص قد يسبب أضراراً.",
      c: C.red
    },
    {
      n: "04",
      title: "إهمال وجبة التعافي",
      body: "قد يؤخر تعويض السوائل والطاقة والعناصر الغذائية.",
      c: C.green
    }
  ];

  mistakes.forEach((a, i) => {
    const x = i % 2 === 0 ? 0.95 : 6.85;
    const y = i < 2 ? 2.2 : 4.25;

    shape(s, pptx.ShapeType.roundRect, x, y, 5.45, 1.55, C.white, C.light);
    dot(s, x + 4.65, y + 0.28, 0.48, a.c);

    txt(s, a.n, x + 4.72, y + 0.38, 0.34, 0.2, {
      fontSize: 10, bold: true, color: C.white, align: "center"
    });

    txt(s, a.title, x + 0.2, y + 0.25, 4.2, 0.35, {
      fontSize: 19, bold: true, color: a.c, align: "right"
    });

    txt(s, a.body, x + 0.2, y + 0.78, 4.8, 0.5, {
      fontSize: 14, color: C.ink, align: "right"
    });
  });

  footer(s, 9);
}
// 10 — العلاقة بين التغذية والأداء الرياضي
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  topBand(s, C.gold);
  heading(s, 10, "العلاقة بين التغذية والأداء الرياضي",
    "التغذية المناسبة تدعم قدرة الجسم على تقديم أفضل أداء ممكن.", true);

  const factors = [
    { x: 0.85, title: "الطاقة", value: "Energy", desc: "دعم النشاط البدني", c: C.gold },
    { x: 3.95, title: "القوة", value: "Strength", desc: "دعم عمل العضلات", c: C.cyan },
    { x: 7.05, title: "التحمل", value: "Endurance", desc: "المحافظة على الجهد", c: C.green },
    { x: 10.15, title: "التركيز", value: "Focus", desc: "دعم الانتباه والاستجابة", c: "A9A9F5" }
  ];

  factors.forEach((a, i) => {
    dot(s, a.x + 0.42, 2.45, 1.25, a.c);
    txt(s, String(i + 1).padStart(2, "0"),
      a.x + 0.7, 2.83, 0.7, 0.35, {
        fontSize: 22, bold: true, color: C.navy, align: "center"
      });

    txt(s, a.title, a.x, 4.0, 2.1, 0.4, {
      fontSize: 21, bold: true, color: a.c, align: "center"
    });
    txt(s, a.value, a.x, 4.43, 2.1, 0.25, {
      fontSize: 11, color: "D7E6EC", align: "center"
    });
    txt(s, a.desc, a.x, 4.92, 2.1, 0.45, {
      fontSize: 13, color: C.white, align: "center"
    });
  });

  line(s, 1.5, 5.75, 11.85, 5.75, "426176", 1.3);

  txt(s, "التغذية ليست بديلاً عن التدريب، بل شريك أساسي في تحسين نتائجه.",
    1.0, 6.05, 11.3, 0.45, {
      fontSize: 19, bold: true, color: C.gold, align: "center"
    });

  footer(s, 10, true);
}
// 11 — دور معلم التربية البدنية
{
  const s = pptx.addSlide();
  s.background = { color: C.pale };
  topBand(s, C.cyan);
  heading(s, 11, "دور معلم التربية البدنية في التوعية الغذائية",
    "المعلم شريك مهم في ترسيخ السلوكيات الصحية لدى الطلاب.");

  const roles = [
    {
      x: 0.85, y: 2.15,
      title: "تعزيز الوعي الغذائي",
      body: "توضيح أهمية الغذاء المتوازن وعلاقته بالنشاط البدني.",
      c: C.blue, n: "01"
    },
    {
      x: 6.85, y: 2.15,
      title: "تشجيع العادات الصحية",
      body: "تحفيز الطلاب على اختيار الأطعمة المفيدة والمتنوعة.",
      c: C.green, n: "02"
    },
    {
      x: 0.85, y: 4.25,
      title: "التوعية بالترطيب",
      body: "تذكير الطلاب بشرب الماء وملاحظة علامات الجفاف.",
      c: C.cyan, n: "03"
    },
    {
      x: 6.85, y: 4.25,
      title: "ملاحظة التعب والإرهاق",
      body: "الانتباه إلى علامات الإجهاد وتوجيه الطلاب للممارسات السليمة.",
      c: C.orange, n: "04"
    }
  ];

  roles.forEach(a => {
    shape(s, pptx.ShapeType.roundRect, a.x, a.y, 5.55, 1.55, C.white, C.light);
    shape(s, pptx.ShapeType.roundRect, a.x + 4.75, a.y + 0.25, 0.55, 0.55, a.c, a.c);
    txt(s, a.n, a.x + 4.82, a.y + 0.4, 0.4, 0.2, {
      fontSize: 11, bold: true, color: C.white, align: "center"
    });
    txt(s, a.title, a.x + 0.2, a.y + 0.22, 4.25, 0.4, {
      fontSize: 18, bold: true, color: a.c, align: "right"
    });
    txt(s, a.body, a.x + 0.2, a.y + 0.78, 4.8, 0.45, {
      fontSize: 13, color: C.ink, align: "right"
    });
  });

  footer(s, 11);
}
// 12 — وجبة غذائية متوازنة
{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  topBand(s, C.green);
  heading(s, 12, "مثال تطبيقي لوجبة غذائية متوازنة",
    "نموذج يجمع بين مصادر الطاقة والبروتين والخضروات.");

  // صورة طبق توضيحي
  dot(s, 1.05, 2.0, 4.15, C.light);
  dot(s, 1.2, 2.15, 3.85, C.pale);

  // أقسام الطبق
  shape(s, pptx.ShapeType.arc, 1.45, 2.4, 3.35, 3.35, C.green, C.green);
  shape(s, pptx.ShapeType.pie, 1.35, 2.3, 3.5, 3.5, C.gold, C.gold, {
    rotate: 90
  });
  line(s, 3.1, 2.4, 3.1, 5.65, C.white, 4);
  line(s, 1.45, 4.05, 4.75, 4.05, C.white, 4);

  txt(s, "خضروات", 1.65, 2.9, 1.2, 0.3, {
    fontSize: 14, bold: true, color: C.white, align: "center"
  });
  txt(s, "حبوب", 3.3, 2.9, 1.0, 0.3, {
    fontSize: 14, bold: true, color: C.navy, align: "center"
  });
  txt(s, "بروتين", 3.25, 4.5, 1.1, 0.3, {
    fontSize: 14, bold: true, color: C.navy, align: "center"
  });

  const meal = [
    {
      y: 2.15, title: "مصدر البروتين",
      body: "الدجاج أو السمك أو البيض أو البقوليات.",
      c: C.blue
    },
    {
      y: 3.25, title: "مصدر الكربوهيدرات",
      body: "الأرز أو البطاطا أو الخبز كامل الحبوب.",
      c: C.gold
    },
    {
      y: 4.35, title: "الخضروات والفواكه",
      body: "توفير الألياف والفيتامينات والمعادن.",
      c: C.green
    },
    {
      y: 5.45, title: "السوائل",
      body: "الماء للمحافظة على الترطيب.",
      c: C.cyan
    }
  ];

  meal.forEach(a => {
    shape(s, pptx.ShapeType.roundRect, 5.55, a.y, 6.55, 0.85, C.pale, C.light);
    dot(s, 11.55, a.y + 0.27, 0.25, a.c);
    txt(s, a.title, 6.0, a.y + 0.1, 5.15, 0.28, {
      fontSize: 16, bold: true, color: a.c, align: "right"
    });
    txt(s, a.body, 5.95, a.y + 0.48, 5.2, 0.22, {
      fontSize: 12, color: C.ink, align: "right"
    });
  });

  txt(s, "تُعدّل الكميات ومكونات الوجبة وفق العمر والحالة الصحية ومستوى النشاط.",
    1.0, 6.45, 11.2, 0.25, {
      fontSize: 11, color: C.gray, align: "center"
    });

  footer(s, 12);
}
// 13 — توصيات عملية
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  topBand(s, C.gold);
  heading(s, 13, "توصيات عملية لتحسين الأداء الرياضي",
    "خطوات بسيطة تصنع فرقاً في صحة الجسم وكفاءة النشاط.", true);

  const tips = [
    ["01", "نوّع غذاءك", "احرص على تناول وجبات متوازنة ومتنوعة.", C.cyan],
    ["02", "اهتم بالترطيب", "اشرب الماء بانتظام قبل النشاط وأثناءه وبعده.", C.gold],
    ["03", "نظّم وجباتك", "اختر توقيتاً مناسباً للوجبات حول التمرين.", C.green],
    ["04", "امنح جسمك الراحة", "اهتم بالنوم الكافي والتعافي بعد الجهد.", "A9A9F5"],
    ["05", "كن قدوة صحية", "شجّع الطلاب على اتباع عادات غذائية سليمة.", C.orange]
  ];

  tips.forEach((a, i) => {
    const y = 2.0 + i * 0.87;

    dot(s, 10.95, y + 0.08, 0.48, a[3]);

    txt(s, a[0], 11.03, y + 0.2, 0.3, 0.18, {
      fontSize: 10, bold: true, color: C.navy, align: "center"
    });

    txt(s, a[1], 7.5, y + 0.05, 3.0, 0.32, {
      fontSize: 17, bold: true, color: a[3], align: "right"
    });

    txt(s, a[2], 1.3, y + 0.08, 5.9, 0.3, {
      fontSize: 14, color: C.white, align: "right"
    });

    if (i < tips.length - 1) {
      line(s, 1.3, y + 0.7, 11.5, y + 0.7, "426176", 0.7);
    }
  });

  footer(s, 13, true);
}
// 14 — الخاتمة
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };

  shape(s, pptx.ShapeType.rect, 0, 0, SW, 0.14, C.gold);
  shape(s, pptx.ShapeType.ellipse, 8.8, -1.2, 6.5, 7.8, C.navy2);
  shape(s, pptx.ShapeType.ellipse, -2.2, 5.1, 5.2, 4.2, C.blue, C.blue, {
    fill: { color: C.blue, transparency: 75 },
    line: { color: C.blue, transparency: 100 }
  });

  txt(s, "التغذية الرياضية والأداء الرياضي",
    1.0, 1.55, 11.3, 0.7, {
      fontSize: 30, bold: true, color: C.white, align: "center"
    });

  line(s, 4.0, 2.55, 9.3, 2.55, C.gold, 2);

  txt(s, "الغذاء المتوازن ليس مجرد احتياج يومي،",
    1.0, 3.05, 11.3, 0.5, {
      fontSize: 22, color: "D9E9EF", align: "center"
    });

  txt(s, "بل هو استثمار في الصحة والطاقة والأداء.",
    1.0, 3.68, 11.3, 0.55, {
      fontSize: 25, bold: true, color: C.cyan, align: "center"
    });

  shape(s, pptx.ShapeType.roundRect, 3.25, 5.0, 6.85, 0.95, C.navy2, "426176");
  txt(s, "شكراً لحسن استماعكم",
    3.5, 5.22, 6.35, 0.45, {
      fontSize: 27, bold: true, color: C.gold, align: "center"
    });

  txt(s, "إعداد وتقديم الدكتورة مفيدة المختار راشد",
    2.0, 6.35, 9.3, 0.3, {
      fontSize: 14, color: C.white, align: "center"
    });

  footer(s, 14, true);
}

// إنشاء ملف PowerPoint
pptx.writeFile({ fileName: path.join(__dirname, "Sports-Nutrition.pptx") })
  .then(() => {
    console.log("Presentation created successfully: 14 slides");
  })
  .catch((err) => {
    console.error("Error creating presentation:", err);
    process.exit(1);
  });
