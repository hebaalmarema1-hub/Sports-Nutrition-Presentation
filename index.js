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

// الشريحة الأولى: الغلاف الرئيسي

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
// ==============================
// الشريحة الثانية: مفهوم التغذية الرياضية
// ==============================

const slide2 = pptx.addSlide();

slide2.background = { color: C.pale };

// شريط علوي
slide2.addShape(pptx.ShapeType.rect, {
  x: 0,
  y: 0,
  w: SW,
  h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide2, "02", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide2, "مفهوم التغذية الرياضية", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// خط زخرفي
slide2.addShape(pptx.ShapeType.line, {
  x: 7.2,
  y: 1.55,
  w: 5.3,
  h: 0,
  line: { color: C.gold, width: 2 }
});

// التعريف
addText(
  slide2,
  "التغذية الرياضية هي العلم الذي يهتم بتوفير العناصر الغذائية والطاقة اللازمة للرياضيين، بما يتناسب مع طبيعة النشاط البدني وشدته ومدته.",
  1.1, 2.0, 11.1, 1.5,
  {
    fontSize: 23,
    color: C.navy,
    bold: true,
    align: "right",
    valign: "mid",
    breakLine: true,
    margin: 0.15
  }
);

// ثلاثة محاور
const topics = [
  {
    x: 0.9,
    title: "توفير الطاقة",
    desc: "دعم الجسم بالطاقة اللازمة للنشاط البدني"
  },
  {
    x: 4.85,
    title: "تحسين الأداء",
    desc: "المساعدة على رفع الكفاءة البدنية"
  },
  {
    x: 8.8,
    title: "سرعة الاستشفاء",
    desc: "دعم تعافي العضلات بعد المجهود"
  }
];

for (const item of topics) {
  slide2.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 4.25,
    w: 3.55,
    h: 1.65,
    rectRadius: 0.12,
    line: { color: C.gray, width: 1 },
    fill: { color: C.white }
  });

  addText(slide2, item.title, item.x + 0.15, 4.55, 3.25, 0.4, {
    fontSize: 20,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide2, item.desc, item.x + 0.2, 5.05, 3.15, 0.55, {
    fontSize: 14,
    color: C.navy,
    align: "center"
  });
}

// تذييل
addText(slide2, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الثالثة: مصادر الطاقة في الجسم
// ==============================

const slide3 = pptx.addSlide();

slide3.background = { color: C.white };

// شريط علوي
slide3.addShape(pptx.ShapeType.rect, {
  x: 0,
  y: 0,
  w: SW,
  h: 0.16,
  line: { color: C.blue, transparency: 100 },
  fill: { color: C.blue }
});

// رقم الشريحة
addText(slide3, "03", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide3, "مصادر الطاقة في الجسم", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide3,
  "يعتمد الجسم على ثلاثة عناصر غذائية رئيسية لإنتاج الطاقة ودعم النشاط البدني:",
  1.0, 1.65, 11.2, 0.8,
  {
    fontSize: 21,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// البطاقات
const energySources = [
  {
    x: 0.8,
    title: "الكربوهيدرات",
    desc: "المصدر الأساسي للطاقة أثناء النشاط البدني، وتساعد على تجديد مخزون الجليكوجين في العضلات."
  },
  {
    x: 4.85,
    title: "الدهون",
    desc: "مصدر مهم للطاقة، خاصة أثناء الأنشطة البدنية طويلة المدة ومنخفضة إلى متوسطة الشدة."
  },
  {
    x: 8.9,
    title: "البروتينات",
    desc: "تساهم في بناء العضلات وإصلاح الأنسجة، ويمكن استخدامها كمصدر للطاقة عند الحاجة."
  }
];

for (const item of energySources) {
  slide3.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 2.5,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.5 },
    fill: { color: C.pale }
  });

  addText(slide3, item.title, item.x + 0.15, 3.35, 3.25, 0.5, {
    fontSize: 22,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide3, item.desc, item.x + 0.25, 4.05, 3.05, 1.1, {
    fontSize: 16,
    color: C.navy,
    align: "center",
    valign: "mid",
    breakLine: true,
    margin: 0.1
  });
}

// تذييل
addText(slide3, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الرابعة: مكونات الغذاء المتوازن
// ==============================

const slide4 = pptx.addSlide();

slide4.background = { color: C.pale };

// الشريط العلوي
slide4.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide4, "04", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide4, "مكونات الغذاء المتوازن", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide4,
  "يحتاج الرياضي إلى تنوع غذائي يضمن حصول الجسم على العناصر الضرورية للطاقة والنمو والتعافي.",
  1.0, 1.55, 11.2, 0.8,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// مكونات الغذاء
const foodItems = [
  {
    x: 0.8, y: 2.8,
    title: "الكربوهيدرات",
    desc: "تمد الجسم بالطاقة اللازمة للنشاط البدني."
  },
  {
    x: 4.85, y: 2.8,
    title: "البروتينات",
    desc: "تساعد على بناء العضلات وإصلاح الأنسجة."
  },
  {
    x: 8.9, y: 2.8,
    title: "الدهون الصحية",
    desc: "تساهم في توفير الطاقة وامتصاص بعض الفيتامينات."
  },
  {
    x: 2.8, y: 4.8,
    title: "الفيتامينات والمعادن",
    desc: "تدعم وظائف الجسم والعضلات والعظام."
  },
  {
    x: 6.85, y: 4.8,
    title: "الماء",
    desc: "يحافظ على توازن السوائل وتنظيم حرارة الجسم."
  }
];

for (const item of foodItems) {
  slide4.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: item.y,
    w: 3.55,
    h: 1.45,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.2 },
    fill: { color: C.white }
  });

  addText(slide4, item.title, item.x + 0.1, item.y + 0.2, 3.35, 0.4, {
    fontSize: 19,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide4, item.desc, item.x + 0.2, item.y + 0.7, 3.15, 0.55, {
    fontSize: 14,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// التذييل
addText(slide4, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الخامسة: أهمية الماء والترطيب
// ==============================

const slide5 = pptx.addSlide();

slide5.background = { color: C.white };

// الشريط العلوي
slide5.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide5, "05", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide5, "أهمية الماء والترطيب للرياضي", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide5,
  "يُعد الماء عنصرًا أساسيًا للحفاظ على كفاءة الجسم أثناء النشاط البدني، ويساعد على تعويض السوائل المفقودة عن طريق التعرق.",
  1.0, 1.65, 11.2, 1.0,
  {
    fontSize: 21,
    color: C.navy,
    align: "right",
    valign: "mid",
    margin: 0.1
  }
);

// فوائد الترطيب
const hydration = [
  {
    x: 0.8,
    title: "تنظيم حرارة الجسم",
    desc: "يساعد على تبريد الجسم أثناء المجهود البدني."
  },
  {
    x: 4.85,
    title: "الحفاظ على الأداء",
    desc: "يقلل تأثير الجفاف على القوة والتحمل والتركيز."
  },
  {
    x: 8.9,
    title: "تعويض السوائل",
    desc: "يعوض الماء والأملاح المفقودة مع التعرق."
  }
];

for (const item of hydration) {
  slide5.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.3,
    w: 3.55,
    h: 1.9,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.3 },
    fill: { color: C.pale }
  });

  addText(slide5, item.title, item.x + 0.1, 3.65, 3.35, 0.45, {
    fontSize: 19,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide5, item.desc, item.x + 0.2, 4.25, 3.15, 0.65, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// ملاحظة
addText(
  slide5,
  "يُنصح بشرب الماء بانتظام قبل النشاط البدني وأثناءه وبعده، وفقًا لاحتياجات الجسم وظروف النشاط.",
  1.0, 5.75, 11.2, 0.65,
  {
    fontSize: 18,
    bold: true,
    color: C.navy,
    align: "center",
    valign: "mid"
  }
);

// التذييل
addText(slide5, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة السادسة: التوقيت الغذائي
// ==============================

const slide6 = pptx.addSlide();

slide6.background = { color: C.pale };

// الشريط العلوي
slide6.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide6, "06", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide6, "التوقيت الغذائي والنشاط البدني", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide6,
  "لا تقتصر أهمية التغذية على نوع الطعام فقط، بل يشمل ذلك توقيت تناوله بما يتناسب مع النشاط البدني.",
  1.0, 1.5, 11.2, 0.8,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// مراحل التوقيت الغذائي
const timing = [
  {
    x: 0.8,
    title: "قبل النشاط",
    desc: "تناول وجبة متوازنة تحتوي على الكربوهيدرات، مع كمية مناسبة من البروتين، وشرب الماء."
  },
  {
    x: 4.85,
    title: "أثناء النشاط",
    desc: "المحافظة على الترطيب بشرب الماء، وقد يحتاج الرياضي إلى الكربوهيدرات في الأنشطة الطويلة."
  },
  {
    x: 8.9,
    title: "بعد النشاط",
    desc: "تناول البروتين لدعم إصلاح العضلات، والكربوهيدرات لتعويض مخزون الجليكوجين."
  }
];

for (const item of timing) {
  slide6.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 2.35,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.4 },
    fill: { color: C.white }
  });

  addText(slide6, item.title, item.x + 0.1, 3.35, 3.35, 0.5, {
    fontSize: 22,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide6, item.desc, item.x + 0.2, 4.05, 3.15, 1.0, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid",
    breakLine: true,
    margin: 0.1
  });
}

// خلاصة
addText(
  slide6,
  "التغذية المناسبة في الوقت المناسب تساعد على تحسين الأداء وتسريع الاستشفاء.",
  1.0, 5.85, 11.2, 0.6,
  {
    fontSize: 19,
    bold: true,
    color: C.navy,
    align: "center",
    valign: "mid"
  }
);

// التذييل
addText(slide6, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة السابعة: التغذية والاستشفاء العضلي
// ==============================

const slide7 = pptx.addSlide();

slide7.background = { color: C.white };

// الشريط العلوي
slide7.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide7, "07", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide7, "التغذية والاستشفاء العضلي", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide7,
  "تؤدي التغذية دورًا مهمًا في تعافي الجسم بعد المجهود البدني، والمساعدة على إصلاح الأنسجة واستعادة الطاقة.",
  1.0, 1.55, 11.2, 0.85,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// عناصر الاستشفاء
const recovery = [
  {
    x: 0.8,
    title: "البروتين",
    desc: "يساعد على إصلاح الأنسجة العضلية ودعم عملية بناء العضلات."
  },
  {
    x: 4.85,
    title: "الكربوهيدرات",
    desc: "تساهم في تعويض مخزون الجليكوجين واستعادة الطاقة."
  },
  {
    x: 8.9,
    title: "الراحة والنوم",
    desc: "يساعدان الجسم على التعافي والتكيف مع التدريب."
  }
];

for (const item of recovery) {
  slide7.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 2.0,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.3 },
    fill: { color: C.pale }
  });

  addText(slide7, item.title, item.x + 0.1, 3.35, 3.35, 0.45, {
    fontSize: 21,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide7, item.desc, item.x + 0.2, 4.0, 3.15, 0.7, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// خلاصة
addText(
  slide7,
  "الاستشفاء الجيد لا يعتمد على الغذاء وحده، بل على تكامل التغذية والترطيب والراحة والنوم.",
  1.0, 5.65, 11.2, 0.75,
  {
    fontSize: 19,
    bold: true,
    color: C.navy,
    align: "center",
    valign: "mid"
  }
);

// التذييل
addText(slide7, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الثامنة: الأخطاء الغذائية الشائعة
// ==============================

const slide8 = pptx.addSlide();

slide8.background = { color: C.pale };

// الشريط العلوي
slide8.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide8, "08", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide8, "الأخطاء الغذائية الشائعة لدى الرياضيين", 2.0, 0.55, 10.5, 0.8, {
  fontSize: 28,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide8,
  "قد تؤثر بعض العادات الغذائية غير الصحيحة سلبًا على الأداء الرياضي وصحة الجسم.",
  1.0, 1.5, 11.2, 0.7,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// الأخطاء
const mistakes = [
  {
    x: 0.8, y: 2.65,
    title: "إهمال وجبة الإفطار",
    desc: "قد يؤدي إلى نقص الطاقة والتركيز."
  },
  {
    x: 4.85, y: 2.65,
    title: "قلة شرب الماء",
    desc: "تزيد من احتمالية الجفاف وانخفاض الأداء."
  },
  {
    x: 8.9, y: 2.65,
    title: "الإفراط في المكملات",
    desc: "استخدامها دون حاجة أو إرشاد متخصص."
  },
  {
    x: 2.8, y: 4.65,
    title: "الاعتماد على نوع واحد من الغذاء",
    desc: "يقلل تنوع العناصر الغذائية التي يحتاجها الجسم."
  },
  {
    x: 6.85, y: 4.65,
    title: "تناول الطعام في توقيت غير مناسب",
    desc: "قد يؤثر على الراحة والأداء أثناء التدريب."
  }
];

for (const item of mistakes) {
  slide8.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: item.y,
    w: 3.55,
    h: 1.45,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.2 },
    fill: { color: C.white }
  });

  addText(slide8, item.title, item.x + 0.1, item.y + 0.2, 3.35, 0.4, {
    fontSize: 17,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide8, item.desc, item.x + 0.2, item.y + 0.7, 3.15, 0.55, {
    fontSize: 14,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// التذييل
addText(slide8, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة التاسعة: علاقة التغذية بالأداء الرياضي
// ==============================

const slide9 = pptx.addSlide();

slide9.background = { color: C.white };

// الشريط العلوي
slide9.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide9, "09", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide9, "علاقة التغذية بالأداء الرياضي", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide9,
  "تؤثر التغذية المتوازنة بصورة مباشرة في قدرة الرياضي على أداء التمارين بكفاءة والمحافظة على مستوى بدني جيد.",
  1.0, 1.5, 11.2, 0.8,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// عناصر الأداء
const performance = [
  {
    x: 0.8,
    title: "القوة العضلية",
    desc: "توفير الطاقة والعناصر اللازمة لدعم عمل العضلات."
  },
  {
    x: 4.85,
    title: "التحمل البدني",
    desc: "المساعدة على استمرار النشاط وتقليل التعب المبكر."
  },
  {
    x: 8.9,
    title: "التركيز وسرعة الاستجابة",
    desc: "دعم وظائف الدماغ والانتباه أثناء النشاط الرياضي."
  }
];

for (const item of performance) {
  slide9.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 2.0,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.3 },
    fill: { color: C.pale }
  });

  addText(slide9, item.title, item.x + 0.1, 3.35, 3.35, 0.5, {
    fontSize: 19,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide9, item.desc, item.x + 0.2, 4.0, 3.15, 0.7, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// خلاصة
addText(
  slide9,
  "كلما كانت التغذية مناسبة لاحتياجات الرياضي، زادت فرص تحسين الأداء والمحافظة على الصحة.",
  1.0, 5.65, 11.2, 0.7,
  {
    fontSize: 19,
    bold: true,
    color: C.navy,
    align: "center",
    valign: "mid"
  }
);

// التذييل
addText(slide9, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة العاشرة: دور معلم التربية البدنية
// ==============================

const slide10 = pptx.addSlide();

slide10.background = { color: C.pale };

// الشريط العلوي
slide10.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide10, "10", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide10, "دور معلم التربية البدنية في التوعية الغذائية", 1.5, 0.55, 11, 0.8, {
  fontSize: 27,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide10,
  "لا يقتصر دور معلم التربية البدنية على تعليم المهارات والتمارين، بل يمتد إلى تعزيز السلوكيات الصحية لدى الطلاب.",
  1.0, 1.55, 11.2, 0.85,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// الأدوار
const teacherRoles = [
  {
    x: 0.8,
    title: "نشر الوعي الغذائي",
    desc: "توضيح أهمية الغذاء المتوازن للطلاب والرياضيين."
  },
  {
    x: 4.85,
    title: "تشجيع العادات الصحية",
    desc: "تعزيز شرب الماء وتناول الوجبات المتوازنة."
  },
  {
    x: 8.9,
    title: "توجيه الطلاب",
    desc: "تشجيعهم على اتباع أساليب صحية تناسب نشاطهم."
  }
];

for (const item of teacherRoles) {
  slide10.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 2.0,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.3 },
    fill: { color: C.white }
  });

  addText(slide10, item.title, item.x + 0.1, 3.35, 3.35, 0.5, {
    fontSize: 19,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide10, item.desc, item.x + 0.2, 4.0, 3.15, 0.7, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// خلاصة
addText(
  slide10,
  "المعلم الواعي يسهم في بناء جيل أكثر صحة ونشاطًا من خلال التوجيه والتثقيف.",
  1.0, 5.65, 11.2, 0.7,
  {
    fontSize: 19,
    bold: true,
    color: C.navy,
    align: "center",
    valign: "mid"
  }
);

// التذييل
addText(slide10, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الحادية عشرة: وجبة غذائية متوازنة
// ==============================

const slide11 = pptx.addSlide();

slide11.background = { color: C.white };

// الشريط العلوي
slide11.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide11, "11", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide11, "مثال عملي لوجبة غذائية متوازنة للرياضي", 1.5, 0.55, 11, 0.8, {
  fontSize: 28,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide11,
  "يمكن تكوين وجبة متوازنة من مجموعة متنوعة من الأطعمة التي توفر العناصر الغذائية الأساسية.",
  1.0, 1.5, 11.2, 0.75,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// مكونات الوجبة
const meal = [
  {
    x: 0.8,
    title: "الأرز",
    desc: "مصدر للكربوهيدرات والطاقة."
  },
  {
    x: 4.85,
    title: "الدجاج",
    desc: "مصدر للبروتين ودعم إصلاح العضلات."
  },
  {
    x: 8.9,
    title: "الخضروات",
    desc: "توفر الفيتامينات والمعادن والألياف."
  }
];

for (const item of meal) {
  slide11.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 1.8,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.3 },
    fill: { color: C.pale }
  });

  addText(slide11, item.title, item.x + 0.1, 3.35, 3.35, 0.45, {
    fontSize: 21,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide11, item.desc, item.x + 0.2, 4.0, 3.15, 0.55, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// الماء
slide11.addShape(pptx.ShapeType.roundRect, {
  x: 4.35,
  y: 5.15,
  w: 4.6,
  h: 0.9,
  rectRadius: 0.12,
  line: { color: C.cyan, width: 1.3 },
  fill: { color: C.pale }
});

addText(slide11, "الماء: عنصر أساسي للحفاظ على الترطيب وتوازن السوائل.", 4.5, 5.4, 4.3, 0.4, {
  fontSize: 16,
  bold: true,
  color: C.blue,
  align: "center"
});

// إنشاء ملف PowerPoint
// ==============================
// الشريحة الثانية: مفهوم التغذية الرياضية
// ==============================

const slide2 = pptx.addSlide();

slide2.background = { color: C.pale };

// شريط علوي
slide2.addShape(pptx.ShapeType.rect, {
  x: 0,
  y: 0,
  w: SW,
  h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide2, "02", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide2, "مفهوم التغذية الرياضية", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// خط زخرفي
slide2.addShape(pptx.ShapeType.line, {
  x: 7.2,
  y: 1.55,
  w: 5.3,
  h: 0,
  line: { color: C.gold, width: 2 }
});

// التعريف
addText(
  slide2,
  "التغذية الرياضية هي العلم الذي يهتم بتوفير العناصر الغذائية والطاقة اللازمة للرياضيين، بما يتناسب مع طبيعة النشاط البدني وشدته ومدته.",
  1.1, 2.0, 11.1, 1.5,
  {
    fontSize: 23,
    color: C.navy,
    bold: true,
    align: "right",
    valign: "mid",
    breakLine: true,
    margin: 0.15
  }
);

// ثلاثة محاور
const topics = [
  {
    x: 0.9,
    title: "توفير الطاقة",
    desc: "دعم الجسم بالطاقة اللازمة للنشاط البدني"
  },
  {
    x: 4.85,
    title: "تحسين الأداء",
    desc: "المساعدة على رفع الكفاءة البدنية"
  },
  {
    x: 8.8,
    title: "سرعة الاستشفاء",
    desc: "دعم تعافي العضلات بعد المجهود"
  }
];

for (const item of topics) {
  slide2.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 4.25,
    w: 3.55,
    h: 1.65,
    rectRadius: 0.12,
    line: { color: C.gray, width: 1 },
    fill: { color: C.white }
  });

  addText(slide2, item.title, item.x + 0.15, 4.55, 3.25, 0.4, {
    fontSize: 20,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide2, item.desc, item.x + 0.2, 5.05, 3.15, 0.55, {
    fontSize: 14,
    color: C.navy,
    align: "center"
  });
}

// تذييل
addText(slide2, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الثالثة: مصادر الطاقة في الجسم
// ==============================

const slide3 = pptx.addSlide();

slide3.background = { color: C.white };

// شريط علوي
slide3.addShape(pptx.ShapeType.rect, {
  x: 0,
  y: 0,
  w: SW,
  h: 0.16,
  line: { color: C.blue, transparency: 100 },
  fill: { color: C.blue }
});

// رقم الشريحة
addText(slide3, "03", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide3, "مصادر الطاقة في الجسم", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide3,
  "يعتمد الجسم على ثلاثة عناصر غذائية رئيسية لإنتاج الطاقة ودعم النشاط البدني:",
  1.0, 1.65, 11.2, 0.8,
  {
    fontSize: 21,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// البطاقات
const energySources = [
  {
    x: 0.8,
    title: "الكربوهيدرات",
    desc: "المصدر الأساسي للطاقة أثناء النشاط البدني، وتساعد على تجديد مخزون الجليكوجين في العضلات."
  },
  {
    x: 4.85,
    title: "الدهون",
    desc: "مصدر مهم للطاقة، خاصة أثناء الأنشطة البدنية طويلة المدة ومنخفضة إلى متوسطة الشدة."
  },
  {
    x: 8.9,
    title: "البروتينات",
    desc: "تساهم في بناء العضلات وإصلاح الأنسجة، ويمكن استخدامها كمصدر للطاقة عند الحاجة."
  }
];

for (const item of energySources) {
  slide3.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 2.5,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.5 },
    fill: { color: C.pale }
  });

  addText(slide3, item.title, item.x + 0.15, 3.35, 3.25, 0.5, {
    fontSize: 22,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide3, item.desc, item.x + 0.25, 4.05, 3.05, 1.1, {
    fontSize: 16,
    color: C.navy,
    align: "center",
    valign: "mid",
    breakLine: true,
    margin: 0.1
  });
}

// تذييل
addText(slide3, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الرابعة: مكونات الغذاء المتوازن
// ==============================

const slide4 = pptx.addSlide();

slide4.background = { color: C.pale };

// الشريط العلوي
slide4.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide4, "04", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide4, "مكونات الغذاء المتوازن", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide4,
  "يحتاج الرياضي إلى تنوع غذائي يضمن حصول الجسم على العناصر الضرورية للطاقة والنمو والتعافي.",
  1.0, 1.55, 11.2, 0.8,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// مكونات الغذاء
const foodItems = [
  {
    x: 0.8, y: 2.8,
    title: "الكربوهيدرات",
    desc: "تمد الجسم بالطاقة اللازمة للنشاط البدني."
  },
  {
    x: 4.85, y: 2.8,
    title: "البروتينات",
    desc: "تساعد على بناء العضلات وإصلاح الأنسجة."
  },
  {
    x: 8.9, y: 2.8,
    title: "الدهون الصحية",
    desc: "تساهم في توفير الطاقة وامتصاص بعض الفيتامينات."
  },
  {
    x: 2.8, y: 4.8,
    title: "الفيتامينات والمعادن",
    desc: "تدعم وظائف الجسم والعضلات والعظام."
  },
  {
    x: 6.85, y: 4.8,
    title: "الماء",
    desc: "يحافظ على توازن السوائل وتنظيم حرارة الجسم."
  }
];

for (const item of foodItems) {
  slide4.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: item.y,
    w: 3.55,
    h: 1.45,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.2 },
    fill: { color: C.white }
  });

  addText(slide4, item.title, item.x + 0.1, item.y + 0.2, 3.35, 0.4, {
    fontSize: 19,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide4, item.desc, item.x + 0.2, item.y + 0.7, 3.15, 0.55, {
    fontSize: 14,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// التذييل
addText(slide4, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الخامسة: أهمية الماء والترطيب
// ==============================

const slide5 = pptx.addSlide();

slide5.background = { color: C.white };

// الشريط العلوي
slide5.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide5, "05", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide5, "أهمية الماء والترطيب للرياضي", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide5,
  "يُعد الماء عنصرًا أساسيًا للحفاظ على كفاءة الجسم أثناء النشاط البدني، ويساعد على تعويض السوائل المفقودة عن طريق التعرق.",
  1.0, 1.65, 11.2, 1.0,
  {
    fontSize: 21,
    color: C.navy,
    align: "right",
    valign: "mid",
    margin: 0.1
  }
);

// فوائد الترطيب
const hydration = [
  {
    x: 0.8,
    title: "تنظيم حرارة الجسم",
    desc: "يساعد على تبريد الجسم أثناء المجهود البدني."
  },
  {
    x: 4.85,
    title: "الحفاظ على الأداء",
    desc: "يقلل تأثير الجفاف على القوة والتحمل والتركيز."
  },
  {
    x: 8.9,
    title: "تعويض السوائل",
    desc: "يعوض الماء والأملاح المفقودة مع التعرق."
  }
];

for (const item of hydration) {
  slide5.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.3,
    w: 3.55,
    h: 1.9,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.3 },
    fill: { color: C.pale }
  });

  addText(slide5, item.title, item.x + 0.1, 3.65, 3.35, 0.45, {
    fontSize: 19,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide5, item.desc, item.x + 0.2, 4.25, 3.15, 0.65, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// ملاحظة
addText(
  slide5,
  "يُنصح بشرب الماء بانتظام قبل النشاط البدني وأثناءه وبعده، وفقًا لاحتياجات الجسم وظروف النشاط.",
  1.0, 5.75, 11.2, 0.65,
  {
    fontSize: 18,
    bold: true,
    color: C.navy,
    align: "center",
    valign: "mid"
  }
);

// التذييل
addText(slide5, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة السادسة: التوقيت الغذائي
// ==============================

const slide6 = pptx.addSlide();

slide6.background = { color: C.pale };

// الشريط العلوي
slide6.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide6, "06", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide6, "التوقيت الغذائي والنشاط البدني", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide6,
  "لا تقتصر أهمية التغذية على نوع الطعام فقط، بل يشمل ذلك توقيت تناوله بما يتناسب مع النشاط البدني.",
  1.0, 1.5, 11.2, 0.8,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// مراحل التوقيت الغذائي
const timing = [
  {
    x: 0.8,
    title: "قبل النشاط",
    desc: "تناول وجبة متوازنة تحتوي على الكربوهيدرات، مع كمية مناسبة من البروتين، وشرب الماء."
  },
  {
    x: 4.85,
    title: "أثناء النشاط",
    desc: "المحافظة على الترطيب بشرب الماء، وقد يحتاج الرياضي إلى الكربوهيدرات في الأنشطة الطويلة."
  },
  {
    x: 8.9,
    title: "بعد النشاط",
    desc: "تناول البروتين لدعم إصلاح العضلات، والكربوهيدرات لتعويض مخزون الجليكوجين."
  }
];

for (const item of timing) {
  slide6.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 2.35,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.4 },
    fill: { color: C.white }
  });

  addText(slide6, item.title, item.x + 0.1, 3.35, 3.35, 0.5, {
    fontSize: 22,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide6, item.desc, item.x + 0.2, 4.05, 3.15, 1.0, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid",
    breakLine: true,
    margin: 0.1
  });
}

// خلاصة
addText(
  slide6,
  "التغذية المناسبة في الوقت المناسب تساعد على تحسين الأداء وتسريع الاستشفاء.",
  1.0, 5.85, 11.2, 0.6,
  {
    fontSize: 19,
    bold: true,
    color: C.navy,
    align: "center",
    valign: "mid"
  }
);

// التذييل
addText(slide6, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة السابعة: التغذية والاستشفاء العضلي
// ==============================

const slide7 = pptx.addSlide();

slide7.background = { color: C.white };

// الشريط العلوي
slide7.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide7, "07", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide7, "التغذية والاستشفاء العضلي", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide7,
  "تؤدي التغذية دورًا مهمًا في تعافي الجسم بعد المجهود البدني، والمساعدة على إصلاح الأنسجة واستعادة الطاقة.",
  1.0, 1.55, 11.2, 0.85,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// عناصر الاستشفاء
const recovery = [
  {
    x: 0.8,
    title: "البروتين",
    desc: "يساعد على إصلاح الأنسجة العضلية ودعم عملية بناء العضلات."
  },
  {
    x: 4.85,
    title: "الكربوهيدرات",
    desc: "تساهم في تعويض مخزون الجليكوجين واستعادة الطاقة."
  },
  {
    x: 8.9,
    title: "الراحة والنوم",
    desc: "يساعدان الجسم على التعافي والتكيف مع التدريب."
  }
];

for (const item of recovery) {
  slide7.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 2.0,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.3 },
    fill: { color: C.pale }
  });

  addText(slide7, item.title, item.x + 0.1, 3.35, 3.35, 0.45, {
    fontSize: 21,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide7, item.desc, item.x + 0.2, 4.0, 3.15, 0.7, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// خلاصة
addText(
  slide7,
  "الاستشفاء الجيد لا يعتمد على الغذاء وحده، بل على تكامل التغذية والترطيب والراحة والنوم.",
  1.0, 5.65, 11.2, 0.75,
  {
    fontSize: 19,
    bold: true,
    color: C.navy,
    align: "center",
    valign: "mid"
  }
);

// التذييل
addText(slide7, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الثامنة: الأخطاء الغذائية الشائعة
// ==============================

const slide8 = pptx.addSlide();

slide8.background = { color: C.pale };

// الشريط العلوي
slide8.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide8, "08", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide8, "الأخطاء الغذائية الشائعة لدى الرياضيين", 2.0, 0.55, 10.5, 0.8, {
  fontSize: 28,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide8,
  "قد تؤثر بعض العادات الغذائية غير الصحيحة سلبًا على الأداء الرياضي وصحة الجسم.",
  1.0, 1.5, 11.2, 0.7,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// الأخطاء
const mistakes = [
  {
    x: 0.8, y: 2.65,
    title: "إهمال وجبة الإفطار",
    desc: "قد يؤدي إلى نقص الطاقة والتركيز."
  },
  {
    x: 4.85, y: 2.65,
    title: "قلة شرب الماء",
    desc: "تزيد من احتمالية الجفاف وانخفاض الأداء."
  },
  {
    x: 8.9, y: 2.65,
    title: "الإفراط في المكملات",
    desc: "استخدامها دون حاجة أو إرشاد متخصص."
  },
  {
    x: 2.8, y: 4.65,
    title: "الاعتماد على نوع واحد من الغذاء",
    desc: "يقلل تنوع العناصر الغذائية التي يحتاجها الجسم."
  },
  {
    x: 6.85, y: 4.65,
    title: "تناول الطعام في توقيت غير مناسب",
    desc: "قد يؤثر على الراحة والأداء أثناء التدريب."
  }
];

for (const item of mistakes) {
  slide8.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: item.y,
    w: 3.55,
    h: 1.45,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.2 },
    fill: { color: C.white }
  });

  addText(slide8, item.title, item.x + 0.1, item.y + 0.2, 3.35, 0.4, {
    fontSize: 17,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide8, item.desc, item.x + 0.2, item.y + 0.7, 3.15, 0.55, {
    fontSize: 14,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// التذييل
addText(slide8, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة التاسعة: علاقة التغذية بالأداء الرياضي
// ==============================

const slide9 = pptx.addSlide();

slide9.background = { color: C.white };

// الشريط العلوي
slide9.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide9, "09", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide9, "علاقة التغذية بالأداء الرياضي", 2.2, 0.55, 10.3, 0.8, {
  fontSize: 30,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide9,
  "تؤثر التغذية المتوازنة بصورة مباشرة في قدرة الرياضي على أداء التمارين بكفاءة والمحافظة على مستوى بدني جيد.",
  1.0, 1.5, 11.2, 0.8,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// عناصر الأداء
const performance = [
  {
    x: 0.8,
    title: "القوة العضلية",
    desc: "توفير الطاقة والعناصر اللازمة لدعم عمل العضلات."
  },
  {
    x: 4.85,
    title: "التحمل البدني",
    desc: "المساعدة على استمرار النشاط وتقليل التعب المبكر."
  },
  {
    x: 8.9,
    title: "التركيز وسرعة الاستجابة",
    desc: "دعم وظائف الدماغ والانتباه أثناء النشاط الرياضي."
  }
];

for (const item of performance) {
  slide9.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 2.0,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.3 },
    fill: { color: C.pale }
  });

  addText(slide9, item.title, item.x + 0.1, 3.35, 3.35, 0.5, {
    fontSize: 19,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide9, item.desc, item.x + 0.2, 4.0, 3.15, 0.7, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// خلاصة
addText(
  slide9,
  "كلما كانت التغذية مناسبة لاحتياجات الرياضي، زادت فرص تحسين الأداء والمحافظة على الصحة.",
  1.0, 5.65, 11.2, 0.7,
  {
    fontSize: 19,
    bold: true,
    color: C.navy,
    align: "center",
    valign: "mid"
  }
);

// التذييل
addText(slide9, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة العاشرة: دور معلم التربية البدنية
// ==============================

const slide10 = pptx.addSlide();

slide10.background = { color: C.pale };

// الشريط العلوي
slide10.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide10, "10", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide10, "دور معلم التربية البدنية في التوعية الغذائية", 1.5, 0.55, 11, 0.8, {
  fontSize: 27,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide10,
  "لا يقتصر دور معلم التربية البدنية على تعليم المهارات والتمارين، بل يمتد إلى تعزيز السلوكيات الصحية لدى الطلاب.",
  1.0, 1.55, 11.2, 0.85,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// الأدوار
const teacherRoles = [
  {
    x: 0.8,
    title: "نشر الوعي الغذائي",
    desc: "توضيح أهمية الغذاء المتوازن للطلاب والرياضيين."
  },
  {
    x: 4.85,
    title: "تشجيع العادات الصحية",
    desc: "تعزيز شرب الماء وتناول الوجبات المتوازنة."
  },
  {
    x: 8.9,
    title: "توجيه الطلاب",
    desc: "تشجيعهم على اتباع أساليب صحية تناسب نشاطهم."
  }
];

for (const item of teacherRoles) {
  slide10.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 2.0,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.3 },
    fill: { color: C.white }
  });

  addText(slide10, item.title, item.x + 0.1, 3.35, 3.35, 0.5, {
    fontSize: 19,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide10, item.desc, item.x + 0.2, 4.0, 3.15, 0.7, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// خلاصة
addText(
  slide10,
  "المعلم الواعي يسهم في بناء جيل أكثر صحة ونشاطًا من خلال التوجيه والتثقيف.",
  1.0, 5.65, 11.2, 0.7,
  {
    fontSize: 19,
    bold: true,
    color: C.navy,
    align: "center",
    valign: "mid"
  }
);

// التذييل
addText(slide10, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الحادية عشرة: وجبة غذائية متوازنة
// ==============================

const slide11 = pptx.addSlide();

slide11.background = { color: C.white };

// الشريط العلوي
slide11.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide11, "11", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide11, "مثال عملي لوجبة غذائية متوازنة للرياضي", 1.5, 0.55, 11, 0.8, {
  fontSize: 28,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide11,
  "يمكن تكوين وجبة متوازنة من مجموعة متنوعة من الأطعمة التي توفر العناصر الغذائية الأساسية.",
  1.0, 1.5, 11.2, 0.75,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// مكونات الوجبة
const meal = [
  {
    x: 0.8,
    title: "الأرز",
    desc: "مصدر للكربوهيدرات والطاقة."
  },
  {
    x: 4.85,
    title: "الدجاج",
    desc: "مصدر للبروتين ودعم إصلاح العضلات."
  },
  {
    x: 8.9,
    title: "الخضروات",
    desc: "توفر الفيتامينات والمعادن والألياف."
  }
];

for (const item of meal) {
  slide11.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 1.8,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.3 },
    fill: { color: C.pale }
  });

  addText(slide11, item.title, item.x + 0.1, 3.35, 3.35, 0.45, {
    fontSize: 21,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide11, item.desc, item.x + 0.2, 4.0, 3.15, 0.55, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// الماء
slide11.addShape(pptx.ShapeType.roundRect, {
  x: 4.35,
  y: 5.15,
  w: 4.6,
  h: 0.9,
  rectRadius: 0.12,
  line: { color: C.cyan, width: 1.3 },
  fill: { color: C.pale }
});

addText(slide11, "الماء: عنصر أساسي للحفاظ على الترطيب وتوازن السوائل.", 4.5, 5.4, 4.3, 0.4, {
  fontSize: 16,
  bold: true,
  color: C.blue,
  align: "center"
});

// التذييل
addText(slide11, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الحادية عشرة: وجبة غذائية متوازنة
// ==============================

const slide11 = pptx.addSlide();

slide11.background = { color: C.white };

// الشريط العلوي
slide11.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.cyan, transparency: 100 },
  fill: { color: C.cyan }
});

// رقم الشريحة
addText(slide11, "11", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.blue
});

// العنوان
addText(slide11, "مثال عملي لوجبة غذائية متوازنة للرياضي", 1.5, 0.55, 11, 0.8, {
  fontSize: 28,
  bold: true,
  color: C.navy,
  align: "right"
});

// مقدمة
addText(
  slide11,
  "يمكن تكوين وجبة متوازنة من مجموعة متنوعة من الأطعمة التي توفر العناصر الغذائية الأساسية.",
  1.0, 1.5, 11.2, 0.75,
  {
    fontSize: 20,
    color: C.navy,
    align: "right",
    valign: "mid"
  }
);

// مكونات الوجبة
const meal = [
  {
    x: 0.8,
    title: "الأرز",
    desc: "مصدر للكربوهيدرات والطاقة."
  },
  {
    x: 4.85,
    title: "الدجاج",
    desc: "مصدر للبروتين ودعم إصلاح العضلات."
  },
  {
    x: 8.9,
    title: "الخضروات",
    desc: "توفر الفيتامينات والمعادن والألياف."
  }
];

for (const item of meal) {
  slide11.addShape(pptx.ShapeType.roundRect, {
    x: item.x,
    y: 3.0,
    w: 3.55,
    h: 1.8,
    rectRadius: 0.12,
    line: { color: C.cyan, width: 1.3 },
    fill: { color: C.pale }
  });

  addText(slide11, item.title, item.x + 0.1, 3.35, 3.35, 0.45, {
    fontSize: 21,
    bold: true,
    color: C.blue,
    align: "center"
  });

  addText(slide11, item.desc, item.x + 0.2, 4.0, 3.15, 0.55, {
    fontSize: 15,
    color: C.navy,
    align: "center",
    valign: "mid"
  });
}

// الماء
slide11.addShape(pptx.ShapeType.roundRect, {
  x: 4.35,
  y: 5.15,
  w: 4.6,
  h: 0.9,
  rectRadius: 0.12,
  line: { color: C.cyan, width: 1.3 },
  fill: { color: C.pale }
});

addText(slide11, "الماء: عنصر أساسي للحفاظ على الترطيب وتوازن السوائل.", 4.5, 5.4, 4.3, 0.4, {
  fontSize: 16,
  bold: true,
  color: C.blue,
  align: "center"
});

// التذييل
addText(slide11, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الثالثة عشرة: الخاتمة والشكر
// ==============================

const slide13 = pptx.addSlide();

slide13.background = { color: C.navy };

// شريط علوي
slide13.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.gold, transparency: 100 },
  fill: { color: C.gold }
});

// رقم الشريحة
addText(slide13, "13", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.white
});

// عنوان الخاتمة
addText(slide13, "الخاتمة", 2.0, 0.85, 9.3, 0.8, {
  fontSize: 34,
  bold: true,
  color: C.gold,
  align: "center"
});

// نص الخاتمة
addText(
  slide13,
  "إن تحقيق الأداء الرياضي المتميز لا يعتمد على التدريب وحده، بل يتطلب تكامل التغذية المتوازنة والترطيب الجيد والراحة والاستشفاء.",
  1.3, 2.0, 10.7, 1.4,
  {
    fontSize: 23,
    bold: true,
    color: C.white,
    align: "center",
    valign: "mid",
    margin: 0.15
  }
);

// عبارة ختامية
addText(
  slide13,
  "التغذية السليمة أساس الأداء الأفضل والحياة الصحية.",
  1.5, 3.85, 10.3, 0.7,
  {
    fontSize: 21,
    color: C.cyan,
    bold: true,
    align: "center"
  }
);

// خط زخرفي
slide13.addShape(pptx.ShapeType.line, {
  x: 4.0,
  y: 5.0,
  w: 5.3,
  h: 0,
  line: { color: C.gold, width: 2 }
});

// الشكر
addText(
  slide13,
  "شكرًا لحُسن استماعكم",
  1.5, 5.35, 10.3, 0.9,
  {
    fontSize: 32,
    bold: true,
    color: C.white,
    align: "center"
  }
);

// التذييل
addText(slide13, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.white
});
async function buildPresentation() {
  await pptx.writeFile({
    fileName: "Sports-Nutrition.pptx"
  });

  console.log("Presentation created successfully!");
}

buildPresentation().catch(console.error);
// التذييل
addText(slide11, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.blue
});
// ==============================
// الشريحة الثالثة عشرة: الخاتمة والشكر
// ==============================

const slide13 = pptx.addSlide();

slide13.background = { color: C.navy };

// شريط علوي
slide13.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: SW, h: 0.16,
  line: { color: C.gold, transparency: 100 },
  fill: { color: C.gold }
});

// رقم الشريحة
addText(slide13, "13", 0.55, 0.45, 1, 0.5, {
  fontSize: 22,
  bold: true,
  color: C.white
});

// عنوان الخاتمة
addText(slide13, "الخاتمة", 2.0, 0.85, 9.3, 0.8, {
  fontSize: 34,
  bold: true,
  color: C.gold,
  align: "center"
});

// نص الخاتمة
addText(
  slide13,
  "إن تحقيق الأداء الرياضي المتميز لا يعتمد على التدريب وحده، بل يتطلب تكامل التغذية المتوازنة والترطيب الجيد والراحة والاستشفاء.",
  1.3, 2.0, 10.7, 1.4,
  {
    fontSize: 23,
    bold: true,
    color: C.white,
    align: "center",
    valign: "mid",
    margin: 0.15
  }
);

// عبارة ختامية
addText(
  slide13,
  "التغذية السليمة أساس الأداء الأفضل والحياة الصحية.",
  1.5, 3.85, 10.3, 0.7,
  {
    fontSize: 21,
    color: C.cyan,
    bold: true,
    align: "center"
  }
);

// خط زخرفي
slide13.addShape(pptx.ShapeType.line, {
  x: 4.0,
  y: 5.0,
  w: 5.3,
  h: 0,
  line: { color: C.gold, width: 2 }
});

// الشكر
addText(
  slide13,
  "شكرًا لحُسن استماعكم",
  1.5, 5.35, 10.3, 0.9,
  {
    fontSize: 32,
    bold: true,
    color: C.white,
    align: "center"
  }
);

// التذييل
addText(slide13, "التغذية الرياضية والأداء الرياضي | 2026", 0.6, 7.05, 6, 0.2, {
  fontSize: 10,
  color: C.white
});
async function buildPresentation() {
  await pptx.writeFile({
    fileName: "Sports-Nutrition.pptx"
  });

  console.log("Presentation created successfully!");
}

buildPresentation().catch(console.error);
