const path = require("path");
const pptxgen = require("pptxgenjs");
const { icon } = require("@fortawesome/fontawesome-svg-core");
const {
  faHammer,
  faThumbsUp,
  faThumbsDown,
  faBagShopping,
  faHeadSideVirus,
} = require("@fortawesome/free-solid-svg-icons");

// Note that we use PowerPoint defaults here ("LAYOUT_WIDE"), which is different from pptxgenjs's
// default ("LAYOUT_16x9") of 10 x 5.625 inches.
const SLIDE_WIDTH = (7.5 / 9) * 16; // inches
const SLIDE_HEIGHT = 7.5; // inches

const FONT_FACE = "Arial";
const FONT_SIZE = {
  PRESENTATION_TITLE: 48,
  PRESENTATION_SUBTITLE: 17,
  SLIDE_TITLE: 32,
  DATE: 17,
  SECTION_TITLE: 18,
  BULLET: 16,
  AGENDA: 24,
  BIG_TEXT: 48,
  DETAIL: 12,
  SECTION: 29,
  QUOTE: 29,
  ATTRIB: 16,
  PLACEHOLDER: 14,
  CITATION: 8,
  SUBHEADER: 28,
};

const CITATION_HEIGHT = calcTextBoxHeight(FONT_SIZE.CITATION);

const MARGINS = {
  DEFAULT_PADDING_BOTTOM: 0.3,
  DEFAULT_CITATION: SLIDE_HEIGHT - CITATION_HEIGHT - 0.15,
  ELEMENT_MEDIUM_PADDING_MEDIUM: 0.4,
  ELEMENT_MEDIUM_PADDING_LARGE: 0.8,
};

const SLIDE_TITLE = { X: 0.4, Y: 0.4, W: "94%" };

const BLACK = "000000";
const WHITE = "FFFFFF";
const NEAR_BLACK_NAVY = "030A18";
const GREYISH_BLUE = "97B1DF";
const LIGHT_GREEN = "A4B6B8";

// If you choose a layout that includes this placeholder, you'll need to replace it with actual
// assets—either generated or sourced from the internet.
const LIGHT_GRAY_BLOCK = path.join(
  __dirname,
  "placeholder_light_gray_block.png"
);

function calcTextBoxHeight(fontSize, lines = 1, leading = 1.2, padding = 0.2) {
  const lineHeightIn = (fontSize / 72) * leading;
  return lines * lineHeightIn + padding;
}

const hSlideTitle = calcTextBoxHeight(FONT_SIZE.SLIDE_TITLE);
function addSlideTitle(slide, title, color = BLACK) {
  slide.addText(title, {
    x: SLIDE_TITLE.X,
    y: SLIDE_TITLE.Y,
    w: SLIDE_TITLE.W,
    h: hSlideTitle,
    fontFace: FONT_FACE,
    fontSize: FONT_SIZE.SLIDE_TITLE,
    color,
  });
}

function getIconSvg(faIcon, color) {
  // CSS color, syntax slightly different from pptxgenjs.
  return icon(faIcon, { styles: { color: `#${color}` } }).html.join("");
}

const svgToDataUri = (svg) =>
  "data:image/svg+xml;base64," + Buffer.from(svg).toString("base64");

(async () => {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";

  // Slide 1: Title slide with subtitle and date
  {
    const slide = pptx.addSlide();

    slide.addImage({
      path: LIGHT_GRAY_BLOCK,
      x: "55%",
      y: 0,
      w: "45%",
      h: "100%",
    });

    const leftMargin = 0.4;
    const hTitle = calcTextBoxHeight(FONT_SIZE.PRESENTATION_TITLE);
    slide.addText("Presentation title", {
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.PRESENTATION_TITLE,
      x: leftMargin,
      y: (SLIDE_HEIGHT - hTitle) / 2,
      w: "50%",
      h: calcTextBoxHeight(FONT_SIZE.PRESENTATION_TITLE),
      valign: "middle",
    });

    slide.addText("Subtitle here", {
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.PRESENTATION_SUBTITLE,
      x: leftMargin,
      y: 4.2,
      w: "50%",
      h: calcTextBoxHeight(FONT_SIZE.PRESENTATION_SUBTITLE),
    });

    let hDate = calcTextBoxHeight(FONT_SIZE.DATE);
    slide.addText("Date here", {
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.DATE,
      x: leftMargin,
      y: SLIDE_HEIGHT - 0.5 - hDate,
      w: 5,
      h: hDate,
    });
  }

  // Slide 2: Stacked column chart with key takeaways
  {
    const slide = pptx.addSlide();

    slide.addImage({
      path: LIGHT_GRAY_BLOCK,
      x: "65%",
      y: 0,
      w: "35%",
      h: "100%",
    });

    addSlideTitle(slide, "Slide title");

    const ySectionTitle = SLIDE_TITLE.Y + hSlideTitle + 0.4;
    const hSectionTitle = calcTextBoxHeight(FONT_SIZE.SECTION_TITLE);
    const yContent = ySectionTitle + hSectionTitle + 0.3;
    slide.addText("Plot Title (units, unit scale)²²", {
      x: 0.4,
      y: ySectionTitle,
      w: "58%",
      h: hSectionTitle,
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.SECTION_TITLE,
    });

    const labels = ["Jan", "Feb", "Mar", "Apr"];
    const dataChart = [
      { name: "Core SaaS", labels, values: [5.2, 5.5, 6.1, 6.4] },
      { name: "Add-ons", labels, values: [1.8, 2.0, 2.5, 2.7] },
      { name: "Services", labels, values: [1.0, 1.1, 1.4, 1.6] },
    ];

    slide.addChart(pptx.ChartType.bar, dataChart, {
      x: SLIDE_TITLE.X,
      y: yContent,
      w: "58%",
      h: 4.0,
      barDir: "col",
      barGrouping: "stacked",
      chartColors: ["002B5B", "1F6BBA", "9CC3E4"],
      showValue: true,
      dataLabelPosition: "ctr",
      dataLabelColor: "DDDDDD",
      catAxisMajorTickMark: "none",
      valAxisMajorTickMark: "none",
      valAxisLineShow: false,
      valGridLine: { color: "EEEEEE" },
      showLegend: true,
      legendPos: "tr",
      legendFontFace: FONT_FACE,
      legendFontSize: 16,
      layout: { x: 0.1, y: 0.1, w: 0.7, h: 1 }, // relative margin of the plot within the chart area.
    });

    const xRight = "67%";
    slide.addText("Takeaways", {
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.SECTION_TITLE,
      x: xRight,
      y: ySectionTitle,
      w: 4,
      h: hSectionTitle,
    });

    slide.addText(
      [
        {
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.",
          options: { bullet: true },
        },
        {
          text: "Nullam pharetra mauris tortor. In hac habitasse platea dictumst.²³",
          options: { bullet: true },
        },
        {
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.",
          options: { bullet: true },
        },
        {
          text: "Nullam pharetra mauris tortor. In hac habitasse platea dictumst.",
          options: { bullet: true },
        },
        {
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.²⁴",
          options: { bullet: true },
        },
      ],
      {
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.BULLET,
        x: xRight,
        y: yContent,
        w: "29%",
        h: calcTextBoxHeight(FONT_SIZE.BULLET, 14),
        paraSpaceAfter: FONT_SIZE.BULLET * 0.3,
      }
    );

    slide.addText(
      [
        {
          text: "[22]",
          options: {
            hyperlink: {
              url: "【67†L9-L11】", // replace with real target
            },
            color: NEAR_BLACK_NAVY,
          },
        },
        {
          text: " ",
          options: { underline: false },
        },
        {
          text: "[23]",
          options: {
            hyperlink: {
              url: "【69†L9-L11】", // replace with real target
            },
            color: NEAR_BLACK_NAVY,
          },
        },
        {
          text: " ",
          options: { underline: false },
        },
        {
          text: "[24]",
          options: {
            hyperlink: {
              url: "【71†L9-L11】", // replace with real target
            },
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 3: Centered big text on blue background
  {
    const slide = pptx.addSlide();
    slide.background = { fill: GREYISH_BLUE };
    const hBigText = calcTextBoxHeight(FONT_SIZE.BIG_TEXT, 4);
    const wBigText = 0.65 * SLIDE_WIDTH;
    slide.addText(
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus.",
      {
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.BIG_TEXT,
        x: (SLIDE_WIDTH - wBigText) / 2,
        y: (SLIDE_HEIGHT - hBigText) / 2,
        w: wBigText,
        h: hBigText,
        align: "center",
        valign: "middle",
      }
    );
  }

  // Slide 4: Big text over image background
  {
    const slide = pptx.addSlide();

    slide.addImage({
      path: LIGHT_GRAY_BLOCK, // Pick a background image with enough contrast to the text.
      x: 0,
      y: 0,
      w: "100%",
      h: "100%",
    });

    const hBigText = calcTextBoxHeight(FONT_SIZE.BIG_TEXT, 4);
    const wBigText = 0.65 * SLIDE_WIDTH;
    slide.addText(
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus.",
      {
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.BIG_TEXT,
        x: (SLIDE_WIDTH - wBigText) / 2,
        y: (SLIDE_HEIGHT - hBigText) / 2,
        w: wBigText,
        h: hBigText,
        align: "center",
        valign: "middle",
      }
    );
  }

  // Slide 5: Dark panel featuring a large quote and image placeholder
  {
    const slide = pptx.addSlide();

    slide.addImage({
      path: LIGHT_GRAY_BLOCK, // Pick a background image with enough contrast to the text.
      x: 0,
      y: 0,
      w: "100%",
      h: "100%",
    });

    const hBigText = 0.8 * SLIDE_HEIGHT;
    const wBigText = hBigText;
    const xPanel = (SLIDE_WIDTH - wBigText) / 2;

    slide.addText(
      "Lorem ipsum dolor sit amet, consectetur ut quam adipiscing ultricies elit.",
      {
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.BIG_TEXT,
        color: WHITE,
        x: xPanel,
        y: (SLIDE_HEIGHT - hBigText) / 2,
        w: wBigText,
        h: hBigText,
        margin: 32,
        valign: "top",
        fill: { color: NEAR_BLACK_NAVY },
      }
    );

    slide.addText(
      [
        {
          text: "[1]",
          options: {
            hyperlink: {
              url: "【1†1050-1150】", // replace with actual citation target
            },
          },
        },
      ],
      {
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
        color: GREYISH_BLUE,
        x: xPanel + 0.3,
        y: (SLIDE_HEIGHT + hBigText) / 2 - 0.2 - CITATION_HEIGHT,
        w: wBigText - 0.4,
        h: CITATION_HEIGHT,
      }
    );
  }

  // Slide 6: Two-column layout—image left, text right
  {
    const slide = pptx.addSlide();

    // Left placeholder image
    slide.addImage({ path: LIGHT_GRAY_BLOCK, x: 0, y: 0, w: "50%", h: "100%" });

    // Right quote
    const RIGHT_X = "52%";
    const RIGHT_W = "45%";
    slide.addText(
      "“Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus.”",
      {
        x: RIGHT_X,
        y: 0.3,
        w: RIGHT_W,
        h: calcTextBoxHeight(FONT_SIZE.QUOTE, 4, 1.25),
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.QUOTE,
        valign: "top",
      }
    );

    // Attribution
    const hAttrib = calcTextBoxHeight(FONT_SIZE.ATTRIB, 2);
    slide.addText("Author name\nJob title", {
      x: RIGHT_X,
      y: SLIDE_HEIGHT - 0.4 - hAttrib,
      w: RIGHT_W,
      h: hAttrib,
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.ATTRIB,
      align: "right",
      valign: "bottom",
    });
  }

  // Slide 7: Two stacked image-text blocks
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Slide title");

    // Two-row image/text layout
    const IMG = { X: "42%", W: "25%", H: 2.8 };
    const TXT = { X: "70%", W: "20%", FS: 13 };
    const ROW_GAP = 0.8;
    const lorem =
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus.";

    const addRow = (y) => {
      slide.addImage({
        path: LIGHT_GRAY_BLOCK,
        x: IMG.X,
        y,
        w: IMG.W,
        h: IMG.H,
      });

      slide.addText(lorem, {
        x: TXT.X,
        y,
        w: TXT.W,
        h: calcTextBoxHeight(TXT.FS, 4),
        fontFace: FONT_FACE,
        fontSize: TXT.FS,
        valign: "top",
      });
    };

    const firstY = SLIDE_TITLE.Y + 0.2;
    addRow(firstY);
    addRow(firstY + IMG.H + ROW_GAP);
  }

  // Slide 8: Three side-by-side image blocks with captions
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Slide title");

    // Three images with captions
    const IMAGE_W = "28%";
    const IMAGE_H = 2.4;
    const LEFT_PCT = 0.05;
    const GAP_PCT = 0.03;
    const IMAGE_Y = (SLIDE_HEIGHT - IMAGE_H) / 2;
    const CAPTION_Y = IMAGE_Y + IMAGE_H + 0.2;

    const xPos = [
      `${LEFT_PCT * 100}%`,
      `${(LEFT_PCT + parseFloat(IMAGE_W) / 100 + GAP_PCT) * 100}%`,
      `${(LEFT_PCT + 2 * (parseFloat(IMAGE_W) / 100 + GAP_PCT)) * 100}%`,
    ];

    const captions = [
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus.",
      "Nullam pharetra mauris tortor. In hac habitasse platea dictumst. Proin mattis nibh risus. In hac habitasse platea dictumst.",
      "Nulla tempor ut massa elementum dapibus. Nam non ante quis enim fringilla tempus nec in lectus.",
    ];

    xPos.forEach((x, i) => {
      slide.addImage({
        path: LIGHT_GRAY_BLOCK,
        x,
        y: IMAGE_Y,
        w: IMAGE_W,
        h: IMAGE_H,
      });
      slide.addText(captions[i], {
        x,
        y: CAPTION_Y,
        w: IMAGE_W,
        h: calcTextBoxHeight(13, 4),
        fontFace: FONT_FACE,
        fontSize: 13,
        valign: "top",
      });
    });
  }

  // Slide 9: Three key points stacked vertically
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Slide title");

    // Numbered text blocks
    const NUMBER = { X: "40%", W: 0.9, FS: 20 };
    const CONTENT = { X: "45%", W: "50%" };
    const TITLE_TEXT_GAP = 0.02;
    const BLOCK_GAP = 0.6;
    let y = SLIDE_TITLE.Y + 0.2;

    const items = [
      {
        num: "01",
        title: "Title for item number one…",
        text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus. Nullam pharetra mauris tortor.² In hac habitasse platea dictumst. Proin mattis nibh risus. Nulla tempor ut massa elementum dapibus. Nam non ante quis enim fringilla tempus nec in lectus.³ In hac habitasse platea dictumst.",
      },
      {
        num: "02",
        title: "Title for item number one…",
        text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus. Nullam pharetra mauris tortor. In hac habitasse platea dictumst.⁵ Proin mattis nibh risus. Nulla tempor ut massa elementum dapibus. Nam non ante quis enim fringilla tempus nec in lectus.⁶ Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus. Nullam pharetra mauris tortor.⁷ In hac habitasse platea dictumst.",
      },
      {
        num: "03",
        title: "Title for item number one…",
        text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus. Nullam pharetra mauris tortor. In hac habitasse platea dictumst.⁸ Proin mattis nibh risus. Nulla tempor ut massa elementum dapibus. Nam non ante quis enim fringilla tempus nec in lectus.⁹ In hac habitasse platea dictumst.¹⁰",
      },
    ];

    items.forEach(({ num, title, text }) => {
      slide.addText(num, {
        x: NUMBER.X,
        y,
        w: NUMBER.W,
        h: calcTextBoxHeight(NUMBER.FS),
        fontFace: FONT_FACE,
        fontSize: NUMBER.FS,
      });

      const hTitle = calcTextBoxHeight(20);
      slide.addText(title, {
        x: CONTENT.X,
        y,
        w: CONTENT.W,
        h: hTitle,
        fontFace: FONT_FACE,
        fontSize: 20,
      });

      const hBody = calcTextBoxHeight(10, 5);
      slide.addText(text, {
        x: CONTENT.X,
        y: y + hTitle + TITLE_TEXT_GAP,
        w: CONTENT.W,
        h: hBody,
        fontFace: FONT_FACE,
        fontSize: 10,
      });

      y += hTitle + TITLE_TEXT_GAP + hBody + BLOCK_GAP;
    });

    // Citations in hyperlink format with random tether IDs and line numbers (e.g., L9-L11)
    slide.addText(
      [
        {
          text: "[2]",
          options: {
            hyperlink: { url: "【8†L9-L11】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[3]",
          options: {
            hyperlink: { url: "【2015010†L12-L14】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[4]",
          options: {
            hyperlink: { url: "【345234535†L15-L17】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[5]",
          options: {
            hyperlink: { url: "【592†L18-L20】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[6]",
          options: {
            hyperlink: { url: "【8691†L21-L23】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[7]",
          options: {
            hyperlink: { url: "【7†L24-L26】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[8]",
          options: {
            hyperlink: { url: "【8†L27-L29】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[9]",
          options: {
            hyperlink: { url: "【9†L30-L32】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[10]",
          options: {
            hyperlink: { url: "【8723†L33-L35】" },
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 10: Three key points in horizontal columns
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Slide title");

    // Three-column text
    const BODY_FS = 17;
    const yCols = SLIDE_TITLE.Y + calcTextBoxHeight(FONT_SIZE.SECTION) + 0.4;
    const TOTAL_W = SLIDE_WIDTH - 2 * SLIDE_TITLE.X;
    const COL_W = 0.3 * SLIDE_WIDTH;
    const GAP_W = (TOTAL_W - 3 * COL_W) / 2;
    const COL = {
      X1: SLIDE_TITLE.X,
      X2: SLIDE_TITLE.X + COL_W + GAP_W,
      X3: SLIDE_TITLE.X + 2 * (COL_W + GAP_W),
    };
    const hColText = calcTextBoxHeight(BODY_FS, 15);

    const lorem =
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus. Nullam pharetra mauris tortor. In hac habitasse platea dictumst. Proin mattis nibh risus. Nulla tempor ut massa elementum dapibus. Nam non ante quis enim fringilla tempus nec in lectus. In hac habitasse platea dictumst.";

    // Columns 1 & 2
    [COL.X1, COL.X2].forEach((x) =>
      slide.addText(lorem, {
        x,
        y: yCols,
        w: COL_W,
        h: hColText,
        fontFace: FONT_FACE,
        fontSize: BODY_FS,
        valign: "top",
      })
    );

    // Column 3
    slide.addText(
      [
        { text: lorem, options: {} },
        {
          text: "In hac habitasse platea dictumst.",
          options: { bullet: true },
        },
        { text: "Proin mattis nibh risus.", options: { bullet: true } },
      ],
      {
        x: COL.X3,
        y: yCols,
        w: COL_W,
        h: hColText,
        fontFace: FONT_FACE,
        fontSize: BODY_FS,
        paraSpaceAfter: BODY_FS * 0.3,
        valign: "top",
      }
    );

    slide.addText(
      [
        {
          text: "[11]",
          options: {
            hyperlink: {
              url: "【60†L9-L11】", // replace with actual citation target
            },
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: calcTextBoxHeight(FONT_SIZE.CITATION),
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 11: Prominent big number with two supporting text columns
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Slide title");

    const BIG_NUM_FS = 96;
    const yBig = SLIDE_TITLE.Y + calcTextBoxHeight(FONT_SIZE.SECTION) + 0.4;
    slide.addText("82%", {
      x: SLIDE_TITLE.X,
      y: yBig,
      w: "35%",
      h: calcTextBoxHeight(BIG_NUM_FS),
      fontFace: FONT_FACE,
      fontSize: BIG_NUM_FS,
      color: "9FB7B9",
    });

    const BODY_FS = 10;
    const loremLong =
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus. Nullam pharetra mauris tortor. In hac habitasse platea dictumst. Proin mattis nibh risus. Nulla tempor ut massa elementum dapibus. Nam non ante quis enim fringilla tempus nec in lectus. In hac habitasse platea dictumst. " +
      "Sed vitae magna sed libero suscipit bibendum. Phasellus aliquam magna eu fringilla aliquet. Morbi hendrerit, lorem a pulvinar consequat, mi mauris viverra felis, sed tempor felis velit a erat. Integer eget orci at tellus fringilla congue. Cras sed justo sed augue eleifend ultricies. In ut justo ut nibh hendrerit gravida. Praesent sit amet purus non urna bibendum auctor. " +
      "Donec feugiat facilisis ipsum, vitae faucibus odio tempus sit amet. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Suspendisse potenti. Maecenas quis porta lorem. Vivamus aliquet metus at mi mattis vestibulum. Integer congue elit vitae nunc dignissim, sed accumsan nibh faucibus.";
    const yBody = yBig;
    const COL_W = "25%";
    const hBody = calcTextBoxHeight(BODY_FS, 29);

    ["42%", "72%"].forEach((x) =>
      slide.addText(loremLong, {
        x,
        y: yBody,
        w: COL_W,
        h: hBody,
        fontFace: FONT_FACE,
        fontSize: BODY_FS,
        valign: "top",
      })
    );

    slide.addText(
      [
        {
          text: "[11]",
          options: {
            hyperlink: { url: "【60†L500-L510】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[12]",
          options: {
            hyperlink: { url: "【5402†L511-L520】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[14]",
          options: {
            hyperlink: { url: "【90017†L521-L530】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[15]",
          options: {
            hyperlink: { url: "【265†L531-L540】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[16]",
          options: {
            hyperlink: { url: "【16†L541-L550】" },
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[17]",
          options: {
            hyperlink: { url: "【17†L551-L560】" },
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 12: Side-by-side comparison tables
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Slide title");

    const lorem =
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus. Nullam pharetra mauris tortor. In hac habitasse platea dictumst. Proin mattis nibh risus. Nulla tempor ut massa elementum dapibus.";

    const makeRow = ({ us, uk, text }) => [
      { text: us, options: { fontFace: FONT_FACE, fontSize: 24 } },
      { text: uk, options: { fontFace: FONT_FACE, fontSize: 24 } },
      { text, options: { fontFace: FONT_FACE, fontSize: 10 } },
    ];

    const headerRow = [
      { text: "US", options: { fontFace: FONT_FACE, fontSize: 14 } },
      { text: "UK", options: { fontFace: FONT_FACE, fontSize: 14 } },
      { text: "", options: { fontFace: FONT_FACE, fontSize: 14 } },
    ];

    const leftRows = [
      headerRow,
      makeRow({ us: "82%", uk: "71%", text: lorem }),
      makeRow({ us: "77%", uk: "68%", text: lorem }),
      makeRow({ us: "80%", uk: "67%", text: lorem }),
    ];

    const rightRows = [
      headerRow,
      makeRow({ us: "71%", uk: "61%", text: lorem }),
      makeRow({ us: "84%", uk: "76%", text: lorem }),
      makeRow({ us: "79%", uk: "68%", text: lorem }),
    ];

    // Layout math
    const CENTER_X = SLIDE_WIDTH / 2;
    const GAP = 0.25;

    const colW = [1.2, 1.2, 3.6];
    const TABLE_W = 6.0;

    const LEFT_X = CENTER_X - GAP - TABLE_W;
    const RIGHT_X = CENTER_X + GAP;
    const yStart = 2.1;

    slide.addTable(leftRows, { x: LEFT_X, y: yStart, colW, rowH: 0.9 });
    slide.addTable(rightRows, { x: RIGHT_X, y: yStart, colW, rowH: 0.9 });

    slide.addShape(pptx.ShapeType.line, {
      x: CENTER_X,
      y: yStart - 0.2,
      w: 0,
      h: 4.8,
      line: { color: "666666", width: 1 },
    });

    slide.addText(
      [
        {
          text: "[17]",
          options: {
            hyperlink: { url: "【17†L223-L225】" }, // replace with real target
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 13: Detailed data table
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Slide title");

    const NONE = { type: "none" };
    const BOTTOM = { color: "000000", pt: 1 };

    // Only a single horizontal rule under the header row
    const headerBorder = [NONE, NONE, BOTTOM, NONE]; // top, right, bottom, left
    const bodyBorder = [NONE, NONE, NONE, NONE];

    const mkHeader = (t) => ({
      text: t,
      options: {
        fontFace: FONT_FACE,
        fontSize: 15,
        fill: "EDEDED",
        border: headerBorder,
        valign: "bottom",
      },
    });

    const mkCell = (t) => ({
      text: t,
      options: {
        fontFace: FONT_FACE,
        fontSize: 15,
        border: bodyBorder,
        valign: "bottom",
      },
    });

    // ── Table rows ────────────────────────────────────────────────────────────
    const dataRows = [
      ["Item 1", "7 years", "$100", "5%", "5", "high"],
      ["Item 2", "3 years", "$500", "6%", "2", "high"],
      ["Item 3", "6 years", "$1,500", "10%", "1", "low"],
      ["Item 4", "6 years", "$1,900", "25%", "4", "medium"],
      ["Item 5", "11 years", "$2,500", "90%", "6", "high"],
      ["Item 6", "3 years", "$50", "14%", "0", "n/a"],
      ["Item 7", "4", "$9,000", "25%", "9", "medium"],
      ["Item 8", "6 years", "$2,500", "3%", "7", "low"],
    ];

    const rows = [
      [
        mkHeader(""),
        mkHeader("🍎 Column 1 ¹⁷"),
        mkHeader("🗑️ Column 2 ¹⁸"),
        mkHeader("🖨️ Column 3 ¹⁹"),
        mkHeader("⛶ Column 4 ²⁰"),
        mkHeader("🔔 Column 5 ²¹"),
      ],
      ...dataRows.map((r) => r.map(mkCell)),
    ];

    const COL_W = Array(6).fill(2.1);
    const TABLE_W = COL_W.reduce((s, w) => s + w, 0);
    const xStart = (SLIDE_WIDTH - TABLE_W) / 2;
    const yStart = 1.6;

    slide.addTable(rows, {
      x: xStart,
      y: yStart,
      colW: COL_W,
      rowH: 0.55,
    });

    slide.addText(
      [
        {
          text: "[17]",
          options: {
            hyperlink: { url: "【17†L101-L103】" },
            underline: true,
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[18]",
          options: {
            hyperlink: { url: "【18†L104-L106】" },
            underline: true,
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[19]",
          options: {
            hyperlink: { url: "【25†L107-L109】" },
            underline: true,
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[20]",
          options: {
            hyperlink: { url: "【12345†L110-L112】" },
            underline: true,
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[21]",
          options: {
            hyperlink: { url: "【51†L113-L115】" },
            underline: true,
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 14: Three icon-feature blocks on black background
  {
    const slide = pptx.addSlide();
    slide.background = { fill: BLACK };

    addSlideTitle(slide, "Slide title", WHITE);

    const ICON_W = 0.4;
    const ICON_H = 0.4;
    const TEXT_BELOW_ICON_W = 0.8;
    const GAP_ICON_TEXT = 0.25;
    const GAP_ICON_TITLE = 0.15;
    const TEXT_W = 2.6;
    const TITLE_FS = 14;
    const BODY_FS = 12;

    const GROUP_W = TEXT_BELOW_ICON_W + GAP_ICON_TEXT + TEXT_W;
    const GAP_GROUPS = 0.9;
    const TOTAL_W = GROUP_W * 3 + GAP_GROUPS * 2;
    const LEFT_X = (SLIDE_WIDTH - TOTAL_W) / 2;

    const hTitle = calcTextBoxHeight(TITLE_FS);
    const hBody = calcTextBoxHeight(BODY_FS, 16, 1.4);
    const groupH = Math.max(ICON_H, hBody) + GAP_ICON_TITLE + hTitle;
    const groupY = (SLIDE_HEIGHT - groupH) / 2;

    const colXs = [
      LEFT_X,
      LEFT_X + GROUP_W + GAP_GROUPS,
      LEFT_X + 2 * (GROUP_W + GAP_GROUPS),
    ];

    const bodyPara =
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus. Nullam pharetra mauris tortor. In hac habitasse platea dictumst. Proin mattis nibh risus.";

    const iconColor = "EEEEEE";

    const icons = [
      getIconSvg(faHammer, iconColor),
      getIconSvg(faThumbsUp, iconColor),
      getIconSvg(faThumbsDown, iconColor),
    ];

    colXs.forEach((x) => {
      const iconX = x;
      const titleX = x;
      const bodyX = x + TEXT_BELOW_ICON_W + GAP_ICON_TEXT;

      slide.addImage({
        data: svgToDataUri(icons[colXs.indexOf(x)]),
        x: iconX + (TEXT_BELOW_ICON_W - ICON_W) / 2,
        y: groupY,
        w: ICON_W,
        h: ICON_H,
      });

      slide.addText("Add title here", {
        x: titleX,
        y: groupY + ICON_H + GAP_ICON_TITLE,
        w: TEXT_BELOW_ICON_W,
        h: hTitle,
        fontFace: FONT_FACE,
        fontSize: TITLE_FS,
        color: "FFFFFF",
        align: "center",
        valign: "top",
      });

      slide.addText(
        [
          { text: bodyPara, options: {} },
          {
            text: "In hac habitasse platea dictumst.",
            options: { bullet: true },
          },
          { text: "Proin mattis nibh risus.", options: { bullet: true } },
        ],
        {
          x: bodyX,
          y: groupY,
          w: TEXT_W,
          h: hBody,
          fontFace: FONT_FACE,
          fontSize: BODY_FS,
          color: "FFFFFF",
          paraSpaceAfter: BODY_FS * 0.3,
          lineSpacingMultiple: 1.4,
          valign: "top",
        }
      );
    });

    slide.addText(
      [
        {
          text: "[19]",
          options: {
            hyperlink: { url: "【25†L600-L610】" }, // replace with actual citation target
            color: "FFFFFF",
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 15: Two columns with icons
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Slide title");

    // Layout constants
    const PADDING_TOP = 0.4;
    const PADDING_BOTTOM = 0.3;
    const PADDING_HORIZONTAL = 0.2;
    const SUBTITLE_W = 1.3;
    const CONTENT_W = 4.5;
    const TEXT_GAP = 0.1;
    const MEDIUM_ICON_SIZE = 0.4;

    const yStart = SLIDE_TITLE.Y + hSlideTitle + PADDING_TOP;
    const yEnd =
      SLIDE_HEIGHT - calcTextBoxHeight(FONT_SIZE.CITATION) - PADDING_BOTTOM;

    const cols = [
      {
        xLabel: SLIDE_TITLE.X,
        xContent: SLIDE_TITLE.X + SUBTITLE_W + PADDING_HORIZONTAL,
      },
      {
        xLabel: SLIDE_WIDTH / 2,
        xContent: SLIDE_WIDTH / 2 + SUBTITLE_W + PADDING_HORIZONTAL,
      },
    ];

    // Draw separators
    cols.forEach(({ xContent }) => {
      slide.addShape(pptx.ShapeType.line, {
        x: xContent - PADDING_HORIZONTAL,
        y: yStart,
        w: 0,
        h: yEnd - yStart,
        line: { color: "CCCCCC", pt: 1 },
      });
    });

    function addMetrics({ xLabel, xContent }, metrics, labelText, svgIcon) {
      // Label text (tight height)
      // Center the icon horizontally above the label text
      const iconX =
        xLabel + (SUBTITLE_W - MEDIUM_ICON_SIZE) / 2 - MEDIUM_ICON_SIZE / 2;
      const svgIconObject = icon(svgIcon, {
        styles: { color: `#${LIGHT_GREEN}` },
      }).html.join(""); // Color must be added here.
      slide.addImage({
        data: svgToDataUri(svgIconObject),
        x: iconX,
        y: yStart + TEXT_GAP,
        w: MEDIUM_ICON_SIZE,
        h: MEDIUM_ICON_SIZE,
      });

      slide.addText(labelText, {
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.DETAIL,
        x: xLabel,
        y: yStart + TEXT_GAP + MEDIUM_ICON_SIZE * 2,
        w: SUBTITLE_W,
        h: calcTextBoxHeight(FONT_SIZE.DETAIL),
        wrap: true,
      });

      let y =
        yStart +
        calcTextBoxHeight(FONT_SIZE.DETAIL, (leading = 1), (padding = 0.01));
      metrics.forEach(({ value, text, lines }) => {
        // Metric value
        const hSubheader = calcTextBoxHeight(FONT_SIZE.SUBHEADER);
        slide.addText(value, {
          fontFace: FONT_FACE,
          fontSize: FONT_SIZE.SUBHEADER,
          color: LIGHT_GREEN,
          x: xContent,
          y,
          w: CONTENT_W,
          h: hSubheader,
        });
        y += hSubheader + TEXT_GAP;

        // Detail text
        const hDetail = calcTextBoxHeight(FONT_SIZE.DETAIL, lines);
        slide.addText(text, {
          fontFace: FONT_FACE,
          fontSize: FONT_SIZE.DETAIL,
          x: xContent,
          y,
          w: CONTENT_W,
          h: hDetail,
          wrap: true,
        });
        // Increment by line height * lines + bottom padding
        y += hDetail + PADDING_BOTTOM;
      });
    }

    // Populate columns
    addMetrics(
      cols[0],
      [
        {
          value: "$5B",
          text: "In hac habitasse platea dictumst. Proin mattis nibh risus. Nulla tempor ut massa elementum dapibus. Nam non ante quis enim fringilla tempus nec in lectus. In hac habitasse platea.",
          lines: 3,
        },
        {
          value: "3x",
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus. Nullam pharetra mauris tortor.",
          lines: 2,
        },
        {
          value: "10x",
          text: "Proin mattis nibh risus. Nulla tempor ut massa elementum dapibus dolor sit.",
          lines: 2,
        },
      ],
      "Title here…",
      faBagShopping
    );

    addMetrics(
      cols[1],
      [
        {
          value: "100s",
          text: "In hac habitasse platea dictumst. Proin mattis nibh risus. Nulla tempor ut massa elementum dapibus. Nam non ante quis enim fringilla tempus nec in lectus. In hac habitasse platea.",
          lines: 3,
        },
        {
          value: "+84%",
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus. Nullam pharetra mauris tortor.",
          lines: 2,
        },
        {
          value: "-25%",
          text: "Proin mattis nibh risus. Nulla tempor ut massa elementum dapibus dolor sit.",
          lines: 2,
        },
      ],
      "Title here…",
      faHeadSideVirus
    );

    // Footer citation
    slide.addText(
      [
        {
          text: "[17]",
          options: {
            hyperlink: { url: "【17†L301-L313】" }, // replace with real target
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 16: Four-up statistics grid
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Slide title");

    // 4-up stat layout
    const STAT_W = "20%";
    const COL_X = ["48%", "77%"]; // wider gap — left column farther left
    const firstY = SLIDE_TITLE.Y; // align stats' top with title
    const ROW_GAP = 2.3;
    const BIG_FS = 48;
    const SMALL_FS = 13;
    const lorem =
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.";

    const addStat = (x, y) => {
      slide.addText("99.99%", {
        x,
        y,
        w: STAT_W,
        h: calcTextBoxHeight(BIG_FS),
        fontFace: FONT_FACE,
        fontSize: BIG_FS,
        color: "9FB7B9",
      });

      const yTitle = y + calcTextBoxHeight(BIG_FS) + 0.05;
      slide.addText("Add your title here", {
        x,
        y: yTitle,
        w: STAT_W,
        h: calcTextBoxHeight(SMALL_FS),
        fontFace: FONT_FACE,
        fontSize: SMALL_FS,
      });

      slide.addText(lorem, {
        x,
        y: yTitle + calcTextBoxHeight(SMALL_FS),
        w: STAT_W,
        h: calcTextBoxHeight(SMALL_FS, 3),
        fontFace: FONT_FACE,
        fontSize: SMALL_FS,
      });
    };

    addStat(COL_X[0], firstY);
    addStat(COL_X[1], firstY);
    addStat(COL_X[0], firstY + ROW_GAP);
    addStat(COL_X[1], firstY + ROW_GAP);

    // Sources
    slide.addText(
      [
        {
          text: "[17]",
          options: {
            hyperlink: { url: "【17†L999-L1001】" }, // replace with actual citation target
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 17: Horizontal six-item table layout
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Horizontal six item table layout");

    const lorem = [
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus.",
      "Nullam pharetra mauris tortor. In hac habitasse platea dictumst. Proin mattis nibh risus. In hac habitasse platea dictumst.",
      "Nulla tempor ut massa elementum dapibus. Nam non ante quis enim fringilla tempus nec in lectus.",
    ];

    const makeCell = (idx) => ({
      text: [
        {
          text: "Title\n",
          options: { fontFace: FONT_FACE, fontSize: 13, paraSpaceAfter: 24 },
        },
        { text: lorem[idx], options: { fontFace: FONT_FACE, fontSize: 18 } },
      ],
      options: {
        valign: "top",
        border: [
          { type: "none" },
          { type: "none" },
          { type: "none" },
          { type: "none" },
        ],
      },
    });

    const rows = [
      [makeCell(0), makeCell(1), makeCell(2)],
      [makeCell(0), makeCell(1), makeCell(2)],
    ];

    const COL_W = Array(3).fill(4.2);
    const TABLE_W = COL_W.reduce((s, w) => s + w, 0);
    const xStart = (SLIDE_WIDTH - TABLE_W) / 2;

    const yStart = SLIDE_TITLE.Y + calcTextBoxHeight(FONT_SIZE.SECTION) + 0.9;

    slide.addTable(rows, {
      x: xStart,
      y: yStart,
      colW: COL_W,
      rowH: 2.4,
    });

    slide.addText(
      [
        {
          text: "[9]",
          options: {
            hyperlink: { url: "【9†L30-L32】" }, // replace with real target if needed
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 18: Timeline with chevron steps and category table
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Timeline or Phases");

    // Geometry & positioning constants
    const STEP_TITLES = [
      "Step title 1",
      "Step title 2",
      "Step title 3",
      "Step title 4",
    ];
    const STEP_W = 2.55; // matches table-column width
    const STEP_H = 0.55; // a little sleeker
    const CAT_COL_W = 1.7; // width of the Category column
    const LEFT_MARGIN = 0.4; // overall slide margin
    const ARROW_Y = SLIDE_TITLE.Y + 1.05;
    const TABLE_X = LEFT_MARGIN;
    const TABLE_Y = ARROW_Y + STEP_H + 0.25;
    const COL_W = [CAT_COL_W, STEP_W, STEP_W, STEP_W, STEP_W];

    // Arrow bar
    STEP_TITLES.forEach((label, i) => {
      slide.addText(label, {
        shape: pptx.ShapeType.chevron,
        x: TABLE_X + CAT_COL_W + i * STEP_W,
        y: ARROW_Y,
        w: STEP_W,
        h: STEP_H,
        fill: { color: "F2F2F2" },
        line: { color: "F2F2F2" },
        fontFace: FONT_FACE,
        fontSize: 16,
        align: "center",
        valign: "middle",
      });
    });

    // Bullet helper
    const bullet = (txt) => ({
      text: txt,
      options: { bullet: true, fontFace: FONT_FACE, fontSize: 14 },
    });

    // Content table
    const longA =
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus.";
    const longB =
      "Nullam pharetra mauris tortor. In hac habitasse platea dictumst.";

    const rows = [
      [
        { text: "Category 1", options: { fontFace: FONT_FACE, fontSize: 16 } },
        { text: [bullet(longA), bullet("Pellentesque ultricies quam ut.³⁹")] },
        { text: [bullet(longA), bullet("Nullam pharetra mauris tortor.⁴⁰")] },
        { text: [bullet(longA), bullet("Pellentesque ultricies quam ut.")] },
        {
          text: [bullet("Lorem ipsum dolor sit amet. " + longA), bullet(longB)],
        },
      ],
      [
        { text: "Category 2", options: { fontFace: FONT_FACE, fontSize: 16 } },
        {
          text: [
            bullet(longA.replace("consectetur ", "elit ")),
            bullet("Pellentesque ultricies quam ut."),
          ],
        },
        {
          text: [
            bullet("Lorem ipsum dolor sit amet, consectetur."),
            bullet("Pellentesque ultricies quam ut."),
          ],
        },
        {
          text: [
            bullet("Lorem ipsum dolor sit amet, consectetur."),
            bullet("Pellentesque ultricies quam ut."),
          ],
        },
        {
          text: [
            bullet("Lorem ipsum dolor sit amet, consectetur."),
            bullet("Pellentesque ultricies quam ut."),
          ],
        },
      ],
    ];

    slide.addTable(rows, {
      x: TABLE_X,
      y: TABLE_Y,
      colW: COL_W,
      rowH: 1.9,
      valign: "top",
      border: [
        { type: "none" },
        { type: "none" },
        { type: "none" },
        { type: "none" },
      ],
    });

    // Sources
    slide.addText(
      [
        {
          text: "[39]",
          options: {
            hyperlink: { url: "【3†L600-L610】" }, // replace with actual citation target
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[40]",
          options: {
            hyperlink: { url: "【102†L600-L610】" }, // replace with actual citation target
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 19: Category comparison data table
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Table and Categories");

    // helper builders
    const fs15 = (t) => ({
      text: t,
      options: { fontFace: FONT_FACE, fontSize: 15 },
    });
    const fs16 = (t) => ({
      text: t,
      options: { fontFace: FONT_FACE, fontSize: 16 },
    });
    const bullets = (...lines) => ({
      text: lines.map((l) => ({ text: l, options: { bullet: true } })),
      options: { fontFace: FONT_FACE, fontSize: 15, lineSpacingMultiple: 1.15 },
    });

    // table rows
    const rows = [
      [
        fs16(""),
        fs16("Category 1"),
        fs16("Category 2"),
        fs16("Category 3"),
        fs16("Category 4"),
      ],
      [
        fs16("Item 1"),
        fs15("25%¹⁰"),
        fs15("$500,000¹³"),
        bullets(
          "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.",
          "Nullam pharetra mauris tortor. In hac habitasse platea dictumst."
        ),
        fs15(
          "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam pharetra mauris tortor. In hac habitasse platea dictumst. Nullam pharetra mauris tortor."
        ),
      ],
      [
        fs16("Item 1"),
        fs15("25%¹¹"),
        fs15("$500,000¹⁴"),
        bullets(
          "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.",
          "Nullam pharetra mauris tortor."
        ),
        fs15(
          "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam pharetra mauris tortor."
        ),
      ],
      [
        fs16("Item 1"),
        fs15("25%¹²"),
        fs15("$500,000¹⁵"),
        bullets(
          "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
          "Nullam pharetra mauris tortor. In hac habitasse platea dictumst. Pellentesque ultricies quam ut."
        ),
        fs15(
          "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam pharetra mauris tortor. Lorem ipsum dolor sit amet, consectetur adipiscing elit."
        ),
      ],
    ];

    // sizing & placement
    const COL_W = [1.4, 1.5, 2.1, 4.1, 3.7];
    const TABLE_W = COL_W.reduce((s, w) => s + w, 0);
    const xStart = (13.33 - TABLE_W) / 2;
    const yStart = SLIDE_TITLE.Y + calcTextBoxHeight(FONT_SIZE.SECTION) + 0.4;

    slide.addTable(rows, {
      x: xStart,
      y: yStart,
      colW: COL_W,
      rowH: [0.5, 1.6, 1.6, 1.6],
      valign: "top",
    });

    // sources
    slide.addText(
      [
        {
          text: "[10]",
          options: {
            hyperlink: { url: "【8723†L100-L110】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[11]",
          options: {
            hyperlink: { url: "【60†L111-L120】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[12]",
          options: {
            hyperlink: { url: "【5402†L121-L130】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[13]",
          options: {
            hyperlink: { url: "【17834†L131-L140】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[14]",
          options: {
            hyperlink: { url: "【90017†L141-L150】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[15]",
          options: {
            hyperlink: { url: "【265†L151-L160】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 20: Donut chart with 4 quadrants
  {
    const slide = pptx.addSlide();

    addSlideTitle(slide, "Slide title");

    const cX = 6.4;
    const cY = 3.9;
    const outerR = 2.0;
    const innerR = 1.0;
    const colors = {
      TOP_LEFT: "0B0F1A",
      TOP_RIGHT: "9CB6E1",
      BOTTOM_RIGHT: "495568",
      BOTTOM_LEFT: "A0BEC2",
    };

    // === Donut chart (4 equal quadrants) ===
    const dataSeries = [
      {
        name: "",
        labels: ["TL", "TR", "BR", "BL"],
        values: [1, 1, 1, 1],
      },
    ];
    slide.addChart(pptx.ChartType.doughnut, dataSeries, {
      x: cX - outerR,
      y: cY - outerR,
      w: outerR * 2,
      h: outerR * 2,
      hole: Math.round((innerR / outerR) * 100), // inner hole %
      chartColors: [
        colors.TOP_LEFT,
        colors.TOP_RIGHT,
        colors.BOTTOM_RIGHT,
        colors.BOTTOM_LEFT,
      ],
      showLegend: false,
      showDataLabels: false,
    });

    // === Center hole overlay & label ===
    // (optional if your slide background is white,
    //  but ensures a crisp circle behind the icon)
    slide.addShape(pptx.ShapeType.ellipse, {
      x: cX - innerR,
      y: cY - innerR,
      w: innerR * 2,
      h: innerR * 2,
      fill: { color: WHITE },
      line: { color: WHITE, width: 0 },
    });
    slide.addText("Placeholder Image", {
      x: cX - innerR,
      y: cY - innerR,
      w: innerR * 2,
      h: innerR * 2,
      align: "center",
      valign: "middle",
      fontFace: FONT_FACE,
      fontSize: 14,
      color: BLACK,
    });

    // === Body text boxes ===
    const textBody = [
      "Lorem ipsum dolor sit amet,",
      "consectetur adipiscing elit.",
      "Pellentesque ultricies quam ut",
      "massa luctus cursus.",
    ].join("\n");
    const textProps = {
      fontFace: FONT_FACE,
      fontSize: 16,
      color: BLACK,
      w: 4.0,
      h: 1.9,
      align: "left",
    };
    [
      { x: 0.4, y: 1.6 }, // TL
      { x: 9.6, y: 1.6 }, // TR
      { x: 0.4, y: 5.333333333 }, // BL
      { x: 9.6, y: 5.333333333 }, // BR
    ].forEach((pos) => {
      slide.addText(textBody, { ...textProps, x: pos.x, y: pos.y });
    });

    // === Footer ===
    slide.addText(
      [
        {
          text: "[5]",
          options: {
            hyperlink: { url: "【592†L999-L1001】" }, // replace with actual citation target
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 21: Three-level flowchart with top, mid, and bottom nodes
  {
    let slide = pptx.addSlide();

    // === Constants ===
    const TOP_MID_NODE_WIDTH = 3.3333333;
    const NODE_HEIGHT = 0.93333333;
    const BOT_NODE_WIDTH = 1.05;
    const BOT_NODE_SPACING = (TOP_MID_NODE_WIDTH - 3 * BOT_NODE_WIDTH) / 2;
    const BOT_NODE_HEIGHT = 1.0666667;
    const COLOR_TOP = "0B0F1A";
    const COLOR_MID = "A0BEC2";
    const COLOR_BOTTOM = "A6C1EE";
    const LINE_COLOR = "000000";
    const CENTER_X = 6.6666667;

    const TOP_Y = 1.6;
    const MID_Y = 3.7333333;
    const BOT_Y = 5.6;

    addSlideTitle(slide, "Slide title");

    // === Top node ===
    const topX = CENTER_X - TOP_MID_NODE_WIDTH / 2;
    slide.addText("Add text here", {
      x: topX,
      y: TOP_Y,
      w: TOP_MID_NODE_WIDTH,
      h: NODE_HEIGHT,
      align: "center",
      valign: "mid",
      fontFace: FONT_FACE,
      fontSize: 16,
      color: WHITE,
      fill: { color: COLOR_TOP },
      line: { color: COLOR_TOP },
    });

    // === Mid nodes ===
    const MID_NODE_SPACING = 2.0; // Increase spacing between mid nodes
    const midXs = [
      CENTER_X - MID_NODE_SPACING - TOP_MID_NODE_WIDTH, // left
      CENTER_X - TOP_MID_NODE_WIDTH / 2, // center
      CENTER_X + MID_NODE_SPACING, // right
    ];

    midXs.forEach((x) => {
      slide.addText("Add text here", {
        x,
        y: MID_Y,
        w: TOP_MID_NODE_WIDTH,
        h: NODE_HEIGHT,
        align: "center",
        valign: "mid",
        fontFace: FONT_FACE,
        fontSize: 16,
        color: BLACK,
        fill: { color: COLOR_MID },
        line: { color: COLOR_MID },
      });
    });

    // === Bottom nodes ===
    midXs.forEach((midX) => {
      // Compute base x for the 3 children of this mid node
      const baseX = midX;
      for (let i = 0; i < 3; i++) {
        const x = baseX + i * (BOT_NODE_WIDTH + BOT_NODE_SPACING);
        slide.addText("Add text here", {
          x,
          y: BOT_Y,
          w: BOT_NODE_WIDTH,
          h: BOT_NODE_HEIGHT,
          align: "center",
          valign: "mid",
          fontFace: FONT_FACE,
          fontSize: 16,
          color: BLACK,
          fill: { color: COLOR_BOTTOM },
          line: { color: COLOR_BOTTOM },
        });
      }
    });

    // === CONNECT TOP → MID ===
    const topAnchorX = CENTER_X; // bottom-center of top box
    const topAnchorY = TOP_Y + NODE_HEIGHT; // bottom edge of top box

    midXs.forEach((midX) => {
      const midCenterX = midX + TOP_MID_NODE_WIDTH / 2;
      const midTopY = MID_Y;

      // Compute the bounding box of the line from (topAnchorX, topAnchorY)
      // to (midCenterX, midTopY)
      const x = Math.min(topAnchorX, midCenterX);
      const y = topAnchorY;
      const w = Math.abs(midCenterX - topAnchorX);
      const h = midTopY - topAnchorY;

      slide.addShape(pptx.ShapeType.line, {
        x,
        y,
        w,
        h,
        line: { color: LINE_COLOR },
        flipH: midCenterX < topAnchorX, // flip only if it needs to go left
      });
    });

    // === CONNECT MID → BOTTOM (middle child only) ===
    const bottomTopY = BOT_Y;
    midXs.forEach((midX) => {
      const midCenterX = midX + TOP_MID_NODE_WIDTH / 2;
      const midBaseY = MID_Y + NODE_HEIGHT;

      // === Connect each mid to its *middle* child only (3 lines total)
      midXs.forEach((midX) => {
        // parent anchor: bottom-center of the mid box
        const startX = midX + TOP_MID_NODE_WIDTH / 2;
        const startY = MID_Y + NODE_HEIGHT;

        // only the middle child (j = 1)
        const botX = midX + BOT_NODE_WIDTH + BOT_NODE_SPACING;
        const endX = botX + BOT_NODE_WIDTH / 2; // top-center of that bottom box
        const endY = BOT_Y;

        // compute the bounding box for the diagonal line
        const x = Math.min(startX, endX);
        const y = Math.min(startY, endY);
        const w = Math.abs(endX - startX);
        const h = Math.abs(endY - startY);

        // flipH if it's a left-going connector
        const opts = { x, y, w, h, line: { color: LINE_COLOR } };
        if (endX < startX) opts.flipH = true;

        slide.addShape(pptx.ShapeType.line, opts);
      });
    });

    // === Footer ===
    slide.addText(
      [
        {
          text: "[5]",
          options: {
            hyperlink: { url: "【592†L999-L1001】" }, // replace with actual citation target
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 22: Competitor analysis bubble graph
  {
    const slide = pptx.addSlide();

    slide.addImage({
      path: LIGHT_GRAY_BLOCK,
      x: "65%",
      y: 0,
      w: "35%",
      h: "100%",
    });

    addSlideTitle(slide, "Slide title");

    const haxisLength =
      SLIDE_WIDTH * 0.65 - MARGINS.ELEMENT_MEDIUM_PADDING_LARGE * 2;
    slide.addText("Visualization Title", {
      x: 0.7,
      y: 1.6,
      w: haxisLength,
      h: calcTextBoxHeight(FONT_SIZE.SECTION_TITLE),
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.SECTION_TITLE,
      color: NEAR_BLACK_NAVY,
    });

    // 3) Draw axes
    const axisCenterX = (SLIDE_WIDTH * 0.65) / 2;
    const axisCenterY = SLIDE_HEIGHT / 2 + 0.25;
    const axisLength = SLIDE_HEIGHT * 0.55;

    // Vertical axis
    slide.addShape("line", {
      x: axisCenterX,
      y: axisCenterY / 2,
      w: 0,
      h: axisLength,
      line: { color: NEAR_BLACK_NAVY, width: 1 },
    });

    // Horizontal axis
    slide.addShape("line", {
      x: axisCenterX - haxisLength / 2,
      y: axisCenterY,
      w: haxisLength,
      h: 0,
      line: { color: NEAR_BLACK_NAVY, width: 1 },
    });

    // 4) Axis labels
    slide.addText("PLACEHOLDER X-Axis title", {
      x: axisCenterX - haxisLength / 2,
      y: axisCenterY + axisLength / 2 + MARGINS.ELEMENT_MEDIUM_PADDING_MEDIUM,
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.DETAIL,
      color: NEAR_BLACK_NAVY,
      h: calcTextBoxHeight(FONT_SIZE.DETAIL),
      align: "center",
      w: haxisLength,
    });

    const yAxisTextBoxLength = SLIDE_HEIGHT / 2;
    slide.addText("PLACEHOLDER Y-Axis title", {
      x: -(yAxisTextBoxLength / 2) + SLIDE_TITLE.X,
      y: SLIDE_HEIGHT / 2 + 0.25,
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.DETAIL,
      color: NEAR_BLACK_NAVY,
      rotate: 270,
      w: yAxisTextBoxLength,
      h: calcTextBoxHeight(FONT_SIZE.DETAIL),
      align: "center",
    });

    // 5) Circles with items
    const circleDiameter = 1;
    const circleRadius = circleDiameter / 2;

    const circles = [
      { label: "Item 1", x: axisCenterX - 1.6, y: axisCenterY - 1.6 },
      { label: "Item 2", x: axisCenterX + 0.8, y: axisCenterY - 1.2 },
      { label: "Item 3", x: axisCenterX - 2.0, y: axisCenterY + 0.7 },
      { label: "Item 4", x: axisCenterX + 1.2, y: axisCenterY + 0.5 },
    ];

    circles.forEach((c) => {
      slide.addShape("ellipse", {
        x: c.x,
        y: c.y,
        w: circleDiameter,
        h: circleDiameter,
        fill: { color: "0B3556" }, // dark blue
        line: { color: "0B3556" },
      });

      const hDetail = calcTextBoxHeight(FONT_SIZE.DETAIL);
      slide.addText(c.label, {
        x: c.x,
        y: c.y + (circleDiameter - hDetail) / 2,
        w: circleDiameter,
        h: hDetail,
        align: "center",
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.DETAIL,
        color: WHITE,
      });
    });

    const ySectionTitle = SLIDE_TITLE.Y + hSlideTitle + 0.4;
    const hSectionTitle = calcTextBoxHeight(FONT_SIZE.SECTION_TITLE);
    const yContent = ySectionTitle + hSectionTitle + 0.3;
    const xRight = "67%";
    slide.addText("Takeaways", {
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.SECTION_TITLE,
      x: xRight,
      y: ySectionTitle,
      w: 4,
      h: hSectionTitle,
    });

    slide.addText(
      [
        {
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.",
          options: { bullet: true },
        },
        {
          text: "Nullam pharetra mauris tortor. In hac habitasse platea dictumst.",
          options: { bullet: true },
        },
        {
          text: "25",
          options: {
            superscript: true,
          },
        },
        {
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.",
          options: { bullet: true },
        },
        {
          text: "Nullam pharetra mauris tortor. In hac habitasse platea dictumst.",
          options: { bullet: true },
        },
        {
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.",
          options: { bullet: true },
        },
        {
          text: "26",
          options: {
            superscript: true,
          },
        },
      ],
      {
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.BULLET,
        x: xRight,
        y: yContent,
        w: "29%",
        h: calcTextBoxHeight(FONT_SIZE.BULLET, 14),
        paraSpaceAfter: FONT_SIZE.BULLET * 0.3,
      }
    );

    // 9) Footer
    slide.addText(
      [
        {
          text: "[25]",
          options: {
            hyperlink: { url: "【12345†L250-L260】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[26]",
          options: {
            hyperlink: { url: "【67890†L261-L270】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 23: Table with takeaways
  {
    let slide = pptx.addSlide();

    addSlideTitle(slide, "Slide title");

    const ySectionTitle = SLIDE_TITLE.Y + hSlideTitle + 0.4;
    const hSectionTitle = calcTextBoxHeight(FONT_SIZE.SECTION_TITLE);
    const yContent = ySectionTitle + hSectionTitle + 0.3;
    const wTable = "60%";
    const xRight = "67%";
    slide.addText("Slide title", {
      x: SLIDE_TITLE.X,
      y: SLIDE_TITLE.Y,
      w: SLIDE_TITLE.W,
      h: hSlideTitle,
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.SLIDE_TITLE,
    });

    const tableColHeaders = ["Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024"];
    const colX = [2.66667, 4.133333, 5.6, 7.0666667];

    slide.addText("Table Title", {
      x: SLIDE_TITLE.X,
      y: ySectionTitle,
      w: wTable,
      h: hSectionTitle,
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.SECTION_TITLE,
      color: NEAR_BLACK_NAVY,
    });

    const hTableHeader = calcTextBoxHeight(16);
    tableColHeaders.forEach((text, i) => {
      slide.addText(text, {
        x: colX[i],
        y: yContent,
        w: 1.5,
        h: hTableHeader,
        fontFace: FONT_FACE,
        fontSize: 16,
        color: BLACK,
      });
    });

    slide.addShape("line", {
      x: SLIDE_TITLE.X,
      y: yContent + hTableHeader + 0.05,
      w: wTable,
      h: 0.01,
      line: { color: BLACK, width: 1 },
    });

    const metrics = [
      { label: "Metric 1", sup: "27", values: ["12.4", "9.8", "8.2", "6.1"] },
      { label: "Metric 2", sup: "27", values: ["24.5", "21", "19.3", "16.2"] },
      { label: "Metric 3", sup: "28", values: ["61%", "68%", "72%", "81%"] },
      { label: "Metric 4", sup: "28", values: ["3.9", "4.1", "4.3", "4.5"] },
    ];

    const rowStartY = yContent + hTableHeader + 0.1;
    const rowSpacing = 0.9;

    metrics.forEach((row, idx) => {
      const y = rowStartY + idx * rowSpacing;

      slide.addText(
        [
          { text: row.label },
          row.sup && { text: row.sup, options: { superscript: true } },
        ].filter(Boolean),
        {
          x: SLIDE_TITLE.X,
          y: y,
          w: 2,
          h: calcTextBoxHeight(15),
          fontFace: FONT_FACE,
          fontSize: 15,
        }
      );

      // Metric values
      row.values.forEach((val, colIdx) => {
        slide.addText(val, {
          x: colX[colIdx],
          y: y,
          w: 1.4,
          h: calcTextBoxHeight(15),
          fontFace: FONT_FACE,
          fontSize: 15,
          color: BLACK,
        });
      });
    });

    slide.addImage({
      path: LIGHT_GRAY_BLOCK,
      x: "65%",
      y: 0,
      w: "35%",
      h: "100%",
    });

    slide.addText("Takeaways", {
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.SECTION_TITLE,
      x: xRight,
      y: ySectionTitle,
      w: 4,
      h: hSectionTitle,
    });

    slide.addText(
      [
        {
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.",
          options: { bullet: true },
        },
        {
          text: "Nullam pharetra mauris tortor. In hac habitasse platea dictumst.",
          options: { bullet: true },
        },
        {
          text: "27",
          options: {
            superscript: true,
          },
        },
        {
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.",
          options: { bullet: true },
        },
        {
          text: "Nullam pharetra mauris tortor. In hac habitasse platea dictumst.",
          options: { bullet: true },
        },
        {
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.",
          options: { bullet: true },
        },
        {
          text: "28",
          options: {
            superscript: true,
          },
        },
      ],
      {
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.BULLET,
        x: xRight,
        y: yContent,
        w: "29%",
        h: calcTextBoxHeight(FONT_SIZE.BULLET, 14),
        paraSpaceAfter: FONT_SIZE.BULLET * 0.3,
      }
    );

    // 9) Footer
    slide.addText(
      [
        {
          text: "[27]",
          options: {
            hyperlink: { url: "【81†L1-L2】" }, // replace with actual citation target
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[28]",
          options: {
            hyperlink: { url: "【89†L3-L4】" }, // replace with actual citation target
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[29]",
          options: {
            hyperlink: { url: "【91†L5-L6】" }, // replace with actual citation target
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[30]",
          options: {
            hyperlink: { url: "【93†L7-L8】" }, // replace with actual citation target
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 24: Three column summary table
  {
    const slide = pptx.addSlide();

    // --- Title ---
    const hTitle = calcTextBoxHeight(FONT_SIZE.SLIDE_TITLE);
    slide.addText("Slide title", {
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.SLIDE_TITLE,
      color: NEAR_BLACK_NAVY,
      x: SLIDE_TITLE.X,
      y: SLIDE_TITLE.Y,
      w: SLIDE_TITLE.W,
      h: hTitle,
    });

    // --- Column headers ---
    const headerY = SLIDE_TITLE.Y + hTitle + 0.3;
    const numCols = 3;
    const colW = (SLIDE_WIDTH - 1.5) / numCols; // leave margin
    const colXs = [0.5, 0.5 + colW + 0.25, 0.5 + 2 * (colW + 0.25)];
    const headerProps = {
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.PLACEHOLDER,
      bold: false,
      color: NEAR_BLACK_NAVY,
      w: colW,
      h: calcTextBoxHeight(FONT_SIZE.PLACEHOLDER),
      align: "left",
      valign: "top",
    };
    colXs.forEach((x) => {
      slide.addText("Header", { x, y: headerY, ...headerProps });
    });

    // --- Data rows ---
    const data = [
      ["Subject", "78%", "+55"],
      ["Subject", "80%", "+22"],
      ["Subject", "24%", "-42"],
      ["Subject", "13%", "+32"],
      ["Subject", "67%", "-28"],
    ];
    const rowHeight = 0.7;
    const rowGap = 0.21;
    let y = headerY + calcTextBoxHeight(FONT_SIZE.PLACEHOLDER) + 0.1;
    data.forEach((row, idx) => {
      // draw underline first
      slide.addShape(pptx.ShapeType.line, {
        x: colXs[0],
        y: y,
        w: SLIDE_WIDTH - 1.0,
        h: 0.01,
        line: { color: "000000", width: 1 },
      });
      // place text boxes so their tops coincide with underline
      row.forEach((text, colIdx) => {
        slide.addText(text, {
          x: colXs[colIdx],
          y: y,
          w: colW,
          h: rowHeight,
          fontFace: FONT_FACE,
          fontSize: FONT_SIZE.BULLET,
          color: NEAR_BLACK_NAVY,
          align: "left",
          valign: "top",
        });
      });
      y += rowHeight + rowGap;
    });
    // --- Sources and slide number ---
    slide.addText(
      [
        {
          text: "[20]",
          options: {
            hyperlink: { url: "【12345†L200-L210】" }, // Placeholder, update.
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 25: Two column subject content table
  {
    const slide = pptx.addSlide();
    const hTitle = calcTextBoxHeight(FONT_SIZE.SLIDE_TITLE);
    slide.addText("Slide title", {
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.SLIDE_TITLE,
      color: NEAR_BLACK_NAVY,
      x: SLIDE_TITLE.X,
      y: SLIDE_TITLE.Y,
      w: SLIDE_TITLE.W,
      h: hTitle,
    });

    // --- Column headers ---
    const headerY = SLIDE_TITLE.Y + hTitle + 0.3;
    const leftColW = 2.1; // narrow left column
    const rightColW = SLIDE_WIDTH - leftColW - 1.0; // wide right column
    const leftColX = 1.0;
    const rightColX = leftColX + leftColW;
    const headerProps = {
      fontFace: FONT_FACE,
      fontSize: FONT_SIZE.PLACEHOLDER,
      bold: false,
      color: NEAR_BLACK_NAVY,
      w: leftColW,
      h: calcTextBoxHeight(FONT_SIZE.PLACEHOLDER),
      align: "left",
    };
    slide.addText("Header", { x: leftColX, y: headerY, ...headerProps });
    slide.addText("Header", {
      x: rightColX,
      y: headerY,
      ...headerProps,
      w: rightColW,
    });

    const MARGIN = 1.0;
    const headerLineY =
      headerY + calcTextBoxHeight(FONT_SIZE.PLACEHOLDER) + 0.05;
    const center_x = SLIDE_WIDTH / 2;
    const line_width = SLIDE_WIDTH - 2 * MARGIN;
    slide.addShape(pptx.ShapeType.line, {
      x: center_x - line_width / 2,
      y: headerLineY,
      w: line_width,
      h: 0.01,
      line: { color: "000000", width: 1 },
    });

    // --- Content rows ---
    const rowHeight = 1.2;
    const rowGap = 0.35;
    const rowYs = [
      headerLineY + 0.135,
      headerLineY + 0.135 + rowHeight + rowGap,
      headerLineY + 0.135 + (rowHeight + rowGap) * 2,
    ];
    rowYs.forEach((y, idx) => {
      slide.addText("Subject", {
        x: leftColX,
        y: y,
        w: leftColW,
        h: rowHeight,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.BULLET,
        color: NEAR_BLACK_NAVY,
        valign: "top",
        align: "left",
      });
      const content = [
        {
          text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
          options: { valign: "top", align: "left" },
        },
        {
          text: "Pellentesque ultricies quam ut massa luctus cursus.",
          options: { bullet: true },
        },
        {
          text: "20",
          options: { superscript: true },
        },
        { text: "\n", break: true },
        {
          text: "Nullam pharetra mauris tortor.",
          options: { bullet: true },
        },
        {
          text: "21",
          options: { superscript: true },
        },
      ].filter(Boolean);
      slide.addText(content, {
        x: rightColX,
        y: y,
        w: rightColW + 0.7,
        h: rowHeight,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.BULLET,
        color: NEAR_BLACK_NAVY,
        valign: "top",
        align: "left",
      });

      const center_x = SLIDE_WIDTH / 2;
      const line_width = SLIDE_WIDTH - 2 * MARGIN;
      slide.addShape(pptx.ShapeType.line, {
        x: center_x - line_width / 2,
        y: y + rowHeight,
        w: line_width,
        h: 0.01,
        line: { color: "000000", width: 1 },
      });
    });

    slide.addText(
      [
        {
          text: "[20]",
          options: {
            hyperlink: { url: "【12345†L1-L2】" }, // replace with actual citation target
            color: NEAR_BLACK_NAVY,
          },
        },
        {
          text: " ",
          options: { underline: false },
        },
        {
          text: "[21]",
          options: {
            hyperlink: { url: "【51†L3-L4】" }, // replace with actual citation target
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
      }
    );
  }

  // Slide 26: Horizontal timeline with categories and events
  {
    const slide = pptx.addSlide();
    addSlideTitle(slide, "Slide title");

    // timeline horizontal line
    const lineY = 0.22 * SLIDE_HEIGHT;
    const COLOR_LINE_GRAY = "888888";
    slide.addShape("line", {
      x: "10%",
      y: lineY,
      w: "83%",
      h: "0.1%",
      line: { color: COLOR_LINE_GRAY, width: 2 },
    });

    // Add timeline dots and labels
    const COLOR_TIMELINE_DOT = "666666";
    const colCount = 4;
    const colWidth = 0.22 * SLIDE_WIDTH;
    const dotDiameter = 0.1;
    const dotOffset = 2 * dotDiameter;
    const verticalLineHeight = 0.044 * SLIDE_HEIGHT;
    const dateY = 0.16 * SLIDE_HEIGHT;

    for (let i = 0; i < colCount; i++) {
      const x = 1.6 + i * colWidth;
      // Date label
      slide.addText("Date", {
        x: x,
        y: dateY,
        w: "20%",
        h: "5.3%",
        fontSize: FONT_SIZE.PLACEHOLDER,
        fontFace: FONT_FACE,
      });
      // Dot on the line
      slide.addShape("ellipse", {
        x: x + dotOffset,
        y: lineY - dotDiameter / 2,
        w: dotDiameter,
        h: dotDiameter,
        fill: { color: COLOR_TIMELINE_DOT },
        line: { color: COLOR_TIMELINE_DOT },
      });

      // Vertical line from dot to header
      slide.addShape("line", {
        x: x + dotOffset + dotDiameter / 2, // center of the dot
        y: lineY + dotDiameter / 2, // bottom of the dot
        w: 0,
        h: verticalLineHeight, // vertical height to header
        line: { color: COLOR_LINE_GRAY, width: 2 },
      });
    }

    const categories = ["Category 1", "Category 2"];
    const headers = ["Header 1", "Header 2", "Header 3", "Header 4"];
    const texts = [
      [
        [
          {
            text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut.",
          },
          { text: "35", options: { superscript: true } },
        ],
        [
          {
            text: "Nullam pharetra mauris tortor. In hac habitasse platea dictumst.",
          },
        ],
        [
          {
            text: "Pellentesque ultricies quam ut. Nullam pharetra mauris tortor.",
          },
          { text: "37", options: { superscript: true } },
          { text: "In hac habitasse platea." },
        ],
        [
          {
            text: "Lorem ipsum dolor sit amet. Pellentesque ultricies quam ut.",
          },
        ],
      ],
      [
        [
          {
            text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut. Ultricies quam ut.",
          },
        ],
        [
          {
            text: "Lorem ipsum dolor sit amet, elit. Nullam pharetra mauris tortor.",
          },
          { text: "36", options: { superscript: true } },
        ],
        [
          {
            text: "Pellentesque ultricies quam ut. In hac habitasse platea.\n\nLorem ipsum dolor sit amet, consectetur elit.",
          },
        ],
        [
          {
            text: "Lorem ipsum dolor sit amet, consectetur. Pellen tesque ultricies quam ut.",
          },
          { text: "38", options: { superscript: true } },
        ],
      ],
    ];

    const headerOffset = 0.293 * SLIDE_HEIGHT;
    const rowHeight = 0.284 * SLIDE_HEIGHT;
    const categoryHeightOffset = 0.06 * SLIDE_HEIGHT;
    categories.forEach((cat, rowIndex) => {
      slide.addText(cat, {
        x: 0.133333,
        y: headerOffset + categoryHeightOffset + rowIndex * rowHeight,
        w: 1.4,
        h: calcTextBoxHeight(FONT_SIZE.BULLET),
        fontSize: FONT_SIZE.BULLET,
        fontFace: FONT_FACE,
        valign: "top",
        align: "left",
      });

      for (let colIndex = 0; colIndex < colCount; colIndex++) {
        const x = 1.6 + colIndex * colWidth;
        const y = headerOffset + rowIndex * rowHeight;

        if (rowIndex === 0) {
          slide.addText(headers[colIndex], {
            x: x,
            y: y,
            w: colWidth - 0.2666667,
            h: 0.4,
            fontSize: FONT_SIZE.PLACEHOLDER,
            fontFace: FONT_FACE,
          });
        }

        slide.addText(texts[rowIndex][colIndex], {
          x: x,
          y: y + 0.4,
          w: colWidth - 0.2666667,
          h: 1.7333333,
          fontSize: FONT_SIZE.PLACEHOLDER,
          fontFace: FONT_FACE,
          lineSpacingMultiple: 1.0,
          valign: "top",
          align: "left",
        });
      }
    });

    slide.addText(
      [
        {
          text: "[35]",
          options: {
            hyperlink: { url: "【12345†L350-L360】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[36]",
          options: {
            hyperlink: { url: "【67890†L361-L370】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[37]",
          options: {
            hyperlink: { url: "【24680†L371-L380】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
        { text: " ", options: { underline: false } },
        {
          text: "[38]",
          options: {
            hyperlink: { url: "【13579†L381-L390】" }, // randomized number
            color: NEAR_BLACK_NAVY,
          },
        },
      ],
      {
        x: SLIDE_TITLE.X,
        y: MARGINS.DEFAULT_CITATION,
        w: SLIDE_TITLE.W,
        h: CITATION_HEIGHT,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.CITATION,
        color: NEAR_BLACK_NAVY,
      }
    );
  }

  // Slide 27: Agenda
  {
    const slide = pptx.addSlide();
    slide.background = { fill: NEAR_BLACK_NAVY };

    addSlideTitle(slide, "Agenda", WHITE);

    // Agenda items
    const agendaItems = [
      "Agenda item one",
      "Agenda item two",
      "Agenda item three",
      "Agenda item four",
      "Agenda item five",
      "Agenda item six",
      "Agenda item seven",
    ];

    const lineH = calcTextBoxHeight(FONT_SIZE.AGENDA);
    const startY = SLIDE_TITLE.Y + hSlideTitle + 0.3;

    agendaItems.forEach((text, idx) => {
      const y = startY + idx * lineH;

      slide.addText(text, {
        x: "50%",
        y: y,
        w: "40%",
        h: lineH,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.AGENDA,
        color: "FFFFFF",
        align: "left",
      });

      slide.addText(String(idx + 1).padStart(2, "0"), {
        x: "88%",
        y,
        w: "7%",
        h: lineH,
        fontFace: FONT_FACE,
        fontSize: FONT_SIZE.AGENDA,
        color: "FFFFFF",
        align: "right",
      });
    });
  }
  await pptx.writeFile({ fileName: "slides_template.pptx" });
})();
