/* ═══════════════════════════════════════════════════════════════
   SMART SLIDES — Complete Fixed Version

   GESTURES:
   👍 Thumbs Up   → Next slide
   ✌️ Victory     → Previous slide
   👋 Open Palm   → Enter fullscreen (hold 0.4 sec)
   ✊ Closed Fist → Exit fullscreen

   ALSO:
   • Two hands → Pinch zoom
   • Voice commands
   • Text-to-speech
   • Hugging Face AI Assistant
   • PPTX-only upload
   • Touch navigation
   ═══════════════════════════════════════════════════════════════ */

"use strict";

/* ─────────────────────────────────────────────
   1. APP STATE
   ───────────────────────────────────────────── */

const State = {
  slides: [],
  current: 0,
  total: 0,
  loaded: false,

  ttsActive: false,
  ttsUtterance: null,

  voiceActive: false,
  gestureActive: false,
  cameraStream: null,

  hfToken: "",
  aiMode: null,

  handsModel: null,
  lastGestureTime: 0,
  gestureCooldown: 900,

  presentationMode: false,
  zoomLevel: 1.0,

  wristHistory: [[], []],
  HISTORY_LEN: 12,
  SWIPE_MIN_DIST: 0.18,
  SWIPE_MAX_TIME: 700,
  SWIPE_LINEARITY: 0.75
};

const pdfTextCache = {};


/* ─────────────────────────────────────────────
   2. DOM REFERENCES
   ───────────────────────────────────────────── */

const $ = id => document.getElementById(id);

const DOM = {
  fileInput: $("fileInput"),
  dropZone: $("dropZone"),
  uploadProgress: $("uploadProgress"),
  progressFill: $("progressFill"),
  progressLabel: $("progressLabel"),

  viewerEmpty: $("viewerEmpty"),
  slideCanvas: $("slideCanvas"),
  slideImg: $("slideImg"),
  swipeOverlay: $("swipeOverlay"),
  gestureAnim: $("gestureAnim"),
  viewerContainer: $("viewerContainer"),

  currentSlideNum: $("currentSlideNum"),
  totalSlideNum: $("totalSlideNum"),

  btnFirst: $("btnFirst"),
  btnPrev: $("btnPrev"),
  btnNext: $("btnNext"),
  btnLast: $("btnLast"),
  btnFullscreen: $("btnFullscreen"),
  btnStartPresent: $("btnStartPresent"),

  webcam: $("webcam"),
  gestureCanvas: $("gestureCanvas"),
  cameraOff: $("cameraOff"),
  btnStartCamera: $("btnStartCamera"),
  btnStopCamera: $("btnStopCamera"),
  cameraError: $("cameraError"),

  btnStartVoice: $("btnStartVoice"),
  btnStopVoice: $("btnStopVoice"),
  voiceError: $("voiceError"),
  voiceLog: $("voiceLog"),

  voiceDot: $("voiceDot"),
  voiceLabel: $("voiceLabel"),
  gestureDot: $("gestureDot"),
  gestureLabel: $("gestureLabel"),
  ttsDot: $("ttsDot"),
  ttsLabel: $("ttsLabel"),
  lastCommandDisplay: $("lastCommandDisplay"),

  btnReadSlide: $("btnReadSlide"),
  btnPauseRead: $("btnPauseRead"),
  btnResumeRead: $("btnResumeRead"),
  btnStopRead: $("btnStopRead"),
  btnRepeatRead: $("btnRepeatRead"),

  hfToken: $("hfToken"),
  btnSaveToken: $("btnSaveToken"),
  btnSummarize: $("btnSummarize"),
  btnExplain: $("btnExplain"),
  btnQuestion: $("btnQuestion"),
  aiQuestionBox: $("aiQuestionBox"),
  aiQuestionInput: $("aiQuestionInput"),
  btnAskQuestion: $("btnAskQuestion"),
  aiResult: $("aiResult"),
  aiResultText: $("aiResultText"),
  aiLoading: $("aiLoading"),
  btnCloseAI: $("btnCloseAI"),

  thumbnailStrip: $("thumbnailStrip"),
  helpModal: $("helpModal"),
  btnHelp: $("btnHelp"),
  btnCloseHelp: $("btnCloseHelp"),

  presentationOverlay: $("presentationOverlay"),
  fsCanvas: $("fsCanvas"),
  fsImg: $("fsImg"),
  fsBtnPrev: $("fsBtnPrev"),
  fsBtnNext: $("fsBtnNext"),
  fsBtnExit: $("fsBtnExit"),
  fsCounter: $("fsCounter"),
  fsGestureAnim: $("fsGestureAnim")
};


/* ═══════════════════════════════════════════════
   3. PPTX LOADING
   ═══════════════════════════════════════════════ */

function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) {
      resolve();
      return;
    }

    const s = document.createElement("script");

    s.src = src;

    s.onload = resolve;

    s.onerror = () =>
      reject(new Error("Script load failed: " + src));

    document.head.appendChild(s);
  });
}


function isPPTX(file) {
  return (
    file.name.toLowerCase().endsWith(".pptx") ||
    file.type ===
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  );
}


async function loadPPTX(file) {

  if (!window.JSZip) {

    showProgress(
      2,
      "Loading JSZip library…"
    );

    await loadScript(
      "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"
    );
  }


  showProgress(
    5,
    "Unzipping PowerPoint…"
  );


  const ab = await file.arrayBuffer();

  let zip;


  try {

    zip = await JSZip.loadAsync(ab);

  } catch (e) {

    throw new Error(
      "Could not read the .pptx file. Make sure it is a valid PowerPoint file."
    );

  }


  const slideKeys = Object.keys(zip.files)

    .filter(k =>
      /^ppt\/slides\/slide\d+\.xml$/i.test(k)
    )

    .sort((a, b) => {

      const na = parseInt(
        a.match(/\d+/)[0]
      );

      const nb = parseInt(
        b.match(/\d+/)[0]
      );

      return na - nb;

    });


  if (!slideKeys.length) {

    throw new Error(
      "No slides found in this .pptx file."
    );

  }


  let slideW = 9144000;
  let slideH = 5143500;


  if (zip.files["ppt/presentation.xml"]) {

    const presXml =
      await zip.files["ppt/presentation.xml"]
        .async("string");


    const szMatch =
      presXml.match(
        /<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/
      );


    if (szMatch) {

      slideW =
        parseInt(szMatch[1]);

      slideH =
        parseInt(szMatch[2]);

    }
  }


  const slides = [];

  const parser =
    new DOMParser();


  for (
    let i = 0;
    i < slideKeys.length;
    i++
  ) {

    showProgress(
      10 +
        (i / slideKeys.length) * 85,
      `Rendering slide ${i + 1} / ${slideKeys.length}…`
    );


    const xmlStr =
      await zip.files[slideKeys[i]]
        .async("string");


    const relPath =
      slideKeys[i]
        .replace(
          "slides/slide",
          "slides/_rels/slide"
        )
        .replace(
          ".xml",
          ".xml.rels"
        );


    const relMap = {};


    if (zip.files[relPath]) {

      const relXml =
        await zip.files[relPath]
          .async("string");


      const relDoc =
        parser.parseFromString(
          relXml,
          "text/xml"
        );


      relDoc
        .querySelectorAll("Relationship")
        .forEach(r => {

          relMap[
            r.getAttribute("Id")
          ] =
            r.getAttribute("Target");

        });
    }


    let bgColor = "#FFFFFF";


    const bgMatch =
      xmlStr.match(
        /<p:bg>.*?<a:srgbClr val="([0-9A-Fa-f]{6})".*?<\/p:bg>/s
      );


    if (bgMatch) {

      bgColor =
        "#" + bgMatch[1];

    }


    const canvas =
      await renderPPTXSlide(
        xmlStr,
        zip,
        relMap,
        slideW,
        slideH,
        bgColor,
        i,
        parser
      );


    slides.push({
      type: "canvas",
      data: canvas
    });


    const text =
      extractPPTXText(xmlStr);


    if (text) {

      pdfTextCache[i] = text;

    }
  }


  return slides;
}


function extractPPTXText(xmlStr) {

  const matches =
    [
      ...xmlStr.matchAll(
        /<a:t[^>]*>([^<]*)<\/a:t>/g
      )
    ];


  return matches

    .map(m => m[1].trim())

    .filter(Boolean)

    .join(" ");
}


const EMU_TO_PX = 1 / 9525;


async function renderPPTXSlide(
  xmlStr,
  zip,
  relMap,
  slideW,
  slideH,
  bgColor,
  slideIndex,
  parser
) {

  const CANVAS_W = 1280;

  const CANVAS_H =
    Math.round(
      CANVAS_W *
        (slideH / slideW)
    );


  const scaleX =
    CANVAS_W / slideW;

  const scaleY =
    CANVAS_H / slideH;


  const canvas =
    document.createElement("canvas");


  canvas.width =
    CANVAS_W;

  canvas.height =
    CANVAS_H;


  const ctx =
    canvas.getContext("2d");


  ctx.fillStyle =
    bgColor;

  ctx.fillRect(
    0,
    0,
    CANVAS_W,
    CANVAS_H
  );


  const doc =
    parser.parseFromString(
      xmlStr,
      "text/xml"
    );


  const spTree =
    doc.querySelector("spTree") ||
    doc.documentElement;


  const children =
    spTree.children;


  for (const el of children) {

    const tag =
      el.tagName;


    if (tag === "p:pic") {

      await drawPicture(
        ctx,
        el,
        zip,
        relMap,
        scaleX,
        scaleY,
        parser
      );

    }

    else if (tag === "p:sp") {

      drawShape(
        ctx,
        el,
        scaleX,
        scaleY,
        CANVAS_W,
        CANVAS_H
      );

    }

    else if (tag === "p:grpSp") {

      for (
        const child of el.children
      ) {

        if (
          child.tagName === "p:sp"
        ) {

          drawShape(
            ctx,
            child,
            scaleX,
            scaleY,
            CANVAS_W,
            CANVAS_H
          );

        }

        if (
          child.tagName === "p:pic"
        ) {

          await drawPicture(
            ctx,
            child,
            zip,
            relMap,
            scaleX,
            scaleY,
            parser
          );

        }
      }
    }
  }


  return canvas;
}


function getXfrm(el) {

  const xfrm =
    el.querySelector("xfrm") ||
    el.querySelector("spPr xfrm");


  if (!xfrm) return null;


  const off =
    xfrm.querySelector("off");

  const ext =
    xfrm.querySelector("ext");


  if (!off || !ext)
    return null;


  return {

    x:
      parseInt(
        off.getAttribute("x") || 0
      ),

    y:
      parseInt(
        off.getAttribute("y") || 0
      ),

    cx:
      parseInt(
        ext.getAttribute("cx") ||
          1000000
      ),

    cy:
      parseInt(
        ext.getAttribute("cy") ||
          500000
      )
  };
}


function drawShape(
  ctx,
  sp,
  scaleX,
  scaleY,
  canvasW,
  canvasH
) {

  const txBody =
    sp.querySelector("txBody");


  if (!txBody) return;


  const xfrm =
    getXfrm(sp);


  let rx = 40;
  let ry = 40;
  let rw = canvasW - 80;
  let rh = canvasH - 80;


  if (xfrm) {

    rx =
      Math.round(
        xfrm.x * scaleX
      );

    ry =
      Math.round(
        xfrm.y * scaleY
      );

    rw =
      Math.round(
        xfrm.cx * scaleX
      );

    rh =
      Math.round(
        xfrm.cy * scaleY
      );


    rx =
      Math.max(
        0,
        Math.min(
          rx,
          canvasW - 4
        )
      );


    ry =
      Math.max(
        0,
        Math.min(
          ry,
          canvasH - 4
        )
      );


    rw =
      Math.min(
        rw,
        canvasW - rx
      );


    rh =
      Math.min(
        rh,
        canvasH - ry
      );
  }


  const solidFill =
    sp.querySelector(
      "spPr solidFill srgbClr"
    );


  if (solidFill) {

    ctx.fillStyle =
      "#" +
      solidFill.getAttribute("val");


    ctx.fillRect(
      rx,
      ry,
      rw,
      rh
    );
  }


  const bodyPr =
    txBody.querySelector("bodyPr");


  const anchor =
    bodyPr
      ? bodyPr.getAttribute("anchor")
      : "ctr";


  const paras =
    [
      ...txBody.querySelectorAll(
        "a\\:p, p"
      )
    ]
    .filter(
      p => p.tagName === "a:p"
    );


  const lines = [];


  for (
    const para of paras
  ) {

    const pPr =
      para.querySelector(
        "a\\:pPr, pPr"
      );


    const defRPr =
      para.querySelector(
        "a\\:defRPr, defRPr"
      );


    const runs =
      [...para.children]
        .filter(
          c => c.tagName === "a:r"
        );


    const brNodes =
      [...para.children]
        .filter(
          c => c.tagName === "a:br"
        );


    let defSz = 18;


    if (defRPr) {

      const sz =
        parseInt(
          defRPr.getAttribute("sz") || 0
        );


      if (sz) {

        defSz =
          Math.max(
            8,
            Math.min(
              96,
              sz / 100
            )
          );

      }
    }


    const align =
      pPr
        ? (
            pPr.getAttribute("algn") ||
            "l"
          )
        : "l";


    const ctxAlign =
      align === "ctr"
        ? "center"
        : align === "r"
          ? "right"
          : "left";


    if (
      !runs.length &&
      !brNodes.length
    ) {

      lines.push({
        text: "",
        fontSize: defSz * 0.9,
        bold: false,
        color: "#000000",
        align: ctxAlign,
        empty: true
      });

      continue;
    }


    for (
      const run of runs
    ) {

      const rPr =
        run.querySelector(
          "a\\:rPr, rPr"
        );


      let fontSize =
        defSz;

      let bold = false;

      let color =
        "#1a1a1a";

      let italic = false;


      if (rPr) {

        const sz =
          parseInt(
            rPr.getAttribute("sz") || 0
          );


        if (sz) {

          fontSize =
            Math.max(
              8,
              Math.min(
                96,
                sz / 100
              )
            );

        }


        bold =
          rPr.getAttribute("b") === "1";


        italic =
          rPr.getAttribute("i") === "1";


        const srgb =
          rPr.querySelector(
            "a\\:solidFill a\\:srgbClr, solidFill srgbClr"
          );


        if (srgb) {

          color =
            "#" +
            srgb.getAttribute("val");

        }
      }


      const t =
        run.querySelector(
          "a\\:t, t"
        );


      const text =
        t
          ? t.textContent
          : "";


      if (text) {

        lines.push({
          text,
          fontSize,
          bold,
          color,
          italic,
          align: ctxAlign
        });

      }
    }
  }


  let totalH = 0;


  for (
    const ln of lines
  ) {

    totalH +=
      ln.empty
        ? ln.fontSize * 0.6
        : ln.fontSize * 1.4;

  }


  let curY;


  if (
    anchor === "ctr"
  ) {

    curY =
      ry +
      Math.max(
        0,
        (rh - totalH) / 2
      );

  }

  else if (
    anchor === "b"
  ) {

    curY =
      ry +
      Math.max(
        0,
        rh -
          totalH -
          8
      );

  }

  else {

    curY =
      ry + 8;

  }


  ctx.save();

  ctx.rect(
    rx,
    ry,
    rw,
    rh
  );

  ctx.clip();


  for (
    const ln of lines
  ) {

    if (ln.empty) {

      curY +=
        ln.fontSize * 0.6;

      continue;

    }


    if (
      curY >
      ry + rh + ln.fontSize
    ) {

      break;

    }


    const fontStr =
      `${ln.italic ? "italic " : ""}` +
      `${ln.bold ? "bold " : ""}` +
      `${ln.fontSize}px Arial, sans-serif`;


    ctx.font =
      fontStr;

    ctx.fillStyle =
      ln.color;

    ctx.textBaseline =
      "top";

    ctx.textAlign =
      ln.align;


    const anchorX =
      ln.align === "center"
        ? rx + rw / 2
        : ln.align === "right"
          ? rx + rw - 8
          : rx + 10;


    const words =
      ln.text.split(" ");


    let line = "";


    for (
      const word of words
    ) {

      const test =
        line
          ? line + " " + word
          : word;


      if (
        ctx.measureText(test).width >
          rw - 16 &&
        line
      ) {

        ctx.fillText(
          line,
          anchorX,
          curY
        );


        curY +=
          ln.fontSize * 1.35;


        line = word;

      }

      else {

        line = test;

      }
    }


    if (line) {

      ctx.fillText(
        line,
        anchorX,
        curY
      );


      curY +=
        ln.fontSize * 1.35;

    }
  }


  ctx.restore();
}


async function drawPicture(
  ctx,
  pic,
  zip,
  relMap,
  scaleX,
  scaleY,
  parser
) {

  const blip =
    pic.querySelector(
      "a\\:blip, blip"
    );


  if (!blip) return;


  const rId =
    blip.getAttribute("r:embed") ||
    blip.getAttribute("embed");


  if (
    !rId ||
    !relMap[rId]
  ) {

    return;

  }


  const rawTarget =
    relMap[rId];


  const target =
    rawTarget.startsWith("../")
      ? "ppt/" +
        rawTarget.slice(3)

      : rawTarget.startsWith("ppt/")
        ? rawTarget

        : "ppt/slides/" +
          rawTarget;


  if (
    !zip.files[target]
  ) {

    return;

  }


  const xfrm =
    getXfrm(pic);


  let px = 0;
  let py = 0;
  let pw = ctx.canvas.width;
  let ph = ctx.canvas.height;


  if (xfrm) {

    px =
      Math.round(
        xfrm.x * scaleX
      );

    py =
      Math.round(
        xfrm.y * scaleY
      );

    pw =
      Math.round(
        xfrm.cx * scaleX
      );

    ph =
      Math.round(
        xfrm.cy * scaleY
      );
  }


  const ext =
    target
      .split(".")
      .pop()
      .toLowerCase();


  const mime =
    ext === "png"
      ? "image/png"

      : ext === "jpg" ||
        ext === "jpeg"

        ? "image/jpeg"

        : ext === "gif"

          ? "image/gif"

          : ext === "webp"

            ? "image/webp"

            : null;


  if (!mime) return;


  const b64 =
    await zip.files[target]
      .async("base64");


  await new Promise(
    resolve => {

      const img =
        new Image();


      img.onload =
        () => {

          ctx.drawImage(
            img,
            px,
            py,
            pw,
            ph
          );

          resolve();

        };


      img.onerror =
        resolve;


      img.src =
        `data:${mime};base64,${b64}`;

    }
  );
}


/* ═══════════════════════════════════════════════
   4. FILE UPLOAD
   ═══════════════════════════════════════════════ */

async function handleFiles(files) {

  if (
    !files ||
    !files.length
  ) return;


  const fileArr =
    Array.from(files);


  const pptxFile =
    fileArr.find(isPPTX);


  if (!pptxFile) {

    showUploadError(
      "Only PowerPoint (.pptx) files are accepted. Please upload a .pptx file."
    );

    return;

  }


  DOM.uploadProgress
    .classList
    .remove("hidden");


  DOM.viewerEmpty
    .classList
    .remove("hidden");


  DOM.slideCanvas
    .classList
    .add("hidden");


  DOM.slideImg
    .classList
    .add("hidden");


  try {

    const slides =
      await loadPPTX(
        pptxFile
      );


    State.slides =
      slides;


    State.total =
      slides.length;


    State.loaded =
      true;


    State.current =
      0;


    State.zoomLevel =
      1.0;


    buildThumbnails();

    renderSlide(0);


    showProgress(
      100,
      `✓ Loaded ${slides.length} slide${
        slides.length !== 1
          ? "s"
          : ""
      }`
    );


    setTimeout(
      () =>
        DOM.uploadProgress
          .classList
          .add("hidden"),
      2500
    );

  }

  catch (err) {

    showUploadError(
      err.message
    );

    console.error(err);

  }
}


function showUploadError(msg) {

  DOM.uploadProgress
    .classList
    .add("hidden");


  const errEl =
    document.createElement(
      "div"
    );


  errEl.className =
    "upload-error-msg";


  errEl.textContent =
    "⚠ " + msg;


  DOM.dropZone
    .appendChild(
      errEl
    );


  setTimeout(
    () => errEl.remove(),
    6000
  );
}


/* ═══════════════════════════════════════════════
   5. SLIDE RENDERER
   ═══════════════════════════════════════════════ */

function renderSlide(index) {

  if (
    !State.loaded ||
    !State.slides.length
  ) return;


  index =
    Math.max(
      0,
      Math.min(
        index,
        State.total - 1
      )
    );


  const slide =
    State.slides[index];


  State.current =
    index;


  DOM.viewerEmpty
    .classList
    .add("hidden");


  DOM.slideCanvas
    .classList
    .add("hidden");


  DOM.slideImg
    .classList
    .add("hidden");


  if (
    slide.type === "canvas"
  ) {

    DOM.slideCanvas
      .classList
      .remove("hidden");


    DOM.slideCanvas.width =
      slide.data.width;


    DOM.slideCanvas.height =
      slide.data.height;


    const ctx =
      DOM.slideCanvas
        .getContext("2d");


    ctx.setTransform(
      State.zoomLevel,
      0,
      0,
      State.zoomLevel,
      0,
      0
    );


    ctx.drawImage(
      slide.data,
      0,
      0
    );


    ctx.setTransform(
      1,
      0,
      0,
      1,
      0,
      0
    );

  }

  else {

    DOM.slideImg
      .classList
      .remove("hidden");


    DOM.slideImg.src =
      slide.data;


    DOM.slideImg.style.transform =
      `scale(${State.zoomLevel})`;

  }


  DOM.currentSlideNum.textContent =
    index + 1;


  DOM.totalSlideNum.textContent =
    State.total;


  if (
    State.presentationMode
  ) {

    renderFullscreen(index);

  }


  updateThumbnailActive(
    index
  );
}


function renderFullscreen(index) {

  if (
    !State.slides.length
  ) return;


  const slide =
    State.slides[index];


  DOM.fsCanvas
    .classList
    .add("hidden");


  DOM.fsImg
    .classList
    .add("hidden");


  if (
    slide.type === "canvas"
  ) {

    DOM.fsCanvas
      .classList
      .remove("hidden");


    DOM.fsCanvas.width =
      slide.data.width;


    DOM.fsCanvas.height =
      slide.data.height;


    DOM.fsCanvas
      .getContext("2d")
      .drawImage(
        slide.data,
        0,
        0
      );

  }

  else {

    DOM.fsImg
      .classList
      .remove("hidden");


    DOM.fsImg.src =
      slide.data;

  }


  DOM.fsCounter.textContent =
    `${index + 1} / ${State.total}`;
}


function goToSlide(n) {

  if (!State.loaded)
    return;


  n =
    Math.max(
      0,
      Math.min(
        n,
        State.total - 1
      )
    );


  if (
    n === State.current
  ) return;


  renderSlide(n);
}


const nextSlide =
  () =>
    goToSlide(
      State.current + 1
    );


const prevSlide =
  () =>
    goToSlide(
      State.current - 1
    );


const firstSlide =
  () =>
    goToSlide(0);


const lastSlide =
  () =>
    goToSlide(
      State.total - 1
    );


function zoomIn() {

  State.zoomLevel =
    Math.min(
      3.0,
      State.zoomLevel + 0.25
    );


  renderSlide(
    State.current
  );


  flashGestureLabel(
    "🔍 Zoom In"
  );
}


function zoomOut() {

  State.zoomLevel =
    Math.max(
      0.5,
      State.zoomLevel - 0.25
    );


  renderSlide(
    State.current
  );


  flashGestureLabel(
    "🔍 Zoom Out"
  );
}


function zoomReset() {

  State.zoomLevel =
    1.0;


  renderSlide(
    State.current
  );
}


/* ═══════════════════════════════════════════════
   6. THUMBNAILS
   ═══════════════════════════════════════════════ */

function buildThumbnails() {

  DOM.thumbnailStrip.innerHTML =
    "";


  State.slides.forEach(
    (slide, i) => {

      const item =
        document.createElement(
          "div"
        );


      item.className =
        "thumb-item" +
        (
          i === 0
            ? " active"
            : ""
        );


      item.dataset.index =
        i;


      if (
        slide.type === "canvas"
      ) {

        const c =
          document.createElement(
            "canvas"
          );


        c.width = 160;
        c.height = 90;


        c.getContext("2d")
          .drawImage(
            slide.data,
            0,
            0,
            160,
            90
          );


        item.appendChild(c);

      }

      else {

        const img =
          document.createElement(
            "img"
          );


        img.src =
          slide.data;


        img.alt =
          `Slide ${i + 1}`;


        item.appendChild(img);

      }


      const num =
        document.createElement(
          "span"
        );


      num.className =
        "thumb-num";


      num.textContent =
        i + 1;


      item.appendChild(num);


      item.addEventListener(
        "click",
        () =>
          goToSlide(i)
      );


      DOM.thumbnailStrip
        .appendChild(item);

    }
  );
}


function updateThumbnailActive(index) {

  document
    .querySelectorAll(
      ".thumb-item"
    )
    .forEach(
      (el, i) => {

        el.classList.toggle(
          "active",
          i === index
        );


        if (
          i === index
        ) {

          el.scrollIntoView({
            block: "nearest"
          });

        }
      }
    );
}


/* ═══════════════════════════════════════════════
   7. GESTURE ENGINE
   ═══════════════════════════════════════════════ */

const LM = {

  WRIST: 0,

  THUMB_TIP: 4,

  INDEX_MCP: 5,
  INDEX_TIP: 8,

  MIDDLE_MCP: 9,
  MIDDLE_TIP: 12,

  RING_MCP: 13,
  RING_TIP: 16,

  PINKY_MCP: 17,
  PINKY_TIP: 20
};


let handsModel = null;

let gestureRAF = null;

let prevPinchDist = null;

let pinchZoomCooldown = 0;

const PINCH_THRESHOLD =
  0.04;


/*
   One-shot gesture latch.

   This is VERY IMPORTANT.

   Without this latch:
   holding 👍 could advance:
   slide 1 → 2 → 3 → 4 → 5 ...

   With the latch:
   👍 → one next slide
   release/change gesture
   👍 → another next slide
*/


let discreteGestureLatched =
  false;


/*
   Open palm fullscreen state
*/


let palmHoldStart = 0;

const PALM_HOLD_MS =
  400;


let palmTriggeredThisHold =
  false;


/* ─────────────────────────────────────────────
   LANDMARK HELPERS
   ───────────────────────────────────────────── */

function landmarkDist(a, b) {

  const dx =
    a.x - b.x;

  const dy =
    a.y - b.y;

  return Math.sqrt(
    dx * dx +
    dy * dy
  );
}


function isFingerExtended(
  lm,
  mcpIdx,
  tipIdx
) {

  const wrist =
    lm[LM.WRIST];


  return (
    landmarkDist(
      lm[tipIdx],
      wrist
    ) >
    landmarkDist(
      lm[mcpIdx],
      wrist
    ) * 1.1
  );
}


/*
   Open palm.

   We deliberately check the four fingers here.
   GestureRecognizer itself also provides the
   high-level "Open_Palm" category.
*/


function isOpenPalm(lm) {

  return (
    isFingerExtended(
      lm,
      LM.INDEX_MCP,
      LM.INDEX_TIP
    ) &&

    isFingerExtended(
      lm,
      LM.MIDDLE_MCP,
      LM.MIDDLE_TIP
    ) &&

    isFingerExtended(
      lm,
      LM.RING_MCP,
      LM.RING_TIP
    ) &&

    isFingerExtended(
      lm,
      LM.PINKY_MCP,
      LM.PINKY_TIP
    )
  );
}


/* ─────────────────────────────────────────────
   MEDIAPIPE GESTURE RECOGNITION
   ───────────────────────────────────────────── */


/*
   MediaPipe GestureRecognizer returns:

   Thumb_Up
   Victory
   Open_Palm
   Closed_Fist
   Pointing_Up
   Thumb_Down
   ILoveYou
   None

   We only use the gestures requested by you.
*/


function getRecognizedGesture(
  results,
  handIndex = 0
) {

  if (
    !results ||
    !results.gestures ||
    !results.gestures[handIndex]
  ) {

    return null;

  }


  const list =
    results.gestures[
      handIndex
    ];


  if (
    !list ||
    !list.length
  ) {

    return null;

  }


  const top =
    list[0];


  const name =
    top.categoryName ||
    top.displayName ||
    "";


  const score =
    Number(
      top.score || 0
    );


  /*
     0.65 is a reasonable confidence threshold.

     If detection is too difficult for your camera,
     you can lower this to 0.55.
  */


  if (
    score < 0.65
  ) {

    return null;

  }


  return name;
}


/* ─────────────────────────────────────────────
   DISCRETE GESTURES
   ───────────────────────────────────────────── */


function handleDiscreteGesture(
  gesture,
  now
) {

  /*
     No gesture:
     release the latch.
  */


  if (!gesture) {

    discreteGestureLatched =
      false;

    return false;

  }


  /*
     If the same gesture is still being held,
     do NOT execute it again.
  */


  if (
    discreteGestureLatched
  ) {

    return true;

  }


  /*
     👍 THUMBS UP
     NEXT SLIDE
  */


  if (
    gesture === "Thumb_Up" &&
    now -
      State.lastGestureTime >
      State.gestureCooldown
  ) {

    discreteGestureLatched =
      true;


    State.lastGestureTime =
      now;


    nextSlide();


    flashGestureLabel(
      "👍 Next"
    );


    return true;
  }


  /*
     ✌️ VICTORY
     PREVIOUS SLIDE
  */


  if (
    gesture === "Victory" &&
    now -
      State.lastGestureTime >
      State.gestureCooldown
  ) {

    discreteGestureLatched =
      true;


    State.lastGestureTime =
      now;


    prevSlide();


    flashGestureLabel(
      "✌️ Previous"
    );


    return true;
  }


  /*
     ✊ CLOSED FIST
     EXIT FULLSCREEN

     Only does anything while
     presentationMode is active.
  */


  if (
    gesture === "Closed_Fist" &&
    State.presentationMode &&
    now -
      State.lastGestureTime >
      State.gestureCooldown
  ) {

    discreteGestureLatched =
      true;


    State.lastGestureTime =
      now;


    closePresentation();


    flashGestureLabel(
      "✊ Exit Fullscreen"
    );


    return true;
  }


  return false;
}


/* ─────────────────────────────────────────────
   SWIPE DETECTION
   ───────────────────────────────────────────── */


function analyseSwipe(history) {

  if (
    history.length < 4
  ) {

    return null;

  }


  const first =
    history[0];


  const last =
    history[
      history.length - 1
    ];


  const dt =
    last.t - first.t;


  const dx =
    last.x - first.x;


  const dy =
    last.y - first.y;


  const dist =
    Math.abs(dx);


  const totalDist =
    Math.sqrt(
      dx * dx +
      dy * dy
    );


  if (
    dt >
    State.SWIPE_MAX_TIME
  ) {

    return null;

  }


  if (
    dist <
    State.SWIPE_MIN_DIST
  ) {

    return null;

  }


  const linearity =
    totalDist > 0
      ? dist / totalDist
      : 0;


  if (
    linearity <
    State.SWIPE_LINEARITY
  ) {

    return null;

  }


  return dx > 0
    ? "right"
    : "left";
}


/* ─────────────────────────────────────────────
   INITIALIZE MEDIAPIPE
   ───────────────────────────────────────────── */

async function initHandsModel() {

  if (
    !window._MPImports
  ) {

    showGestureError(
      "MediaPipe library not loaded. Refresh and try again."
    );


    return false;

  }


  const {
    GestureRecognizer,
    FilesetResolver
  } =
    window._MPImports;


  try {

    const vision =
      await FilesetResolver
        .forVisionTasks(
          "https://cdn.jsdelivr.net/npm/@mediapipe/tasks-vision@0.10.3/wasm"
        );


    handsModel =
      await GestureRecognizer
        .createFromOptions(
          vision,
          {

            baseOptions: {

              modelAssetPath:
                "https://storage.googleapis.com/mediapipe-models/gesture_recognizer/gesture_recognizer/float16/1/gesture_recognizer.task",

              delegate:
                "GPU"

            },


            runningMode:
              "VIDEO",


            numHands:
              2

          }
        );


    return true;

  }

  catch (err) {

    console.warn(
      "Hands model load error:",
      err
    );


    showGestureError(
      "Gesture model failed to load. Check your internet connection."
    );


    return false;

  }
}


/* ─────────────────────────────────────────────
   MAIN GESTURE LOOP
   ───────────────────────────────────────────── */

function processGestureFrame(
  timestamp
) {

  if (
    !State.gestureActive ||
    !handsModel
  ) {

    return;

  }


  gestureRAF =
    requestAnimationFrame(
      processGestureFrame
    );


  const video =
    DOM.webcam;


  const canvas =
    DOM.gestureCanvas;


  const ctx =
    canvas.getContext(
      "2d"
    );


  if (
    video.readyState < 2
  ) {

    return;

  }


  canvas.width =
    video.videoWidth;


  canvas.height =
    video.videoHeight;


  ctx.clearRect(
    0,
    0,
    canvas.width,
    canvas.height
  );


  let results;


  try {

    results =
      handsModel.recognizeForVideo(
        video,
        timestamp
      );

  }

  catch (e) {

    return;

  }


  const numHands =
    results.landmarks
      ? results.landmarks.length
      : 0;


  const now =
    Date.now();


  /*
     Draw hand skeletons.
  */


  for (
    let h = 0;
    h < numHands;
    h++
  ) {

    drawHandSkeleton(
      ctx,
      results.landmarks[h],
      canvas.width,
      canvas.height,
      h
    );

  }


  /* ═══════════════════════════════════════════
     TWO HANDS = PINCH ZOOM
     ═══════════════════════════════════════════ */

  if (
    numHands === 2
  ) {

    palmHoldStart =
      0;


    palmTriggeredThisHold =
      false;


    /*
       IMPORTANT:
       We do not reset the discrete gesture
       latch here because the user may move from
       a gesture to two hands.
    */


    const lm0 =
      results.landmarks[0];


    const lm1 =
      results.landmarks[1];


    const tip0 =
      lm0[LM.INDEX_TIP];


    const tip1 =
      lm1[LM.INDEX_TIP];


    const interDist =
      landmarkDist(
        tip0,
        tip1
      );


    if (
      prevPinchDist !== null &&
      now >
        pinchZoomCooldown
    ) {

      const delta =
        interDist -
        prevPinchDist;


      if (
        delta >
        PINCH_THRESHOLD
      ) {

        zoomIn();


        pinchZoomCooldown =
          now + 600;


        flashGestureLabel(
          "🔍 Zoom In"
        );

      }


      else if (
        delta <
        -PINCH_THRESHOLD
      ) {

        zoomOut();


        pinchZoomCooldown =
          now + 600;


        flashGestureLabel(
          "🔍 Zoom Out"
        );

      }
    }


    prevPinchDist =
      interDist;


    /*
       Draw pinch line.
    */


    ctx.beginPath();


    ctx.moveTo(
      tip0.x *
        canvas.width,
      tip0.y *
        canvas.height
    );


    ctx.lineTo(
      tip1.x *
        canvas.width,
      tip1.y *
        canvas.height
    );


    ctx.strokeStyle =
      "rgba(255,200,0,0.7)";


    ctx.lineWidth =
      3;


    ctx.stroke();


    /*
       Clear swipe history.
    */


    State.wristHistory[0] =
      [];


    State.wristHistory[1] =
      [];


    return;
  }


  prevPinchDist =
    null;


  /* ═══════════════════════════════════════════
     ONE HAND
     ═══════════════════════════════════════════ */

  if (
    numHands === 1
  ) {

    const lm =
      results.landmarks[0];


    const wrist =
      lm[LM.WRIST];


    /*
       First check MediaPipe high-level gesture.

       This is BEFORE swipe detection.

       Therefore:
       👍 = next
       ✌️ = previous
       ✊ = exit fullscreen
    */


    const recognized =
      getRecognizedGesture(
        results,
        0
      );


    const discreteHandled =
      handleDiscreteGesture(
        recognized,
        now
      );


    if (
      recognized === "Thumb_Up" ||
      recognized === "Victory" ||
      (
        recognized === "Closed_Fist" &&
        State.presentationMode
      )
    ) {

      /*
         Do not let these gestures
         also become swipes.
      */


      State.wristHistory[0] =
        [];


      palmHoldStart =
        0;


      palmTriggeredThisHold =
        false;


      return;
    }


    /*
       Release the latch when the user
       changes hand shape.
    */


    if (
      recognized !== "Thumb_Up" &&
      recognized !== "Victory" &&
      recognized !== "Closed_Fist"
    ) {

      discreteGestureLatched =
        false;

    }


    /*
       Add wrist position to trajectory.
    */


    const hist =
      State.wristHistory[0];


    hist.push({

      x: wrist.x,
      y: wrist.y,
      t: now

    });


    /*
       Remove old frames.
    */


    while (
      hist.length > 0 &&
      now - hist[0].t >
        State.SWIPE_MAX_TIME +
        100
    ) {

      hist.shift();

    }


    if (
      hist.length >
      State.HISTORY_LEN
    ) {

      hist.shift();

    }


    State.wristHistory[1] =
      [];


    /* ═══════════════════════════════════════
       OPEN PALM = ENTER FULLSCREEN
       ═══════════════════════════════════════ */


    const palm =
      recognized === "Open_Palm" ||
      isOpenPalm(lm);


    if (
      palm
    ) {

      /*
         Start the hold timer.
      */


      if (
        palmHoldStart === 0
      ) {

        palmHoldStart =
          now;


        palmTriggeredThisHold =
          false;

      }


      /*
         After 400ms:
         ENTER fullscreen.

         IMPORTANT:
         Open palm now only ENTERS fullscreen.

         It does NOT exit fullscreen.

         Closed fist exits fullscreen.
      */


      else if (
        !palmTriggeredThisHold &&
        now -
          palmHoldStart >=
          PALM_HOLD_MS
      ) {

        palmTriggeredThisHold =
          true;


        /*
           Only open fullscreen
           if we are currently normal.
        */


        if (
          !State.presentationMode
        ) {

          openPresentation();


          flashGestureLabel(
            "👋 Fullscreen"
          );

        }


        State.lastGestureTime =
          now;

      }


      /*
         Draw hold progress ring.
      */


      const wx =
        wrist.x *
        canvas.width;


      const wy =
        wrist.y *
        canvas.height;


      const progress =
        Math.min(
          1,
          (
            now -
            palmHoldStart
          ) /
          PALM_HOLD_MS
        );


      ctx.beginPath();


      ctx.arc(
        wx,
        wy,
        22,
        -Math.PI / 2,
        -Math.PI / 2 +
          progress *
          2 *
          Math.PI
      );


      ctx.strokeStyle =
        "rgba(110,87,255,0.9)";


      ctx.lineWidth =
        4;


      ctx.stroke();

    }


    else {

      palmHoldStart =
        0;


      palmTriggeredThisHold =
        false;


      /*
         Legacy swipe support.

         This remains available if the user
         wants to swipe with an unrecognized hand.

         But 👍 and ✌️ are handled above first.
      */


      if (
        now -
          State.lastGestureTime >
        State.gestureCooldown
      ) {

        const dir =
          analyseSwipe(hist);


        if (dir) {

          State.lastGestureTime =
            now;


          State.wristHistory[0] =
            [];


          if (
            dir === "right"
          ) {

            nextSlide();


            flashGestureLabel(
              "▶ Next"
            );

          }

          else {

            prevSlide();


            flashGestureLabel(
              "◀ Previous"
            );

          }
        }
      }
    }

  }


  /* ═══════════════════════════════════════════
     NO HANDS
     ═══════════════════════════════════════════ */

  else {

    State.wristHistory[0] =
      [];


    State.wristHistory[1] =
      [];


    palmHoldStart =
      0;


    palmTriggeredThisHold =
      false;


    /*
       IMPORTANT:
       Releasing the hand unlocks the
       next 👍 / ✌️ / ✊ gesture.
    */


    discreteGestureLatched =
      false;

  }
}


/* ─────────────────────────────────────────────
   DRAW HAND
   ───────────────────────────────────────────── */

function drawHandSkeleton(
  ctx,
  lm,
  w,
  h,
  handIdx
) {

  const color =
    handIdx === 0
      ? "rgba(110,87,255,0.85)"
      : "rgba(255,140,0,0.85)";


  const dotColor =
    handIdx === 0
      ? "rgba(180,160,255,1)"
      : "rgba(255,200,80,1)";


  const CONNECTIONS = [

    [0, 1],
    [1, 2],
    [2, 3],
    [3, 4],

    [0, 5],
    [5, 6],
    [6, 7],
    [7, 8],

    [0, 9],
    [9, 10],
    [10, 11],
    [11, 12],

    [0, 13],
    [13, 14],
    [14, 15],
    [15, 16],

    [0, 17],
    [17, 18],
    [18, 19],
    [19, 20],

    [5, 9],
    [9, 13],
    [13, 17],
    [0, 17]

  ];


  ctx.strokeStyle =
    color;


  ctx.lineWidth =
    2;


  for (
    const [a, b]
    of CONNECTIONS
  ) {

    ctx.beginPath();


    ctx.moveTo(
      lm[a].x * w,
      lm[a].y * h
    );


    ctx.lineTo(
      lm[b].x * w,
      lm[b].y * h
    );


    ctx.stroke();

  }


  ctx.fillStyle =
    dotColor;


  for (
    const pt of lm
  ) {

    ctx.beginPath();


    ctx.arc(
      pt.x * w,
      pt.y * h,
      4,
      0,
      Math.PI * 2
    );


    ctx.fill();

  }
}


/* ─────────────────────────────────────────────
   GESTURE LABEL
   ───────────────────────────────────────────── */

function flashGestureLabel(text) {

  const el =
    State.presentationMode
      ? DOM.fsGestureAnim
      : DOM.gestureAnim;


  el.textContent =
    text;


  el.classList.remove(
    "hidden",
    "show"
  );


  void el.offsetWidth;


  el.classList.add(
    "show"
  );


  setTimeout(
    () =>
      el.classList.remove(
        "show"
      ),
    800
  );


  setLastCommand(
    "Gesture: " + text
  );
}


/* ─────────────────────────────────────────────
   CAMERA
   ───────────────────────────────────────────── */

async function startCamera() {

  DOM.cameraError
    .classList
    .add("hidden");


  try {

    const stream =
      await navigator.mediaDevices
        .getUserMedia({

          video: {

            width: {
              ideal: 640
            },

            height: {
              ideal: 480
            },

            facingMode:
              "user"

          },

          audio: false

        });


    State.cameraStream =
      stream;


    DOM.webcam.srcObject =
      stream;


    DOM.cameraOff.style.display =
      "none";


    await new Promise(
      r => {

        DOM.webcam.onloadedmetadata =
          r;

      }
    );


    await DOM.webcam.play();


    showGestureStatus(
      "loading",
      "Loading gesture model…"
    );


    const ok =
      await initHandsModel();


    if (ok) {

      State.gestureActive =
        true;


      setGestureStatus(
        true
      );


      discreteGestureLatched =
        false;


      gestureRAF =
        requestAnimationFrame(
          processGestureFrame
        );

    }

  }


  catch (err) {

    let msg =
      "Camera error.";


    if (
      err.name ===
      "NotAllowedError"
    ) {

      msg =
        "Camera permission denied. Allow access in your browser settings.";

    }


    else if (
      err.name ===
      "NotFoundError"
    ) {

      msg =
        "No camera found on this device.";

    }


    else {

      msg =
        "Camera error: " +
        err.message;

    }


    DOM.cameraError.textContent =
      msg;


    DOM.cameraError
      .classList
      .remove("hidden");


    DOM.cameraOff.style.display =
      "";


    setGestureStatus(
      false
    );

  }
}


function stopCamera() {

  State.gestureActive =
    false;


  if (gestureRAF) {

    cancelAnimationFrame(
      gestureRAF
    );


    gestureRAF =
      null;

  }


  if (
    State.cameraStream
  ) {

    State.cameraStream
      .getTracks()
      .forEach(
        t => t.stop()
      );


    State.cameraStream =
      null;

  }


  DOM.webcam.srcObject =
    null;


  DOM.cameraOff.style.display =
    "";


  setGestureStatus(
    false
  );


  const ctx =
    DOM.gestureCanvas
      .getContext("2d");


  ctx.clearRect(
    0,
    0,
    DOM.gestureCanvas.width,
    DOM.gestureCanvas.height
  );


  discreteGestureLatched =
    false;

}


/* ═══════════════════════════════════════════════
   8. VOICE RECOGNITION
   ═══════════════════════════════════════════════ */

let recognition =
  null;


const VOICE_COMMANDS = [

  {
    triggers: [
      "next",
      "forward",
      "ahead",
      "advance"
    ],
    action:
      nextSlide,
    label:
      "▶ Next"
  },

  {
    triggers: [
      "back",
      "previous",
      "prev",
      "before",
      "light"
    ],
    action:
      prevSlide,
    label:
      "◀ Back"
  },

  {
    triggers: [
      "first",
      "start",
      "beginning",
      "begin",
      "home"
    ],
    action:
      firstSlide,
    label:
      "⏮ First"
  },

  {
    triggers: [
      "last",
      "end",
      "final",
      "finish"
    ],
    action:
      lastSlide,
    label:
      "⏭ Last"
  },

  {
    triggers: [
      "read",
      "speak",
      "say",
      "aloud",
      "tell",
      "narrate"
    ],
    action:
      readCurrentSlide,
    label:
      "🔊 Read"
  },

  {
    triggers: [
      "stop",
      "quiet",
      "silence",
      "mute",
      "cancel",
      "halt"
    ],
    action:
      stopReading,
    label:
      "⏹ Stop"
  },

  {
    triggers: [
      "pause",
      "wait",
      "hold",
      "freeze"
    ],
    action:
      pauseReading,
    label:
      "⏸ Pause"
  },

  {
    triggers: [
      "resume",
      "continue",
      "unpause",
      "play"
    ],
    action:
      resumeReading,
    label:
      "▶ Resume"
  },

  {
    triggers: [
      "repeat",
      "replay",
      "again"
    ],
    action:
      readCurrentSlide,
    label:
      "↺ Repeat"
  },

  {
    triggers: [
      "summarize",
      "summarise",
      "summary",
      "brief",
      "sum"
    ],
    action:
      () =>
        triggerAI("summarize"),
    label:
      "✦ Summarize"
  },

  {
    triggers: [
      "explain",
      "elaborate",
      "detail",
      "describe"
    ],
    action:
      () =>
        triggerAI("explain"),
    label:
      "✦ Explain"
  },

  {
    triggers: [
      "present",
      "fullscreen",
      "full",
      "show",
      "open",
      "launch"
    ],
    action:
      openPresentation,
    label:
      "⤢ Present"
  },

  {
    triggers: [
      "exit",
      "close",
      "quit",
      "leave",
      "escape"
    ],
    action:
      closePresentation,
    label:
      "✕ Exit"
  },

  {
    triggers: [
      "zoom"
    ],
    action:
      zoomIn,
    label:
      "🔍 Zoom"
  }

];


function handleVoiceCommand(raw) {

  const words =
    raw
      .toLowerCase()
      .trim()
      .split(/\s+/);


  let matched =
    null;


  outer:

  for (
    const cmd of VOICE_COMMANDS
  ) {

    for (
      const w of words
    ) {

      if (
        cmd.triggers.includes(w)
      ) {

        matched =
          cmd;

        break outer;

      }
    }
  }


  if (!matched) {

    const phrase =
      words.join(" ");


    for (
      const cmd of VOICE_COMMANDS
    ) {

      if (
        cmd.triggers.some(
          t =>
            phrase.includes(t)
        )
      ) {

        matched =
          cmd;

        break;

      }
    }
  }


  if (matched) {

    logVoice(
      raw,
      true,
      matched.label
    );


    matched.action();


    setLastCommand(
      matched.label
    );

  }

  else {

    logVoice(
      raw,
      false
    );

  }
}


function initVoiceRecognition() {

  const SR =
    window.SpeechRecognition ||
    window.webkitSpeechRecognition;


  if (!SR) {

    DOM.voiceError.textContent =
      "Voice recognition not supported. Use Chrome or Edge.";


    DOM.voiceError
      .classList
      .remove("hidden");


    return false;

  }


  recognition =
    new SR();


  recognition.continuous =
    true;


  recognition.interimResults =
    false;


  recognition.lang =
    "en-US";


  recognition.maxAlternatives =
    3;


  recognition.onresult =
    e => {

      const result =
        e.results[
          e.results.length - 1
        ];


      for (
        let i = 0;
        i < result.length;
        i++
      ) {

        const transcript =
          result[i]
            .transcript
            .trim()
            .toLowerCase();


        const words =
          transcript.split(/\s+/);


        for (
          const cmd
          of VOICE_COMMANDS
        ) {

          for (
            const w
            of words
          ) {

            if (
              cmd.triggers.includes(w)
            ) {

              logVoice(
                transcript,
                true,
                cmd.label
              );


              cmd.action();


              setLastCommand(
                cmd.label
              );


              return;

            }
          }
        }
      }


      logVoice(
        e.results[
          e.results.length - 1
        ][0]
          .transcript
          .trim()
          .toLowerCase(),
        false
      );

    };


  recognition.onerror =
    e => {

      if (
        e.error ===
        "not-allowed"
      ) {

        DOM.voiceError.textContent =
          "Microphone permission denied.";


        DOM.voiceError
          .classList
          .remove("hidden");


        setVoiceStatus(
          false
        );

      }

    };


  recognition.onend =
    () => {

      if (
        State.voiceActive
      ) {

        try {

          recognition.start();

        }

        catch (e) {}

      }

    };


  return true;
}


function startVoice() {

  DOM.voiceError
    .classList
    .add("hidden");


  if (
    !recognition &&
    !initVoiceRecognition()
  ) {

    return;

  }


  try {

    recognition.start();


    State.voiceActive =
      true;


    setVoiceStatus(
      true
    );


    clearVoiceLog();

  }

  catch (e) {

    console.warn(
      "Voice start error:",
      e
    );

  }
}


function stopVoice() {

  State.voiceActive =
    false;


  if (recognition) {

    try {

      recognition.stop();

    }

    catch (e) {}

  }


  setVoiceStatus(
    false
  );
}


function logVoice(
  text,
  recognised,
  label = ""
) {

  const entry =
    document.createElement(
      "div"
    );


  entry.className =
    "voice-log-entry";


  if (recognised) {

    entry.innerHTML =
      `<span class="voice-log-cmd">${escapeHTML(label)}</span>` +
      `<span class="voice-log-raw"> ← "${escapeHTML(text)}"</span>`;

  }

  else {

    entry.innerHTML =
      `<span class="voice-log-miss">? "${escapeHTML(text)}"</span>`;

  }


  const empty =
    DOM.voiceLog.querySelector(
      ".voice-log-empty"
    );


  if (empty) {

    empty.remove();

  }


  DOM.voiceLog.insertBefore(
    entry,
    DOM.voiceLog.firstChild
  );


  while (
    DOM.voiceLog.children.length >
    25
  ) {

    DOM.voiceLog.removeChild(
      DOM.voiceLog.lastChild
    );

  }
}


function clearVoiceLog() {

  DOM.voiceLog.innerHTML =
    '<p class="voice-log-empty">Listening for commands…</p>';

}


/* ═══════════════════════════════════════════════
   9. TEXT-TO-SPEECH
   ═══════════════════════════════════════════════ */

function getSlideText() {

  return (
    pdfTextCache[
      State.current
    ] || null
  );

}


function readCurrentSlide() {

  stopReading();


  const text =
    getSlideText();


  if (!text) {

    speak(
      `Slide ${
        State.current + 1
      } of ${
        State.total
      }. No text found on this slide.`
    );

  }

  else {

    speak(text);

  }
}


function speak(text) {

  if (
    !window.speechSynthesis
  ) return;


  stopReading();


  const utt =
    new SpeechSynthesisUtterance(
      text
    );


  utt.rate =
    0.95;


  utt.pitch =
    1;


  utt.volume =
    1;


  utt.onstart =
    () => {

      State.ttsActive =
        true;


      setTTSStatus(
        "speaking"
      );

    };


  utt.onend =
    () => {

      State.ttsActive =
        false;


      setTTSStatus(
        "idle"
      );

    };


  utt.onerror =
    () => {

      State.ttsActive =
        false;


      setTTSStatus(
        "idle"
      );

    };


  State.ttsUtterance =
    utt;


  window.speechSynthesis
    .speak(utt);
}


function stopReading() {

  window.speechSynthesis?.cancel();


  State.ttsActive =
    false;


  setTTSStatus(
    "idle"
  );
}


function pauseReading() {

  window.speechSynthesis?.pause();


  setTTSStatus(
    "paused"
  );
}


function resumeReading() {

  window.speechSynthesis?.resume();


  setTTSStatus(
    "speaking"
  );
}


/* ═══════════════════════════════════════════════
   10. HUGGING FACE AI
   ═══════════════════════════════════════════════ */


/*
   IMPORTANT:

   The old version used:

   https://api-inference.huggingface.co/models/...

   That endpoint/model combination can produce
   "Failed to fetch" or provider errors.

   This version uses the Hugging Face router
   OpenAI-compatible chat endpoint.
*/


const HF_MODEL =
  "openai/gpt-oss-20b:fastest";


const HF_CHAT_URL =
  "https://router.huggingface.co/v1/chat/completions";


async function callHF(prompt) {

  /*
     A token is required.
  */


  if (
    !State.hfToken
  ) {

    throw new Error(
      "Add your Hugging Face token and click Save. AI requires an HF token."
    );

  }


  const response =
    await fetch(
      HF_CHAT_URL,
      {

        method:
          "POST",


        headers: {

          "Content-Type":
            "application/json",

          "Authorization":
            `Bearer ${State.hfToken}`

        },


        body:
          JSON.stringify({

            model:
              HF_MODEL,


            messages: [

              {

                role:
                  "system",

                content:
                  "You are the SmartSlides AI assistant. Answer clearly and concisely using only the supplied slide content when possible."

              },

              {

                role:
                  "user",

                content:
                  prompt

              }

            ],


            max_tokens:
              300,


            temperature:
              0.5

          })

      }
    );


  let data = {};


  try {

    data =
      await response.json();

  }

  catch (e) {

    throw new Error(
      "Hugging Face returned an invalid response."
    );

  }


  /*
     Authentication errors.
  */


  if (
    response.status === 401 ||
    response.status === 403
  ) {

    throw new Error(
      "Invalid or unauthorized Hugging Face token. Make sure your token has Inference Providers access."
    );

  }


  /*
     Rate / credits.
  */


  if (
    response.status === 429
  ) {

    throw new Error(
      "Hugging Face rate or credit limit reached. Check your Inference Providers credits."
    );

  }


  /*
     Other HTTP errors.
  */


  if (
    !response.ok
  ) {

    const detail =
      data?.error?.message ||
      data?.error ||
      `API error ${response.status}`;


    throw new Error(
      String(detail)
    );

  }


  /*
     Standard OpenAI-compatible response.
  */


  const content =
    data?.choices?.[0]?.message?.content;


  if (
    typeof content ===
      "string" &&
    content.trim()
  ) {

    return content.trim();

  }


  /*
     Some providers may return
     structured content blocks.
  */


  if (
    Array.isArray(content)
  ) {

    const text =
      content
        .map(
          x =>
            x?.text || ""
        )
        .join("")
        .trim();


    if (text) {

      return text;

    }
  }


  /*
     Hugging Face error response.
  */


  if (
    data?.error
  ) {

    throw new Error(
      typeof data.error ===
        "string"
        ? data.error
        : JSON.stringify(
            data.error
          )
    );

  }


  throw new Error(
    "The AI provider returned an unexpected response."
  );
}


/* ─────────────────────────────────────────────
   AI TRIGGER
   ───────────────────────────────────────────── */

async function triggerAI(
  mode,
  question = ""
) {

  if (
    !State.loaded
  ) {

    showAIResult(
      "Load a presentation first."
    );


    return;

  }


  const text =
    pdfTextCache[
      State.current
    ] || "";


  if (!text) {

    showAIResult(
      "No text found on this slide. AI works with PPTX files that have text content."
    );


    return;

  }


  State.aiMode =
    mode;


  DOM.aiResult
    .classList
    .remove("hidden");


  DOM.aiResultText.textContent =
    "";


  DOM.aiLoading
    .classList
    .remove("hidden");


  DOM.aiResultText
    .classList
    .add("hidden");


  const ctx =
    `Slide ${
      State.current + 1
    }:\n"${text.slice(
      0,
      800
    )}"`;


  const prompts = {

    summarize:
      `Summarize this slide in 2-3 clear sentences.\n\n${ctx}`,

    explain:
      `Explain this slide clearly for a student. Use simple language and focus on the important points.\n\n${ctx}`,

    question:
      `Answer this question based only on the slide content.\n\n${ctx}\n\nQuestion: ${question}`

  };


  try {

    const ans =
      await callHF(
        prompts[mode]
      );


    showAIResult(
      ans
    );

  }

  catch (e) {

    console.error(
      "AI request failed:",
      e
    );


    showAIResult(
      "⚠ " +
      e.message
    );

  }

  finally {

    DOM.aiLoading
      .classList
      .add("hidden");


    DOM.aiResultText
      .classList
      .remove("hidden");

  }
}


function showAIResult(text) {

  DOM.aiResult
    .classList
    .remove("hidden");


  DOM.aiResultText.textContent =
    text;


  DOM.aiLoading
    .classList
    .add("hidden");


  DOM.aiResultText
    .classList
    .remove("hidden");
}


/* ═══════════════════════════════════════════════
   11. PRESENTATION / FULLSCREEN
   ═══════════════════════════════════════════════ */


/*
   OPEN PRESENTATION

   Open palm:
   Normal screen
       ↓
   hold 👋
       ↓
   Presentation overlay
       ↓
   browser fullscreen when allowed
*/


async function openPresentation() {

  if (
    !State.loaded
  ) {

    alert(
      "Load a PPTX file first."
    );


    return;

  }


  /*
     Already in presentation mode?
     Do nothing.
  */


  if (
    State.presentationMode
  ) {

    return;

  }


  State.presentationMode =
    true;


  DOM.presentationOverlay
    .classList
    .remove("hidden");


  renderFullscreen(
    State.current
  );


  /*
     Request native browser fullscreen.

     Browser may reject this when the call
     originates from a camera gesture instead
     of a direct user click. The presentation
     overlay still works in that case.
  */


  try {

    if (
      !document.fullscreenElement
    ) {

      await DOM.presentationOverlay
        .requestFullscreen();

    }

  }

  catch (e) {

    console.warn(
      "Native fullscreen request was blocked:",
      e
    );

  }
}


/*
   CLOSE PRESENTATION

   Closed fist:
   Fullscreen
       ↓
   ✊
       ↓
   Normal screen
*/


async function closePresentation() {

  State.presentationMode =
    false;


  DOM.presentationOverlay
    .classList
    .add("hidden");


  try {

    if (
      document.fullscreenElement
    ) {

      await document.exitFullscreen();

    }

  }

  catch (e) {

    console.warn(
      "Exit fullscreen failed:",
      e
    );

  }
}


/*
   Browser Esc / native fullscreen exit.

   If the user presses Esc, make sure the app
   also returns to normal presentation state.
*/


document.addEventListener(
  "fullscreenchange",
  () => {

    if (
      !document.fullscreenElement &&
      State.presentationMode
    ) {

      State.presentationMode =
        false;


      DOM.presentationOverlay
        .classList
        .add("hidden");


      /*
         Reset gesture latch so the next
         hand gesture can work.
      */


      discreteGestureLatched =
        false;

    }

  }
);


/* ═══════════════════════════════════════════════
   12. STATUS HELPERS
   ═══════════════════════════════════════════════ */

function setVoiceStatus(on) {

  DOM.voiceDot.className =
    "status-dot" +
    (
      on
        ? " listening"
        : ""
    );


  DOM.voiceLabel.textContent =
    on
      ? "Voice: listening"
      : "Voice: off";
}


function setGestureStatus(on) {

  DOM.gestureDot.className =
    "status-dot" +
    (
      on
        ? " active"
        : ""
    );


  DOM.gestureLabel.textContent =
    on
      ? "Gesture: active"
      : "Gesture: off";
}


function showGestureStatus(
  type,
  text
) {

  DOM.gestureDot.className =
    "status-dot" +
    (
      type === "active"
        ? " active"
        : type === "loading"
          ? " listening"
          : ""
    );


  DOM.gestureLabel.textContent =
    text;
}


function setTTSStatus(s) {

  const cls = {

    speaking:
      " speaking",

    paused:
      " listening",

    idle:
      ""

  };


  DOM.ttsDot.className =
    "status-dot" +
    (
      cls[s] || ""
    );


  DOM.ttsLabel.textContent =
    "TTS: " + s;
}


function setLastCommand(
  cmd
) {

  DOM.lastCommandDisplay.textContent =
    "Last: " + cmd;
}


function showProgress(
  pct,
  label
) {

  DOM.uploadProgress
    .classList
    .remove("hidden");


  DOM.progressFill.style.width =
    pct + "%";


  DOM.progressLabel.textContent =
    label;
}


function showGestureError(
  msg
) {

  DOM.cameraError.textContent =
    msg;


  DOM.cameraError
    .classList
    .remove("hidden");
}


function escapeHTML(s) {

  return String(s)

    .replace(
      /&/g,
      "&amp;"
    )

    .replace(
      /</g,
      "&lt;"
    )

    .replace(
      />/g,
      "&gt;"
    );

}


/* ═══════════════════════════════════════════════
   13. TOUCH SWIPE
   ═══════════════════════════════════════════════ */

(function initSwipe() {

  let startX = 0;

  let startY = 0;


  DOM.swipeOverlay
    .addEventListener(
      "touchstart",
      e => {

        startX =
          e.touches[0]
            .clientX;


        startY =
          e.touches[0]
            .clientY;

      },
      {
        passive: true
      }
    );


  DOM.swipeOverlay
    .addEventListener(
      "touchend",
      e => {

        const dx =
          e.changedTouches[0]
            .clientX -
          startX;


        const dy =
          e.changedTouches[0]
            .clientY -
          startY;


        if (
          Math.abs(dx) > 50 &&
          Math.abs(dx) >
            Math.abs(dy)
        ) {

          dx < 0
            ? nextSlide()
            : prevSlide();

        }

      },
      {
        passive: true
      }
    );

})();


/* ═══════════════════════════════════════════════
   14. EVENT WIRING
   ═══════════════════════════════════════════════ */

function wire() {

  /* ─────────────────────────────────────────
     FILE UPLOAD
     ───────────────────────────────────────── */

  DOM.fileInput
    .addEventListener(
      "change",
      e =>
        handleFiles(
          e.target.files
        )
    );


  DOM.dropZone
    .addEventListener(
      "dragover",
      e => {

        e.preventDefault();


        DOM.dropZone
          .classList
          .add("drag-over");

      }
    );


  DOM.dropZone
    .addEventListener(
      "dragleave",
      () =>
        DOM.dropZone
          .classList
          .remove(
            "drag-over"
          )
    );


  DOM.dropZone
    .addEventListener(
      "drop",
      e => {

        e.preventDefault();


        DOM.dropZone
          .classList
          .remove(
            "drag-over"
          );


        handleFiles(
          e.dataTransfer.files
        );

      }
    );


  DOM.dropZone
    .addEventListener(
      "click",
      e => {

        if (
          !e.target.closest("label") &&
          !e.target.closest("input")
        ) {

          DOM.fileInput.click();

        }

      }
    );


  /* ─────────────────────────────────────────
     NAVIGATION
     ───────────────────────────────────────── */

  DOM.btnFirst
    .addEventListener(
      "click",
      firstSlide
    );


  DOM.btnPrev
    .addEventListener(
      "click",
      prevSlide
    );


  DOM.btnNext
    .addEventListener(
      "click",
      nextSlide
    );


  DOM.btnLast
    .addEventListener(
      "click",
      lastSlide
    );


  /* ─────────────────────────────────────────
     KEYBOARD
     ───────────────────────────────────────── */

  document.addEventListener(
    "keydown",
    e => {

      if (
        e.target.tagName ===
          "INPUT" ||
        e.target.tagName ===
          "TEXTAREA"
      ) {

        return;

      }


      const map = {

        ArrowRight:
          nextSlide,

        ArrowDown:
          nextSlide,

        PageDown:
          nextSlide,

        ArrowLeft:
          prevSlide,

        ArrowUp:
          prevSlide,

        PageUp:
          prevSlide,

        Home:
          firstSlide,

        End:
          lastSlide,

        "+":
          zoomIn,

        "=":
          zoomIn,

        "-":
          zoomOut,

        Escape:
          closePresentation,

        F5:
          () => {

            e.preventDefault();

            openPresentation();

          }

      };


      if (
        map[e.key]
      ) {

        map[e.key]();

      }

    }
  );


  /* ─────────────────────────────────────────
     NORMAL FULLSCREEN BUTTON
     ───────────────────────────────────────── */

  DOM.btnFullscreen
    .addEventListener(
      "click",
      async () => {

        try {

          if (
            !document.fullscreenElement
          ) {

            await DOM.viewerContainer
              .requestFullscreen();

          }

          else {

            await document
              .exitFullscreen();

          }

        }

        catch (e) {

          console.warn(
            "Fullscreen error:",
            e
          );

        }

      }
    );


  /* ─────────────────────────────────────────
     PRESENTATION
     ───────────────────────────────────────── */

  DOM.btnStartPresent
    .addEventListener(
      "click",
      openPresentation
    );


  DOM.fsBtnPrev
    .addEventListener(
      "click",
      prevSlide
    );


  DOM.fsBtnNext
    .addEventListener(
      "click",
      nextSlide
    );


  DOM.fsBtnExit
    .addEventListener(
      "click",
      closePresentation
    );


  /* ─────────────────────────────────────────
     CAMERA
     ───────────────────────────────────────── */

  DOM.btnStartCamera
    .addEventListener(
      "click",
      startCamera
    );


  DOM.btnStopCamera
    .addEventListener(
      "click",
      stopCamera
    );


  /* ─────────────────────────────────────────
     VOICE
     ───────────────────────────────────────── */

  DOM.btnStartVoice
    .addEventListener(
      "click",
      startVoice
    );


  DOM.btnStopVoice
    .addEventListener(
      "click",
      stopVoice
    );


  /* ─────────────────────────────────────────
     TEXT TO SPEECH
     ───────────────────────────────────────── */

  DOM.btnReadSlide
    .addEventListener(
      "click",
      readCurrentSlide
    );


  DOM.btnPauseRead
    .addEventListener(
      "click",
      pauseReading
    );


  DOM.btnResumeRead
    .addEventListener(
      "click",
      resumeReading
    );


  DOM.btnStopRead
    .addEventListener(
      "click",
      stopReading
    );


  DOM.btnRepeatRead
    .addEventListener(
      "click",
      readCurrentSlide
    );


  /* ═══════════════════════════════════════
     AI TOKEN
     ═══════════════════════════════════════ */

  DOM.btnSaveToken
    .addEventListener(
      "click",
      () => {

        State.hfToken =
          DOM.hfToken.value.trim();


        if (
          State.hfToken
        ) {

          sessionStorage.setItem(
            "hf_token",
            State.hfToken
          );

        }


        DOM.btnSaveToken.textContent =
          "✓ Saved";


        setTimeout(
          () => {

            DOM.btnSaveToken.textContent =
              "Save";

          },
          2000
        );

      }
    );


  /* ─────────────────────────────────────────
     AI BUTTONS
     ───────────────────────────────────────── */

  DOM.btnSummarize
    .addEventListener(
      "click",
      () =>
        triggerAI(
          "summarize"
        )
    );


  DOM.btnExplain
    .addEventListener(
      "click",
      () =>
        triggerAI(
          "explain"
        )
    );


  DOM.btnQuestion
    .addEventListener(
      "click",
      () => {

        DOM.aiQuestionBox
          .classList
          .toggle(
            "hidden"
          );


        if (
          !DOM.aiQuestionBox
            .classList
            .contains("hidden")
        ) {

          DOM.aiQuestionInput
            .focus();

        }

      }
    );


  DOM.btnAskQuestion
    .addEventListener(
      "click",
      () => {

        const q =
          DOM.aiQuestionInput
            .value
            .trim();


        if (q) {

          triggerAI(
            "question",
            q
          );

        }

      }
    );


  DOM.aiQuestionInput
    .addEventListener(
      "keydown",
      e => {

        if (
          e.key === "Enter"
        ) {

          DOM.btnAskQuestion
            .click();

        }

      }
    );


  DOM.btnCloseAI
    .addEventListener(
      "click",
      () =>
        DOM.aiResult
          .classList
          .add("hidden")
    );


  /* ─────────────────────────────────────────
     HELP
     ───────────────────────────────────────── */

  DOM.btnHelp
    .addEventListener(
      "click",
      () =>
        DOM.helpModal
          .classList
          .remove("hidden")
    );


  DOM.btnCloseHelp
    .addEventListener(
      "click",
      () =>
        DOM.helpModal
          .classList
          .add("hidden")
    );


  DOM.helpModal
    .addEventListener(
      "click",
      e => {

        if (
          e.target ===
          DOM.helpModal
        ) {

          DOM.helpModal
            .classList
            .add("hidden");

        }

      }
    );

}


/* ═══════════════════════════════════════════════
   15. INITIALIZATION
   ═══════════════════════════════════════════════ */

(function init() {

  wire();


  if (
    !window.SpeechRecognition &&
    !window.webkitSpeechRecognition
  ) {

    DOM.voiceError.textContent =
      "Voice recognition needs Chrome or Edge.";


    DOM.voiceError
      .classList
      .remove("hidden");

  }


  if (
    !window.speechSynthesis
  ) {

    DOM.ttsLabel.textContent =
      "TTS: unsupported";

  }


  /*
     Restore HF token for this browser session.
  */


  const saved =
    sessionStorage.getItem(
      "hf_token"
    );


  if (saved) {

    State.hfToken =
      saved;


    DOM.hfToken.value =
      saved;

  }


  console.log(
    "SmartSlides ready."
  );


  console.log(
    "Gestures: 👍 Next | ✌️ Previous | 👋 Fullscreen | ✊ Exit Fullscreen"
  );

})();