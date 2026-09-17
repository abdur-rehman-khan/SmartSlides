/* ═══════════════════════════════════════════════════════════════
   SMART SLIDES — Main Script
   Modules: State · PDF · Images · Renderer · Gesture · Voice ·
            TTS · AI (HF) · UI Wiring
   ═══════════════════════════════════════════════════════════════ */

"use strict";

/* ─────────────────────────────────────────────
   1. APP STATE
   ───────────────────────────────────────────── */
const State = {
  slides: [],          // { type: 'canvas'|'img', data: canvas|url }
  current: 0,
  total: 0,
  loaded: false,

  ttsActive: false,
  ttsUtterance: null,

  voiceActive: false,
  gestureActive: false,
  cameraStream: null,

  hfToken: '',
  aiMode: null,        // 'summarize'|'explain'|'question'

  gestureRecognizer: null,
  lastGestureTime: 0,
  lastGestureDir: null,
  gestureCooldown: 800, // ms

  presentationMode: false,
};

/* ─────────────────────────────────────────────
   2. DOM REFS
   ───────────────────────────────────────────── */
const $ = id => document.getElementById(id);
const DOM = {
  fileInput:       $('fileInput'),
  dropZone:        $('dropZone'),
  uploadProgress:  $('uploadProgress'),
  progressFill:    $('progressFill'),
  progressLabel:   $('progressLabel'),

  viewerEmpty:     $('viewerEmpty'),
  slideCanvas:     $('slideCanvas'),
  slideImg:        $('slideImg'),
  swipeOverlay:    $('swipeOverlay'),
  gestureAnim:     $('gestureAnim'),
  viewerContainer: $('viewerContainer'),

  currentSlideNum: $('currentSlideNum'),
  totalSlideNum:   $('totalSlideNum'),

  btnFirst:        $('btnFirst'),
  btnPrev:         $('btnPrev'),
  btnNext:         $('btnNext'),
  btnLast:         $('btnLast'),
  btnFullscreen:   $('btnFullscreen'),
  btnStartPresent: $('btnStartPresent'),

  webcam:          $('webcam'),
  gestureCanvas:   $('gestureCanvas'),
  cameraOff:       $('cameraOff'),
  btnStartCamera:  $('btnStartCamera'),
  btnStopCamera:   $('btnStopCamera'),
  cameraError:     $('cameraError'),

  btnStartVoice:   $('btnStartVoice'),
  btnStopVoice:    $('btnStopVoice'),
  voiceError:      $('voiceError'),
  voiceLog:        $('voiceLog'),

  voiceDot:        $('voiceDot'),
  voiceLabel:      $('voiceLabel'),
  gestureDot:      $('gestureDot'),
  gestureLabel:    $('gestureLabel'),
  ttsDot:          $('ttsDot'),
  ttsLabel:        $('ttsLabel'),
  lastCommandDisplay: $('lastCommandDisplay'),

  btnReadSlide:    $('btnReadSlide'),
  btnPauseRead:    $('btnPauseRead'),
  btnResumeRead:   $('btnResumeRead'),
  btnStopRead:     $('btnStopRead'),
  btnRepeatRead:   $('btnRepeatRead'),

  hfToken:         $('hfToken'),
  btnSaveToken:    $('btnSaveToken'),
  btnSummarize:    $('btnSummarize'),
  btnExplain:      $('btnExplain'),
  btnQuestion:     $('btnQuestion'),
  aiQuestionBox:   $('aiQuestionBox'),
  aiQuestionInput: $('aiQuestionInput'),
  btnAskQuestion:  $('btnAskQuestion'),
  aiResult:        $('aiResult'),
  aiResultText:    $('aiResultText'),
  aiLoading:       $('aiLoading'),
  btnCloseAI:      $('btnCloseAI'),

  thumbnailStrip:  $('thumbnailStrip'),

  helpModal:       $('helpModal'),
  btnHelp:         $('btnHelp'),
  btnCloseHelp:    $('btnCloseHelp'),

  // Fullscreen
  presentationOverlay: $('presentationOverlay'),
  fsCanvas:        $('fsCanvas'),
  fsImg:           $('fsImg'),
  fsBtnPrev:       $('fsBtnPrev'),
  fsBtnNext:       $('fsBtnNext'),
  fsBtnExit:       $('fsBtnExit'),
  fsCounter:       $('fsCounter'),
  fsGestureAnim:   $('fsGestureAnim'),
};

/* ─────────────────────────────────────────────
   3. PDF LOADING
   ───────────────────────────────────────────── */
async function loadPDF(file) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  showProgress(0, 'Reading PDF…');
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const pages = [];

  for (let i = 1; i <= pdf.numPages; i++) {
    showProgress((i / pdf.numPages) * 100, `Rendering page ${i} of ${pdf.numPages}…`);
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 2.0 });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');
    await page.render({ canvasContext: ctx, viewport }).promise;
    pages.push({ type: 'canvas', data: canvas });
  }

  return pages;
}

/* ─────────────────────────────────────────────
   4. IMAGE LOADING
   ───────────────────────────────────────────── */
function loadImages(files) {
  const imageFiles = Array.from(files).filter(f => f.type.startsWith('image/'));
  return new Promise(resolve => {
    const slides = [];
    let loaded = 0;
    if (!imageFiles.length) { resolve([]); return; }
    imageFiles.forEach((f, idx) => {
      showProgress((idx / imageFiles.length) * 100, `Loading image ${idx + 1}…`);
      const url = URL.createObjectURL(f);
      slides.push({ type: 'img', data: url });
      loaded++;
      if (loaded === imageFiles.length) resolve(slides);
    });
  });
}

/* ─────────────────────────────────────────────
   5. SLIDE RENDERER
   ───────────────────────────────────────────── */
function renderSlide(index) {
  if (!State.loaded || !State.slides.length) return;
  const slide = State.slides[index];
  State.current = index;

  DOM.viewerEmpty.classList.add('hidden');
  DOM.slideCanvas.classList.add('hidden');
  DOM.slideImg.classList.add('hidden');

  if (slide.type === 'canvas') {
    DOM.slideCanvas.classList.remove('hidden');
    const ctx = DOM.slideCanvas.getContext('2d');
    DOM.slideCanvas.width = slide.data.width;
    DOM.slideCanvas.height = slide.data.height;
    ctx.drawImage(slide.data, 0, 0);
  } else {
    DOM.slideImg.classList.remove('hidden');
    DOM.slideImg.src = slide.data;
  }

  // Update counter
  DOM.currentSlideNum.textContent = index + 1;
  DOM.totalSlideNum.textContent = State.total;

  // Update fullscreen if open
  if (State.presentationMode) renderFullscreen(index);

  // Update thumbnails
  updateThumbnailActive(index);
}

function renderFullscreen(index) {
  if (!State.slides.length) return;
  const slide = State.slides[index];
  DOM.fsCanvas.classList.add('hidden');
  DOM.fsImg.classList.add('hidden');

  if (slide.type === 'canvas') {
    DOM.fsCanvas.classList.remove('hidden');
    const ctx = DOM.fsCanvas.getContext('2d');
    DOM.fsCanvas.width = slide.data.width;
    DOM.fsCanvas.height = slide.data.height;
    ctx.drawImage(slide.data, 0, 0);
  } else {
    DOM.fsImg.classList.remove('hidden');
    DOM.fsImg.src = slide.data;
  }
  DOM.fsCounter.textContent = `${index + 1} / ${State.total}`;
}

function goToSlide(n) {
  if (!State.loaded) return;
  n = Math.max(0, Math.min(n, State.total - 1));
  renderSlide(n);
}
function nextSlide() { goToSlide(State.current + 1); }
function prevSlide() { goToSlide(State.current - 1); }
function firstSlide() { goToSlide(0); }
function lastSlide()  { goToSlide(State.total - 1); }

/* ─────────────────────────────────────────────
   6. THUMBNAIL STRIP
   ───────────────────────────────────────────── */
function buildThumbnails() {
  DOM.thumbnailStrip.innerHTML = '';
  State.slides.forEach((slide, i) => {
    const item = document.createElement('div');
    item.className = 'thumb-item' + (i === 0 ? ' active' : '');
    item.dataset.index = i;

    if (slide.type === 'canvas') {
      const c = document.createElement('canvas');
      c.width = 160; c.height = 90;
      c.getContext('2d').drawImage(slide.data, 0, 0, 160, 90);
      item.appendChild(c);
    } else {
      const img = document.createElement('img');
      img.src = slide.data;
      item.appendChild(img);
    }

    const num = document.createElement('span');
    num.className = 'thumb-num';
    num.textContent = i + 1;
    item.appendChild(num);

    item.addEventListener('click', () => goToSlide(i));
    DOM.thumbnailStrip.appendChild(item);
  });
}

function updateThumbnailActive(index) {
  document.querySelectorAll('.thumb-item').forEach((el, i) => {
    el.classList.toggle('active', i === index);
    if (i === index) el.scrollIntoView({ block: 'nearest' });
  });
}

/* ─────────────────────────────────────────────
   7. FILE UPLOAD HANDLER
   ───────────────────────────────────────────── */
async function handleFiles(files) {
  if (!files || !files.length) return;
  DOM.uploadProgress.classList.remove('hidden');

  try {
    let slides = [];
    const hasPDF = Array.from(files).some(f => f.type === 'application/pdf');

    if (hasPDF) {
      const pdfFile = Array.from(files).find(f => f.type === 'application/pdf');
      slides = await loadPDF(pdfFile);
      // Also load any images alongside
      const imgs = await loadImages(files);
      slides = slides.concat(imgs);
    } else {
      slides = await loadImages(files);
    }

    if (!slides.length) {
      showError('No valid slides found. Please upload a PDF or image files.');
      return;
    }

    State.slides = slides;
    State.total = slides.length;
    State.loaded = true;

    buildThumbnails();
    renderSlide(0);
    showProgress(100, `Loaded ${slides.length} slide${slides.length !== 1 ? 's' : ''}!`);

    setTimeout(() => DOM.uploadProgress.classList.add('hidden'), 2000);
  } catch (err) {
    showError(`Failed to load presentation: ${err.message}`);
    console.error(err);
  }
}

/* ─────────────────────────────────────────────
   8. GESTURE RECOGNITION (MediaPipe)
   ───────────────────────────────────────────── */
let gestureAnimFrame = null;

async function initGestureRecognizer() {
  if (!window._MPImports) {
    showGestureError('MediaPipe not loaded. Gesture detection unavailable.');
    return false;
  }
  const { GestureRecognizer, FilesetResolver } = window._MPImports;
  try {
    const vision = await FilesetResolver.forVisionTasks(
      'https://cdn.jsdelivr.net/npm/@mediapipe/tasks-vision@0.10.3/wasm'
    );
    State.gestureRecognizer = await GestureRecognizer.createFromOptions(vision, {
      baseOptions: {
        modelAssetPath:
          'https://storage.googleapis.com/mediapipe-models/gesture_recognizer/gesture_recognizer/float16/1/gesture_recognizer.task',
        delegate: 'GPU',
      },
      runningMode: 'VIDEO',
      numHands: 1,
    });
    return true;
  } catch (err) {
    console.warn('Gesture recognizer init failed:', err);
    showGestureError('Gesture detection failed to initialise. Hand tracking unavailable.');
    return false;
  }
}

function startGestureLoop() {
  const video = DOM.webcam;
  const canvas = DOM.gestureCanvas;
  const ctx = canvas.getContext('2d');

  let prevX = null;

  function processFrame(timestamp) {
    if (!State.gestureActive || !State.gestureRecognizer) return;
    gestureAnimFrame = requestAnimationFrame(processFrame);

    if (video.readyState < 2) return;

    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    try {
      const results = State.gestureRecognizer.recognizeForVideo(video, timestamp);

      if (results.landmarks && results.landmarks.length > 0) {
        // Draw landmarks
        const lm = results.landmarks[0];
        drawLandmarks(ctx, lm, canvas.width, canvas.height);

        // Track wrist X movement (landmark 0 = wrist)
        const wristX = lm[0].x; // 0..1, normalized

        if (prevX !== null) {
          const delta = wristX - prevX;
          const now = Date.now();

          if (Math.abs(delta) > 0.08 && (now - State.lastGestureTime) > State.gestureCooldown) {
            const dir = delta > 0 ? 'right' : 'left';
            State.lastGestureTime = now;
            State.lastGestureDir = dir;

            // Mirrored video: moving right = swipe left visually
            if (dir === 'right') {
              prevSlide();
              flashGestureAnim('◀ Prev', true);
            } else {
              nextSlide();
              flashGestureAnim('Next ▶', false);
            }
          }
        }
        prevX = wristX;
      } else {
        prevX = null;
      }
    } catch (e) {
      // silently ignore per-frame errors
    }
  }

  gestureAnimFrame = requestAnimationFrame(processFrame);
}

function drawLandmarks(ctx, landmarks, w, h) {
  ctx.fillStyle = 'rgba(110, 87, 255, 0.8)';
  landmarks.forEach(pt => {
    ctx.beginPath();
    ctx.arc(pt.x * w, pt.y * h, 4, 0, Math.PI * 2);
    ctx.fill();
  });
  // Draw skeleton lines
  const connections = [
    [0,1],[1,2],[2,3],[3,4],
    [0,5],[5,6],[6,7],[7,8],
    [0,9],[9,10],[10,11],[11,12],
    [0,13],[13,14],[14,15],[15,16],
    [0,17],[17,18],[18,19],[19,20],
    [5,9],[9,13],[13,17],
  ];
  ctx.strokeStyle = 'rgba(110, 87, 255, 0.4)';
  ctx.lineWidth = 2;
  connections.forEach(([a, b]) => {
    ctx.beginPath();
    ctx.moveTo(landmarks[a].x * w, landmarks[a].y * h);
    ctx.lineTo(landmarks[b].x * w, landmarks[b].y * h);
    ctx.stroke();
  });
}

function flashGestureAnim(text, prev) {
  const el = State.presentationMode ? DOM.fsGestureAnim : DOM.gestureAnim;
  el.textContent = text;
  el.classList.remove('hidden', 'show');
  // Trigger reflow
  void el.offsetWidth;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 700);
  setLastCommand(`Gesture: ${text}`);
}

async function startCamera() {
  DOM.cameraError.classList.add('hidden');
  try {
    const stream = await navigator.mediaDevices.getUserMedia({
      video: { width: 640, height: 480, facingMode: 'user' },
      audio: false,
    });
    State.cameraStream = stream;
    DOM.webcam.srcObject = stream;
    DOM.cameraOff.style.display = 'none';

    // Wait for video to be ready
    await new Promise(r => { DOM.webcam.onloadedmetadata = r; });

    const ok = await initGestureRecognizer();
    if (ok) {
      State.gestureActive = true;
      setGestureStatus(true);
      startGestureLoop();
    }
  } catch (err) {
    let msg = 'Camera error.';
    if (err.name === 'NotAllowedError') msg = 'Camera permission denied. Please allow camera access.';
    else if (err.name === 'NotFoundError') msg = 'No camera found on this device.';
    else msg = `Camera error: ${err.message}`;
    DOM.cameraError.textContent = msg;
    DOM.cameraError.classList.remove('hidden');
    DOM.cameraOff.style.display = '';
  }
}

function stopCamera() {
  State.gestureActive = false;
  if (gestureAnimFrame) cancelAnimationFrame(gestureAnimFrame);
  if (State.cameraStream) {
    State.cameraStream.getTracks().forEach(t => t.stop());
    State.cameraStream = null;
  }
  DOM.webcam.srcObject = null;
  DOM.cameraOff.style.display = '';
  setGestureStatus(false);

  const ctx = DOM.gestureCanvas.getContext('2d');
  ctx.clearRect(0, 0, DOM.gestureCanvas.width, DOM.gestureCanvas.height);
}

/* ─────────────────────────────────────────────
   9. VOICE RECOGNITION
   ───────────────────────────────────────────── */
let recognition = null;

function initVoiceRecognition() {
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SR) {
    DOM.voiceError.textContent = 'Speech recognition not supported in this browser. Try Chrome or Edge.';
    DOM.voiceError.classList.remove('hidden');
    return false;
  }

  recognition = new SR();
  recognition.continuous = true;
  recognition.interimResults = false;
  recognition.lang = 'en-US';

  recognition.onresult = (e) => {
    const transcript = e.results[e.results.length - 1][0].transcript.trim().toLowerCase();
    handleVoiceCommand(transcript);
  };

  recognition.onerror = (e) => {
    if (e.error === 'not-allowed') {
      DOM.voiceError.textContent = 'Microphone permission denied.';
      DOM.voiceError.classList.remove('hidden');
      setVoiceStatus(false);
    } else if (e.error === 'no-speech') {
      // normal timeout, just restart
    } else {
      console.warn('Speech error:', e.error);
    }
  };

  recognition.onend = () => {
    if (State.voiceActive) {
      try { recognition.start(); } catch(e) {}
    }
  };

  return true;
}

function handleVoiceCommand(raw) {
  logVoice(raw);
  const r = raw;

  if (r.includes('next slide') || r.includes('next page')) {
    nextSlide(); setLastCommand('next slide');
  } else if (r.includes('previous slide') || r.includes('prev slide') || r.includes('go back')) {
    prevSlide(); setLastCommand('previous slide');
  } else if (r.includes('first slide') || r.includes('go to start')) {
    firstSlide(); setLastCommand('first slide');
  } else if (r.includes('last slide') || r.includes('go to end')) {
    lastSlide(); setLastCommand('last slide');
  } else if (r.includes('read this slide') || r.includes('read slide') || r.includes('read this')) {
    readCurrentSlide(); setLastCommand('read slide');
  } else if (r.includes('stop reading') || r.includes('stop read')) {
    stopReading(); setLastCommand('stop reading');
  } else if (r.includes('pause reading') || r === 'pause') {
    pauseReading(); setLastCommand('pause');
  } else if (r.includes('resume reading') || r === 'resume') {
    resumeReading(); setLastCommand('resume');
  } else if (r.includes('repeat slide') || r.includes('repeat')) {
    readCurrentSlide(); setLastCommand('repeat slide');
  } else if (r.includes('summarize') || r.includes('summarise')) {
    triggerAI('summarize'); setLastCommand('summarize slide');
  } else if (r.includes('explain')) {
    triggerAI('explain'); setLastCommand('explain slide');
  } else if (r.includes('start presentation')) {
    openPresentation(); setLastCommand('start presentation');
  } else if (r.includes('stop presentation') || r.includes('exit presentation')) {
    closePresentation(); setLastCommand('stop presentation');
  } else {
    logVoice(raw, false); // unrecognised
  }
}

function startVoice() {
  DOM.voiceError.classList.add('hidden');
  if (!recognition) {
    if (!initVoiceRecognition()) return;
  }
  try {
    recognition.start();
    State.voiceActive = true;
    setVoiceStatus(true);
    clearVoiceLog();
  } catch(e) {
    console.warn('Voice start error:', e);
  }
}

function stopVoice() {
  State.voiceActive = false;
  if (recognition) {
    try { recognition.stop(); } catch(e) {}
  }
  setVoiceStatus(false);
}

function logVoice(text, recognized = true) {
  const entry = document.createElement('div');
  entry.className = 'voice-log-entry';
  if (recognized) {
    entry.innerHTML = `<span class="voice-log-cmd">▸ ${escapeHTML(text)}</span>`;
  } else {
    entry.innerHTML = `<span class="voice-log-raw">heard: ${escapeHTML(text)}</span>`;
  }
  const empty = DOM.voiceLog.querySelector('.voice-log-empty');
  if (empty) empty.remove();
  DOM.voiceLog.insertBefore(entry, DOM.voiceLog.firstChild);
  // Keep max 20 entries
  while (DOM.voiceLog.children.length > 20) DOM.voiceLog.removeChild(DOM.voiceLog.lastChild);
}

function clearVoiceLog() {
  DOM.voiceLog.innerHTML = '<p class="voice-log-empty">Listening for commands…</p>';
}

/* ─────────────────────────────────────────────
   10. TEXT-TO-SPEECH
   ───────────────────────────────────────────── */
function extractSlideText() {
  if (!State.loaded) return '';
  const slide = State.slides[State.current];

  if (slide.type === 'canvas') {
    // OCR via canvas is not feasible client-side without a library;
    // we extract text from a hidden div if user has pasted text there,
    // or we notify user
    return ''; // PDF text extraction below
  }
  return '';
}

// For PDF slides we extract text using pdf.js text layer separately
const pdfTextCache = {};

async function extractPDFText(slideIndex) {
  if (pdfTextCache[slideIndex] !== undefined) return pdfTextCache[slideIndex];
  pdfTextCache[slideIndex] = ''; // default empty
  return '';
}

function getSlideTextForTTS() {
  // Best effort: if it's an image-based slide we cannot easily OCR
  // We'll speak a placeholder message for image slides, and try
  // text for canvas (PDF-rendered) slides via a cached text layer
  const cached = pdfTextCache[State.current];
  if (cached) return cached;
  return null; // signal: no text
}

function readCurrentSlide() {
  stopReading();
  const text = getSlideTextForTTS();
  if (text === null || text === '') {
    speak(`Slide ${State.current + 1} of ${State.total}. No readable text found on this slide. For PDF slides with extractable text, use a text-layer PDF.`);
  } else {
    speak(text);
  }
}

function speak(text) {
  if (!('speechSynthesis' in window)) {
    alert('Text-to-speech is not supported in your browser.');
    return;
  }
  stopReading();
  const utter = new SpeechSynthesisUtterance(text);
  utter.rate = 0.95;
  utter.pitch = 1;
  utter.volume = 1;
  utter.onstart = () => { State.ttsActive = true; setTTSStatus('speaking'); };
  utter.onend = () => { State.ttsActive = false; setTTSStatus('idle'); };
  utter.onerror = () => { State.ttsActive = false; setTTSStatus('idle'); };
  State.ttsUtterance = utter;
  window.speechSynthesis.speak(utter);
}

function stopReading() {
  if (window.speechSynthesis) window.speechSynthesis.cancel();
  State.ttsActive = false;
  setTTSStatus('idle');
}

function pauseReading() {
  if (window.speechSynthesis) window.speechSynthesis.pause();
  setTTSStatus('paused');
}

function resumeReading() {
  if (window.speechSynthesis) window.speechSynthesis.resume();
  setTTSStatus('speaking');
}

/* ─────────────────────────────────────────────
   11. PDF TEXT EXTRACTION (background)
   ───────────────────────────────────────────── */
async function extractAllPDFText(file) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  try {
    const ab = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const text = content.items.map(s => s.str).join(' ').trim();
      pdfTextCache[i - 1] = text;
    }
  } catch(e) {
    console.warn('PDF text extraction failed:', e);
  }
}

/* ─────────────────────────────────────────────
   12. HUGGING FACE AI
   ───────────────────────────────────────────── */
// The HF token is NEVER hardcoded. Users paste it in the UI.
// For Hugging Face Spaces: set HF_TOKEN as a Space Secret and
// use a minimal proxy. In pure static mode, users supply their own token.

const HF_MODEL = 'mistralai/Mistral-7B-Instruct-v0.3'; // free tier accessible

async function callHF(prompt) {
  const token = State.hfToken;

  const body = {
    inputs: prompt,
    parameters: {
      max_new_tokens: 300,
      temperature: 0.5,
      return_full_text: false,
    },
  };

  const headers = { 'Content-Type': 'application/json' };
  if (token) headers['Authorization'] = `Bearer ${token}`;

  const res = await fetch(`https://api-inference.huggingface.co/models/${HF_MODEL}`, {
    method: 'POST',
    headers,
    body: JSON.stringify(body),
  });

  if (res.status === 503) throw new Error('Model is loading. Please wait 20 seconds and try again.');
  if (res.status === 429) throw new Error('Rate limit reached. Please wait a moment or add your HF token.');
  if (res.status === 401) throw new Error('Invalid or missing HF token.');
  if (!res.ok) throw new Error(`API error ${res.status}`);

  const data = await res.json();
  if (Array.isArray(data) && data[0]?.generated_text) return data[0].generated_text.trim();
  if (data.generated_text) return data.generated_text.trim();
  if (data.error) throw new Error(data.error);
  throw new Error('Unexpected response format from API.');
}

async function triggerAI(mode, question = '') {
  if (!State.loaded) {
    showAIResult('Please load a presentation first.'); return;
  }

  const slideText = pdfTextCache[State.current] || '';
  if (!slideText) {
    showAIResult('No extractable text found on this slide. AI features work best with text-layer PDFs.'); return;
  }

  State.aiMode = mode;
  showAILoading(true);
  showAIResultPanel();

  let prompt = '';
  const context = `Slide ${State.current + 1} content:\n"${slideText.slice(0, 800)}"`;

  if (mode === 'summarize') {
    prompt = `[INST] Summarize the following slide content in 2-3 concise sentences:\n\n${context} [/INST]`;
  } else if (mode === 'explain') {
    prompt = `[INST] Explain the following slide content clearly for a student or general audience:\n\n${context} [/INST]`;
  } else if (mode === 'question') {
    prompt = `[INST] Based on this slide content, answer the following question:\n\n${context}\n\nQuestion: ${question} [/INST]`;
  }

  try {
    const answer = await callHF(prompt);
    showAIResult(answer);
  } catch(err) {
    showAIResult(`⚠ ${err.message}`);
  } finally {
    showAILoading(false);
  }
}

function showAIResultPanel() {
  DOM.aiResult.classList.remove('hidden');
  DOM.aiResultText.textContent = '';
}

function showAIResult(text) {
  DOM.aiResult.classList.remove('hidden');
  DOM.aiResultText.textContent = text;
}

function showAILoading(show) {
  DOM.aiLoading.classList.toggle('hidden', !show);
  DOM.aiResultText.classList.toggle('hidden', show);
}

/* ─────────────────────────────────────────────
   13. PRESENTATION (FULLSCREEN) MODE
   ───────────────────────────────────────────── */
function openPresentation() {
  if (!State.loaded) { alert('Load a presentation first.'); return; }
  State.presentationMode = true;
  DOM.presentationOverlay.classList.remove('hidden');
  renderFullscreen(State.current);
  // Try native fullscreen
  try { DOM.presentationOverlay.requestFullscreen(); } catch(e) {}
}

function closePresentation() {
  State.presentationMode = false;
  DOM.presentationOverlay.classList.add('hidden');
  if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
}

/* ─────────────────────────────────────────────
   14. SWIPE TOUCH SUPPORT
   ───────────────────────────────────────────── */
(function initSwipe() {
  let startX = 0;
  DOM.swipeOverlay.addEventListener('touchstart', e => {
    startX = e.touches[0].clientX;
  }, { passive: true });
  DOM.swipeOverlay.addEventListener('touchend', e => {
    const diff = e.changedTouches[0].clientX - startX;
    if (Math.abs(diff) > 50) {
      diff < 0 ? nextSlide() : prevSlide();
    }
  }, { passive: true });
})();

/* ─────────────────────────────────────────────
   15. STATUS INDICATORS
   ───────────────────────────────────────────── */
function setVoiceStatus(active) {
  DOM.voiceDot.className = 'status-dot' + (active ? ' listening' : '');
  DOM.voiceLabel.textContent = active ? 'Voice: listening' : 'Voice: off';
}
function setGestureStatus(active) {
  DOM.gestureDot.className = 'status-dot' + (active ? ' active' : '');
  DOM.gestureLabel.textContent = active ? 'Gesture: active' : 'Gesture: off';
}
function setTTSStatus(state) {
  const classes = { speaking: ' speaking', paused: ' listening', idle: '' };
  DOM.ttsDot.className = 'status-dot' + (classes[state] || '');
  DOM.ttsLabel.textContent = `TTS: ${state}`;
}
function setLastCommand(cmd) {
  DOM.lastCommandDisplay.textContent = `Last: ${cmd}`;
}

/* ─────────────────────────────────────────────
   16. HELPERS
   ───────────────────────────────────────────── */
function showProgress(pct, label) {
  DOM.uploadProgress.classList.remove('hidden');
  DOM.progressFill.style.width = `${pct}%`;
  DOM.progressLabel.textContent = label;
}

function showError(msg) {
  DOM.uploadProgress.classList.add('hidden');
  alert(`Error: ${msg}`);
}

function showGestureError(msg) {
  DOM.cameraError.textContent = msg;
  DOM.cameraError.classList.remove('hidden');
}

function escapeHTML(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

/* ─────────────────────────────────────────────
   17. EVENT WIRING
   ───────────────────────────────────────────── */
function wire() {
  // File upload
  DOM.fileInput.addEventListener('change', e => handleFiles(e.target.files));

  // Drag & drop
  DOM.dropZone.addEventListener('dragover', e => {
    e.preventDefault(); DOM.dropZone.classList.add('drag-over');
  });
  DOM.dropZone.addEventListener('dragleave', () => DOM.dropZone.classList.remove('drag-over'));
  DOM.dropZone.addEventListener('drop', e => {
    e.preventDefault();
    DOM.dropZone.classList.remove('drag-over');
    handleFiles(e.dataTransfer.files);
  });
  DOM.dropZone.addEventListener('click', e => {
    if (e.target === DOM.dropZone || e.target.closest('.drop-zone') && !e.target.closest('label')) {
      DOM.fileInput.click();
    }
  });

  // Navigation
  DOM.btnFirst.addEventListener('click', firstSlide);
  DOM.btnPrev.addEventListener('click', prevSlide);
  DOM.btnNext.addEventListener('click', nextSlide);
  DOM.btnLast.addEventListener('click', lastSlide);

  // Keyboard navigation
  document.addEventListener('keydown', e => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
    switch(e.key) {
      case 'ArrowRight': case 'ArrowDown': case 'PageDown': nextSlide(); break;
      case 'ArrowLeft': case 'ArrowUp': case 'PageUp': prevSlide(); break;
      case 'Home': firstSlide(); break;
      case 'End': lastSlide(); break;
      case 'Escape': closePresentation(); break;
      case 'F5': e.preventDefault(); openPresentation(); break;
    }
  });

  // Fullscreen / Presentation
  DOM.btnFullscreen.addEventListener('click', () => {
    const el = DOM.viewerContainer;
    if (!document.fullscreenElement) {
      el.requestFullscreen().catch(() => {});
    } else {
      document.exitFullscreen();
    }
  });
  DOM.btnStartPresent.addEventListener('click', openPresentation);
  DOM.fsBtnPrev.addEventListener('click', () => { prevSlide(); });
  DOM.fsBtnNext.addEventListener('click', () => { nextSlide(); });
  DOM.fsBtnExit.addEventListener('click', closePresentation);
  document.addEventListener('fullscreenchange', () => {
    if (!document.fullscreenElement && State.presentationMode) closePresentation();
  });

  // Camera
  DOM.btnStartCamera.addEventListener('click', startCamera);
  DOM.btnStopCamera.addEventListener('click', stopCamera);

  // Voice
  DOM.btnStartVoice.addEventListener('click', startVoice);
  DOM.btnStopVoice.addEventListener('click', stopVoice);

  // TTS
  DOM.btnReadSlide.addEventListener('click', readCurrentSlide);
  DOM.btnPauseRead.addEventListener('click', pauseReading);
  DOM.btnResumeRead.addEventListener('click', resumeReading);
  DOM.btnStopRead.addEventListener('click', stopReading);
  DOM.btnRepeatRead.addEventListener('click', readCurrentSlide);

  // AI
  DOM.btnSaveToken.addEventListener('click', () => {
    State.hfToken = DOM.hfToken.value.trim();
    DOM.btnSaveToken.textContent = State.hfToken ? '✓ Saved' : 'Save';
    setTimeout(() => { DOM.btnSaveToken.textContent = 'Save'; }, 2000);
  });
  DOM.btnSummarize.addEventListener('click', () => triggerAI('summarize'));
  DOM.btnExplain.addEventListener('click', () => triggerAI('explain'));
  DOM.btnQuestion.addEventListener('click', () => {
    DOM.aiQuestionBox.classList.toggle('hidden');
    if (!DOM.aiQuestionBox.classList.contains('hidden')) DOM.aiQuestionInput.focus();
  });
  DOM.btnAskQuestion.addEventListener('click', () => {
    const q = DOM.aiQuestionInput.value.trim();
    if (q) triggerAI('question', q);
  });
  DOM.aiQuestionInput.addEventListener('keydown', e => {
    if (e.key === 'Enter') DOM.btnAskQuestion.click();
  });
  DOM.btnCloseAI.addEventListener('click', () => DOM.aiResult.classList.add('hidden'));

  // Help modal
  DOM.btnHelp.addEventListener('click', () => DOM.helpModal.classList.remove('hidden'));
  DOM.btnCloseHelp.addEventListener('click', () => DOM.helpModal.classList.add('hidden'));
  DOM.helpModal.addEventListener('click', e => {
    if (e.target === DOM.helpModal) DOM.helpModal.classList.add('hidden');
  });

  // Special handling for files with PDF text
  const origHandleFiles = window.handleFiles;
  const _origHandleFiles = handleFiles;

  // Intercept PDF to also extract text
  DOM.fileInput.removeEventListener('change', DOM.fileInput._listener);
  DOM.fileInput.addEventListener('change', async e => {
    const files = e.target.files;
    // Extract PDF text in background
    const pdfFile = Array.from(files).find(f => f.type === 'application/pdf');
    if (pdfFile) extractAllPDFText(pdfFile);
    handleFiles(files);
  });
}

/* ─────────────────────────────────────────────
   18. INIT
   ───────────────────────────────────────────── */
(function init() {
  wire();

  // Check speech recognition support
  if (!window.SpeechRecognition && !window.webkitSpeechRecognition) {
    DOM.voiceError.textContent = 'Voice recognition not supported. Use Chrome or Edge for this feature.';
    DOM.voiceError.classList.remove('hidden');
  }

  // Check TTS support
  if (!('speechSynthesis' in window)) {
    DOM.ttsLabel.textContent = 'TTS: unsupported';
  }

  // Restore HF token from sessionStorage (not localStorage — more secure)
  const saved = sessionStorage.getItem('hf_token');
  if (saved) {
    State.hfToken = saved;
    DOM.hfToken.value = saved;
  }
  DOM.btnSaveToken.addEventListener('click', () => {
    if (State.hfToken) sessionStorage.setItem('hf_token', State.hfToken);
  });

  console.log('SmartSlides ready. Upload a presentation to begin.');
})();
