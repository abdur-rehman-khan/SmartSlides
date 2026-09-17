# ◈ SmartSlides

> A gesture + voice controlled presentation viewer that runs entirely in the browser — no backend required.

---

## Features

| Feature | How it works |
|---|---|
| **PDF & Image upload** | PDF.js renders each page to canvas; images load natively |
| **Hand gesture control** | MediaPipe Tasks Vision detects your wrist position via webcam |
| **Voice commands** | Web Speech API (SpeechRecognition) listens continuously |
| **Read aloud** | SpeechSynthesis API reads slide text with pause/resume/stop |
| **AI: Summarize / Explain / Q&A** | Hugging Face Inference API with Mistral-7B |
| **Slide thumbnails** | Auto-generated strip for quick navigation |
| **Fullscreen mode** | Native browser fullscreen with on-hover controls |
| **Touch swipe** | Left/right swipe on mobile |
| **Keyboard shortcuts** | Arrow keys, Page Up/Down, Home, End |

---

## Gesture Guide

| Movement | Action |
|---|---|
| Move open hand **→ Right** | Previous slide |
| Move open hand **← Left** | Next slide |

> The video is mirrored, so move your hand in the direction you see it move on screen.

---

## Voice Commands

| Say | Action |
|---|---|
| `next slide` | Go forward one slide |
| `previous slide` | Go back one slide |
| `first slide` | Jump to slide 1 |
| `last slide` | Jump to the last slide |
| `read this slide` | Read slide text aloud |
| `stop reading` | Stop TTS |
| `pause` | Pause TTS |
| `resume` | Resume TTS |
| `summarize slide` | AI summary |
| `explain slide` | AI explanation |
| `start presentation` | Open fullscreen mode |
| `stop presentation` | Exit fullscreen mode |

---

## AI Setup (Optional)

AI features (Summarize, Explain, Ask a Question) use the **Hugging Face Inference API** with [`mistralai/Mistral-7B-Instruct-v0.3`](https://huggingface.co/mistralai/Mistral-7B-Instruct-v0.3), which is **free and publicly accessible**.

**To use AI features:**
1. Get a free HF token at [huggingface.co/settings/tokens](https://huggingface.co/settings/tokens) (free account required)
2. Paste the token into the **AI Assistant** panel in the app
3. Click **Save**

The token is stored only in `sessionStorage` (cleared when you close the tab). It is **never** embedded in the source code.

> **Note:** AI features only work on **text-layer PDFs** (PDFs with selectable text). Scanned image PDFs will show "No extractable text found."

### Hugging Face Spaces: Secure Token Handling

When deploying to HF Spaces, you can expose the token to users via a **Space Secret**. The correct approach for a Static Space is:

1. In your Space settings → **Variables and Secrets** → add `HF_TOKEN` as a **secret**
2. Static Spaces do **not** serve server-side code, so the token cannot be automatically injected. Users must paste their own token.
3. If you want automatic token injection, upgrade to a **Docker Space** with a minimal reverse proxy (see `proxy/` folder in advanced setup).

For a pure static deployment, the current approach (user-supplied token) is the correct and secure one.

---

## File Structure

```
smart-slides/
├── index.html      ← Main app (single-page)
├── style.css       ← All styles
├── script.js       ← All logic
└── README.md       ← This file
```

---

## Deploying to Hugging Face Spaces (Static)

### Step-by-step

**1. Create a new Space**
- Go to [huggingface.co/new-space](https://huggingface.co/new-space)
- Give it a name, e.g. `smart-slides`
- Set **Space SDK** → **Static**
- Set visibility to **Public** (or Private)
- Click **Create Space**

**2. Upload your files**

Option A — via the web UI:
- In your new Space, click **Files** → **Add file** → **Upload files**
- Upload `index.html`, `style.css`, `script.js`, `README.md`
- Click **Commit changes**

Option B — via Git:
```bash
git clone https://huggingface.co/spaces/YOUR_USERNAME/smart-slides
cd smart-slides
# copy index.html style.css script.js README.md here
git add .
git commit -m "Initial upload"
git push
```

**3. Done!**
- Your app is live at `https://YOUR_USERNAME-smart-slides.hf.space`
- No build step, no server, no costs

---

## Browser Compatibility

| Feature | Chrome | Edge | Firefox | Safari |
|---|---|---|---|---|
| Voice Recognition | ✅ | ✅ | ❌ | ❌ |
| TTS / Read Aloud | ✅ | ✅ | ✅ | ✅ |
| Hand Gestures | ✅ | ✅ | ✅* | ✅* |
| PDF Upload | ✅ | ✅ | ✅ | ✅ |

> *Gesture detection uses WebGL/WASM — works in most modern browsers but GPU delegate may fall back to CPU.

**Recommended browser: Google Chrome or Microsoft Edge** for full feature support including voice recognition.

---

## Keyboard Shortcuts

| Key | Action |
|---|---|
| `→` / `↓` / `Page Down` | Next slide |
| `←` / `↑` / `Page Up` | Previous slide |
| `Home` | First slide |
| `End` | Last slide |
| `F5` | Start presentation (fullscreen) |
| `Esc` | Exit fullscreen |

---

## Privacy

- **Webcam feed** stays on your device — frames are processed locally by MediaPipe WASM
- **Voice audio** is processed by your browser's built-in speech recognition (on-device or browser vendor's server)
- **HF token** is stored in `sessionStorage` only — cleared when you close the tab
- **Slide content** never leaves your browser unless you use AI features

---

## Limitations & Known Issues

- Voice recognition requires Chrome or Edge (Web Speech API)
- AI features require a PDF with a text layer (not scanned images)
- Hand gesture detection requires good lighting and a clear background
- MediaPipe model (~25MB) loads on first camera start — may take 5-10 seconds on slow connections
- The Hugging Face free tier has rate limits; if you hit them, wait 60 seconds

---

## License

MIT — free to use, modify, and deploy.
