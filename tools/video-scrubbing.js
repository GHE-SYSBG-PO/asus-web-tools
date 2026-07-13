(function () {
  'use strict';

  // ─── MiniZip ─────────────────────────────────────────────────────────────────
  const CRC_TABLE = (function () {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[i] = c;
    }
    return t;
  })();

  function crc32(buf) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  async function deflateRaw(data) {
    const cs = new CompressionStream('deflate-raw');
    const w = cs.writable.getWriter();
    w.write(data); w.close();
    return new Uint8Array(await new Response(cs.readable).arrayBuffer());
  }

  const le16 = n => { const b = new Uint8Array(2); new DataView(b.buffer).setUint16(0, n, true); return b; };
  const le32 = n => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n, true); return b; };
  function concat(...arrs) {
    const out = new Uint8Array(arrs.reduce((s, a) => s + a.length, 0));
    let off = 0; for (const a of arrs) { out.set(a, off); off += a.length; }
    return out;
  }

  class MiniZip {
    constructor() { this._e = []; this._off = 0; }
    async add(name, data) {
      const n = new TextEncoder().encode(name);
      const crc = crc32(data);
      let body, method;
      try { const d = await deflateRaw(data); body = d.length < data.length ? d : data; method = d.length < data.length ? 8 : 0; }
      catch (_) { body = data; method = 0; }
      const lh = concat(
        new Uint8Array([0x50,0x4B,0x03,0x04]), le16(20), le16(0), le16(method),
        le16(0), le16(0), le32(crc), le32(body.length), le32(data.length),
        le16(n.length), le16(0), n
      );
      this._e.push({ n, off: this._off, crc, method, cs: body.length, us: data.length, lh, body });
      this._off += lh.length + body.length;
    }
    generate() {
      const parts = [];
      for (const e of this._e) parts.push(e.lh, e.body);
      const cdOff = this._off; let cdLen = 0;
      for (const e of this._e) {
        const cd = concat(
          new Uint8Array([0x50,0x4B,0x01,0x02]), le16(20), le16(20), le16(0), le16(e.method),
          le16(0), le16(0), le32(e.crc), le32(e.cs), le32(e.us),
          le16(e.n.length), le16(0), le16(0), le16(0), le16(0), le32(0), le32(e.off), e.n
        );
        parts.push(cd); cdLen += cd.length;
      }
      parts.push(concat(
        new Uint8Array([0x50,0x4B,0x05,0x06]), le16(0), le16(0),
        le16(this._e.length), le16(this._e.length), le32(cdLen), le32(cdOff), le16(0)
      ));
      return concat(...parts);
    }
  }

  // ─── CSS ─────────────────────────────────────────────────────────────────────
  const css = `
    #video-scrubbing-app {
      --vs-card: rgba(25, 25, 38, 0.4);
      --vs-glass: rgba(255, 255, 255, 0.04);
      --vs-border: rgba(255, 255, 255, 0.08);
      --vs-border-a: rgba(139, 92, 246, 0.5);
      --vs-text: #f0f0f5;
      --vs-text-2: #8888a0;
      --vs-text-3: #55556a;
      --vs-accent: #8b5cf6;
      --vs-glow: rgba(139, 92, 246, 0.25);
      --vs-ok: #34d399;
      --vs-warn: #f59e0b;
      --vs-r-sm: 8px; --vs-r-md: 12px; --vs-r-lg: 16px;
      --vs-font: 'Inter', sans-serif;
      --vs-mono: 'JetBrains Mono', monospace;
      --vs-ease: 0.2s cubic-bezier(0.4, 0, 0.2, 1);

      font-family: var(--vs-font);
      color: var(--vs-text);
      height: 100%;
      overflow-y: auto;
      padding: 32px 20px;
      display: flex;
      justify-content: center;
    }

    #video-scrubbing-app .vs-app {
      width: 100%; max-width: 900px;
      display: flex; flex-direction: column; gap: 24px;
    }

    #video-scrubbing-app .vs-header { text-align: center; }
    #video-scrubbing-app .vs-title {
      font-size: 1.8rem; font-weight: 700;
      display: flex; align-items: center; justify-content: center; gap: 10px;
    }
    #video-scrubbing-app .vs-title-icon {
      background: linear-gradient(135deg, #8b5cf6, #ec4899, #f97316);
      -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    #video-scrubbing-app .vs-subtitle { color: var(--vs-text-2); font-size: 0.9rem; margin-top: 6px; }

    #video-scrubbing-app .vs-upload {
      border: 2px dashed var(--vs-border);
      border-radius: var(--vs-r-lg);
      padding: 40px 20px; text-align: center; cursor: pointer;
      transition: all var(--vs-ease);
      background: var(--vs-card); backdrop-filter: blur(12px);
    }
    #video-scrubbing-app .vs-upload:hover,
    #video-scrubbing-app .vs-upload.drag-over {
      border-color: var(--vs-accent); background: var(--vs-glow);
    }
    #video-scrubbing-app .vs-upload__icon { font-size: 2.5rem; margin-bottom: 12px; }
    #video-scrubbing-app .vs-upload__text { color: var(--vs-text-2); font-size: 0.9rem; }
    #video-scrubbing-app .vs-upload__hint { color: var(--vs-text-3); font-size: 0.8rem; margin-top: 6px; }
    #video-scrubbing-app .vs-upload input[type="file"] { display: none; }

    #video-scrubbing-app .vs-info-bar {
      background: var(--vs-glass); border: 1px solid var(--vs-border);
      border-radius: var(--vs-r-md); padding: 12px 16px;
      display: flex; gap: 20px; flex-wrap: wrap;
      font-size: 0.8rem; color: var(--vs-text-2);
    }
    #video-scrubbing-app .vs-info-bar strong { color: var(--vs-text); }

    #video-scrubbing-app .vs-grid {
      display: grid; grid-template-columns: 260px 1fr;
      gap: 20px; align-items: start;
    }
    @media (max-width: 680px) { #video-scrubbing-app .vs-grid { grid-template-columns: 1fr; } }

    #video-scrubbing-app .vs-panel {
      background: var(--vs-card); border: 1px solid var(--vs-border);
      border-radius: var(--vs-r-lg); backdrop-filter: blur(12px);
    }
    #video-scrubbing-app .vs-panel--pad { padding: 20px; display: flex; flex-direction: column; gap: 18px; }
    #video-scrubbing-app .vs-panel-head {
      padding: 12px 16px; border-bottom: 1px solid var(--vs-border);
      display: flex; align-items: center; justify-content: space-between;
      font-size: 0.8rem; color: var(--vs-text-2); font-weight: 600;
    }
    #video-scrubbing-app .vs-sec-label {
      font-size: 0.75rem; font-weight: 600; text-transform: uppercase;
      letter-spacing: 0.08em; color: var(--vs-text-2);
    }

    #video-scrubbing-app .vs-field { display: flex; flex-direction: column; gap: 6px; }
    #video-scrubbing-app .vs-field-row {
      display: flex; justify-content: space-between; align-items: center;
      font-size: 0.8rem; color: var(--vs-text-2);
    }
    #video-scrubbing-app .vs-field-val {
      font-family: var(--vs-mono); font-size: 0.75rem; color: var(--vs-accent);
    }
    #video-scrubbing-app .vs-range {
      width: 100%; -webkit-appearance: none;
      height: 4px; border-radius: 2px; outline: none; cursor: pointer;
      background: linear-gradient(to right, var(--vs-accent) var(--pct,0%), var(--vs-border) var(--pct,0%));
    }
    #video-scrubbing-app .vs-range::-webkit-slider-thumb {
      -webkit-appearance: none; width: 16px; height: 16px;
      border-radius: 50%; background: var(--vs-accent);
      border: 2px solid #1a1a2e; cursor: pointer;
      box-shadow: 0 0 0 3px var(--vs-glow);
    }
    #video-scrubbing-app .vs-toggle-row {
      display: flex; align-items: center; justify-content: space-between;
      font-size: 0.8rem; color: var(--vs-text-2);
    }
    #video-scrubbing-app .vs-toggle { position: relative; width: 36px; height: 20px; flex-shrink: 0; }
    #video-scrubbing-app .vs-toggle input { opacity: 0; width: 0; height: 0; position: absolute; }
    #video-scrubbing-app .vs-toggle-track {
      position: absolute; inset: 0; background: var(--vs-border);
      border-radius: 10px; cursor: pointer; transition: background var(--vs-ease);
    }
    #video-scrubbing-app .vs-toggle-thumb {
      position: absolute; top: 3px; left: 3px;
      width: 14px; height: 14px; border-radius: 50%;
      background: var(--vs-text-2); transition: all var(--vs-ease); pointer-events: none;
    }
    #video-scrubbing-app .vs-toggle input:checked ~ .vs-toggle-track { background: var(--vs-accent); }
    #video-scrubbing-app .vs-toggle input:checked ~ .vs-toggle-thumb { left: 19px; background: #fff; }

    #video-scrubbing-app .vs-badge {
      font-size: 0.7rem; font-weight: 500;
      background: var(--vs-glow); color: var(--vs-accent);
      border: 1px solid var(--vs-border-a);
      border-radius: 4px; padding: 2px 8px;
      animation: vs-blink 2.4s ease-in-out infinite;
    }
    @keyframes vs-blink { 0%,100%{opacity:1} 50%{opacity:0.5} }

    /* Preview mode tabs */
    #video-scrubbing-app .vs-prev-tabs { display: flex; gap: 2px; }
    #video-scrubbing-app .vs-prev-tab {
      padding: 3px 9px; border: 1px solid var(--vs-border);
      border-radius: var(--vs-r-sm); background: var(--vs-glass);
      color: var(--vs-text-3); font-size: 0.7rem; cursor: pointer;
      transition: all var(--vs-ease); font-family: var(--vs-font);
    }
    #video-scrubbing-app .vs-prev-tab:hover { color: var(--vs-text-2); }
    #video-scrubbing-app .vs-prev-tab--active {
      color: var(--vs-accent); border-color: var(--vs-border-a); background: var(--vs-glow);
    }
    #video-scrubbing-app .vs-prev-tab:disabled {
      opacity: 0.35; cursor: not-allowed; pointer-events: none;
    }

    /* DOM preview frames */
    #video-scrubbing-app .vs-dom-frames {
      position: absolute; inset: 0;
      display: flex; align-items: center; justify-content: center;
      background: #000;
    }
    #video-scrubbing-app .vs-dom-frame {
      position: absolute;
      max-width: 100%; max-height: 320px;
      opacity: 0; pointer-events: none;
      display: block;
    }
    #video-scrubbing-app .vs-dom-frame.active { opacity: 1; }

    #video-scrubbing-app .vs-imgseq-loading {
      position: absolute; inset: 0; background: rgba(13,13,18,0.85);
      display: flex; align-items: center; justify-content: center; z-index: 2;
      flex-direction: column; gap: 10px;
      font-size: 0.8rem; color: var(--vs-text-2);
    }
    #video-scrubbing-app .vs-imgseq-loading .vs-load-bar {
      width: 140px; height: 3px; background: var(--vs-border); border-radius: 2px; overflow: hidden;
    }
    #video-scrubbing-app .vs-imgseq-loading .vs-load-fill {
      height: 100%; background: var(--vs-accent); width: 0%; transition: width 0.1s;
    }

    #video-scrubbing-app .vs-placeholder {
      height: 320px; display: flex; flex-direction: column;
      align-items: center; justify-content: center;
      color: var(--vs-text-3); gap: 10px; font-size: 0.85rem;
    }
    #video-scrubbing-app .vs-placeholder-icon { font-size: 3rem; opacity: 0.35; }

    #video-scrubbing-app .vs-scroll-box {
      height: 320px; overflow-y: scroll; position: relative;
      scrollbar-width: thin; scrollbar-color: var(--vs-border) transparent;
    }
    #video-scrubbing-app .vs-scroll-box::-webkit-scrollbar { width: 4px; }
    #video-scrubbing-app .vs-scroll-box::-webkit-scrollbar-thumb { background: var(--vs-border); border-radius: 2px; }
    #video-scrubbing-app .vs-spacer { position: relative; }
    #video-scrubbing-app .vs-canvas-wrap {
      position: sticky; top: 0; height: 320px;
      display: flex; align-items: center; justify-content: center; background: #000;
    }
    #video-scrubbing-app .vs-canvas-wrap canvas { display: block; }

    #video-scrubbing-app .vs-progress { height: 3px; background: var(--vs-border); }
    #video-scrubbing-app .vs-progress-fill {
      height: 100%; width: 0%;
      background: linear-gradient(90deg, #8b5cf6, #ec4899); transition: width 0.05s linear;
    }
    #video-scrubbing-app .vs-time-row {
      padding: 8px 16px; border-top: 1px solid var(--vs-border);
      display: flex; justify-content: space-between;
      font-family: var(--vs-mono); font-size: 0.72rem; color: var(--vs-text-3);
    }

    /* Extract Panel */
    #video-scrubbing-app .vs-ext-body {
      padding: 20px; display: flex; flex-direction: column; gap: 16px;
    }
    #video-scrubbing-app .vs-ext-grid {
      display: grid; grid-template-columns: 1fr 1fr; gap: 16px;
    }
    @media (max-width: 600px) { #video-scrubbing-app .vs-ext-grid { grid-template-columns: 1fr; } }

    #video-scrubbing-app .vs-fmt-row {
      display: flex; align-items: center; gap: 12px;
      font-size: 0.8rem; color: var(--vs-text-2);
    }
    #video-scrubbing-app .vs-fmt-btns { display: flex; gap: 4px; }
    #video-scrubbing-app .vs-fmt-btn {
      padding: 4px 12px; border: 1px solid var(--vs-border);
      border-radius: var(--vs-r-sm); background: var(--vs-glass);
      color: var(--vs-text-2); font-size: 0.78rem; cursor: pointer;
      transition: all var(--vs-ease); font-family: var(--vs-font);
    }
    #video-scrubbing-app .vs-fmt-btn:hover { color: var(--vs-accent); border-color: var(--vs-accent); }
    #video-scrubbing-app .vs-fmt-btn--active {
      color: var(--vs-accent); border-color: var(--vs-accent); background: var(--vs-glow);
    }

    #video-scrubbing-app .vs-ext-footer {
      display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
    }
    #video-scrubbing-app .vs-estimate {
      font-size: 0.78rem; color: var(--vs-text-3); font-family: var(--vs-mono);
    }
    #video-scrubbing-app .vs-estimate strong { color: var(--vs-text-2); }

    #video-scrubbing-app .vs-ext-progress {
      flex: 1; display: flex; flex-direction: column; gap: 4px; min-width: 120px;
    }
    #video-scrubbing-app .vs-ext-bar {
      height: 6px; background: var(--vs-border); border-radius: 3px; overflow: hidden;
    }
    #video-scrubbing-app .vs-ext-bar-fill {
      height: 100%; width: 0%; background: linear-gradient(90deg, #8b5cf6, #34d399);
      transition: width 0.1s linear; border-radius: 3px;
    }
    #video-scrubbing-app .vs-ext-status {
      font-size: 0.72rem; color: var(--vs-text-3); font-family: var(--vs-mono);
    }

    /* Code Output */
    #video-scrubbing-app .vs-code-tabs {
      display: flex; gap: 2px;
    }
    #video-scrubbing-app .vs-code-tab {
      padding: 5px 12px; border: 1px solid transparent;
      border-radius: var(--vs-r-sm); background: transparent;
      color: var(--vs-text-3); font-size: 0.75rem; cursor: pointer;
      transition: all var(--vs-ease); font-family: var(--vs-font);
    }
    #video-scrubbing-app .vs-code-tab:hover { color: var(--vs-text-2); }
    #video-scrubbing-app .vs-code-tab--active {
      color: var(--vs-accent); border-color: var(--vs-border-a);
      background: var(--vs-glow);
    }
    #video-scrubbing-app .vs-code-body { overflow: auto; max-height: 320px; }
    #video-scrubbing-app .vs-code-body pre {
      margin: 0; padding: 16px;
      font-family: var(--vs-mono); font-size: 0.78rem;
      color: var(--vs-text); line-height: 1.65; white-space: pre;
    }

    /* Buttons */
    #video-scrubbing-app .vs-btn {
      display: inline-flex; align-items: center; gap: 6px;
      padding: 7px 14px; border: 1px solid var(--vs-border);
      border-radius: var(--vs-r-sm); background: var(--vs-glass);
      color: var(--vs-text-2); font-size: 0.78rem; cursor: pointer;
      transition: all var(--vs-ease); font-family: var(--vs-font); white-space: nowrap;
    }
    #video-scrubbing-app .vs-btn:hover { color: var(--vs-accent); border-color: var(--vs-accent); background: var(--vs-glow); }
    #video-scrubbing-app .vs-btn:disabled { opacity: 0.4; cursor: not-allowed; pointer-events: none; }
    #video-scrubbing-app .vs-btn--primary {
      background: var(--vs-glow); border-color: var(--vs-border-a); color: var(--vs-accent);
    }
    #video-scrubbing-app .vs-btn--primary:hover { background: rgba(139,92,246,0.35); }
    #video-scrubbing-app .vs-btn--cancel { color: #ef4444; border-color: rgba(239,68,68,0.4); background: rgba(239,68,68,0.08); }
    #video-scrubbing-app .vs-btn--cancel:hover { background: rgba(239,68,68,0.15); }
    #video-scrubbing-app .vs-btn--dl {
      color: var(--vs-ok); border-color: rgba(52,211,153,0.4); background: rgba(52,211,153,0.08);
    }
    #video-scrubbing-app .vs-btn--dl:hover { background: rgba(52,211,153,0.15); }
    #video-scrubbing-app .vs-btn--ok { color: var(--vs-ok); border-color: rgba(52,211,153,0.4); background: rgba(52,211,153,0.1); }
    #video-scrubbing-app .vs-btn--ok:hover { background: rgba(52,211,153,0.15); }

    #video-scrubbing-app .vs-hidden { display: none !important; }
  `;

  // ─── HTML ─────────────────────────────────────────────────────────────────────
  const html = `
    <div class="vs-app">
      <header class="vs-header">
        <h1 class="vs-title">
          <span class="vs-title-icon">▶</span>
          Video Scrubbing
        </h1>
        <p class="vs-subtitle">上傳影片，滾動逐格播放 · 支援拆幀匯出為 Image Sequence</p>
      </header>

      <div class="vs-upload" id="vs-upload" role="button" tabindex="0">
        <input type="file" id="vs-file-input" accept="video/*">
        <div class="vs-upload__icon">🎬</div>
        <div class="vs-upload__text">點擊或拖曳影片至此處</div>
        <div class="vs-upload__hint">支援 MP4、WebM、MOV 等格式</div>
      </div>

      <div class="vs-info-bar vs-hidden" id="vs-info-bar">
        <span>📄 <strong id="vs-fname">—</strong></span>
        <span>時長 <strong id="vs-dur">—</strong></span>
        <span>尺寸 <strong id="vs-dim">—</strong></span>
        <span>估算幀數 <strong id="vs-frames">—</strong></span>
      </div>

      <div class="vs-grid">
        <div class="vs-panel vs-panel--pad">
          <div class="vs-sec-label">滾動設定</div>

          <div class="vs-field">
            <div class="vs-field-row">滾動高度 <span class="vs-field-val" id="vs-sh-lbl">3000px</span></div>
            <input class="vs-range" type="range" id="vs-sh" min="500" max="15000" step="500" value="3000">
          </div>
          <div class="vs-field">
            <div class="vs-field-row">開始時間 <span class="vs-field-val" id="vs-st-lbl">0%</span></div>
            <input class="vs-range" type="range" id="vs-st" min="0" max="100" step="1" value="0">
          </div>
          <div class="vs-field">
            <div class="vs-field-row">結束時間 <span class="vs-field-val" id="vs-et-lbl">100%</span></div>
            <input class="vs-range" type="range" id="vs-et" min="0" max="100" step="1" value="100">
          </div>
          <div class="vs-toggle-row">
            反向播放
            <label class="vs-toggle">
              <input type="checkbox" id="vs-rev">
              <span class="vs-toggle-track"></span>
              <span class="vs-toggle-thumb"></span>
            </label>
          </div>
        </div>

        <div class="vs-panel">
          <div class="vs-panel-head">
            預覽
            <div style="display:flex;align-items:center;gap:8px">
              <div class="vs-prev-tabs vs-hidden" id="vs-prev-tabs">
                <button class="vs-prev-tab vs-prev-tab--active" id="vs-prev-canvas">Canvas</button>
                <button class="vs-prev-tab" id="vs-prev-imgseq" disabled>Img Seq</button>
                <button class="vs-prev-tab" id="vs-prev-dom" disabled>DOM</button>
              </div>
              <span class="vs-badge">↓ 在此滾動</span>
            </div>
          </div>
          <div id="vs-placeholder" class="vs-placeholder">
            <div class="vs-placeholder-icon">🎞️</div>
            上傳影片後即可預覽滾動效果
          </div>
          <div class="vs-scroll-box vs-hidden" id="vs-scroll-box">
            <div id="vs-spacer" class="vs-spacer" style="height:3000px">
              <div class="vs-canvas-wrap">
                <canvas id="vs-canvas"></canvas>
              </div>
            </div>
          </div>
          <div class="vs-progress vs-hidden" id="vs-progress">
            <div class="vs-progress-fill" id="vs-pfill"></div>
          </div>
          <div class="vs-time-row vs-hidden" id="vs-time-row">
            <span id="vs-cur">0.000s</span>
            <span id="vs-tot">0.000s</span>
          </div>
        </div>
      </div>

      <!-- 拆幀匯出 -->
      <div class="vs-panel" id="vs-ext-panel">
        <div class="vs-panel-head">
          拆幀匯出
          <span class="vs-estimate" id="vs-estimate">上傳影片後可拆幀</span>
        </div>
        <div class="vs-ext-body">
          <div class="vs-ext-grid">
            <div class="vs-field">
              <div class="vs-field-row">FPS <span class="vs-field-val" id="vs-fps-lbl">24fps</span></div>
              <input class="vs-range" type="range" id="vs-fps" min="1" max="60" step="1" value="24">
            </div>
            <div class="vs-field">
              <div class="vs-field-row">品質 <span class="vs-field-val" id="vs-qual-lbl">80%</span></div>
              <input class="vs-range" type="range" id="vs-qual" min="30" max="100" step="5" value="80">
            </div>
          </div>
          <div class="vs-fmt-row">
            <span>格式</span>
            <div class="vs-fmt-btns">
              <button class="vs-fmt-btn vs-fmt-btn--active" id="vs-fmt-jpg">JPEG</button>
              <button class="vs-fmt-btn" id="vs-fmt-webp">WebP</button>
            </div>
          </div>
          <div class="vs-ext-footer">
            <button class="vs-btn vs-btn--primary" id="vs-ext-btn" disabled>開始拆幀</button>
            <div class="vs-ext-progress vs-hidden" id="vs-ext-prog">
              <div class="vs-ext-bar"><div class="vs-ext-bar-fill" id="vs-ext-fill"></div></div>
              <span class="vs-ext-status" id="vs-ext-status">0 / 0</span>
            </div>
            <button class="vs-btn vs-btn--dl vs-hidden" id="vs-dl-btn">↓ 下載 ZIP</button>
          </div>
        </div>
      </div>

      <!-- Code Output -->
      <div class="vs-panel">
        <div class="vs-panel-head">
          <div class="vs-code-tabs">
            <button class="vs-code-tab vs-code-tab--active" data-tab="canvas">Canvas 版</button>
            <button class="vs-code-tab" data-tab="imgseq">Img Seq 版</button>
            <button class="vs-code-tab" data-tab="dom">DOM 版</button>
          </div>
          <button class="vs-btn" id="vs-copy-btn">
            <svg width="13" height="13" viewBox="0 0 18 18" fill="none">
              <rect x="6" y="6" width="10" height="10" rx="2" stroke="currentColor" stroke-width="1.5"/>
              <path d="M12 6V4a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h2" stroke="currentColor" stroke-width="1.5"/>
            </svg>
            複製代碼
          </button>
        </div>
        <div class="vs-code-body">
          <pre id="vs-code">// 上傳影片後將顯示對應代碼</pre>
        </div>
      </div>
    </div>
  `;

  const app = document.getElementById('video-scrubbing-app');
  if (!app) return;
  const styleEl = document.createElement('style');
  styleEl.textContent = css;
  app.appendChild(styleEl);
  app.insertAdjacentHTML('beforeend', html);

  // ─── Helpers ──────────────────────────────────────────────────────────────────
  const get = id => document.getElementById(id);

  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text);
    const ta = document.createElement('textarea');
    ta.value = text; ta.style.cssText = 'position:fixed;top:-9999px;left:-9999px';
    document.body.appendChild(ta); ta.focus(); ta.select();
    document.execCommand('copy'); document.body.removeChild(ta);
    return Promise.resolve();
  }

  function syncRange(id, lblId, fmt) {
    const el = get(id);
    get(lblId).textContent = fmt(el.value);
    el.style.setProperty('--pct', ((el.value - el.min) / (el.max - el.min) * 100) + '%');
  }

  // ─── Video element (preview) ──────────────────────────────────────────────────
  const vid = document.createElement('video');
  vid.preload = 'auto'; vid.muted = true; vid.playsInline = true;

  let ctx = null;
  let currentFile = null;

  const S = {
    loaded: false, duration: 0, vw: 0, vh: 0,
    sh: 3000, startPct: 0, endPct: 1, reverse: false,
    targetTime: 0, seeking: false,
    // extract
    extFps: 24, extFormat: 'image/jpeg', extQuality: 0.8,
    extracting: false, extractedCount: 0, totalFrames: 0,
    zipData: null, extractedFrameCount: 0,
    // preview mode
    previewMode: 'canvas',   // 'canvas' | 'imgseq'
    previewImgs: null,        // HTMLImageElement[]
    previewBlobUrls: [],      // for cleanup
    // code tab
    codeTab: 'canvas',
  };

  // ─── Upload ───────────────────────────────────────────────────────────────────
  const uploadEl = get('vs-upload');
  const fileInput = get('vs-file-input');

  uploadEl.addEventListener('click', () => fileInput.click());
  uploadEl.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') fileInput.click(); });
  uploadEl.addEventListener('dragover', e => { e.preventDefault(); uploadEl.classList.add('drag-over'); });
  uploadEl.addEventListener('dragleave', e => { if (!uploadEl.contains(e.relatedTarget)) uploadEl.classList.remove('drag-over'); });
  uploadEl.addEventListener('drop', e => {
    e.preventDefault(); uploadEl.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f && f.type.startsWith('video/')) loadFile(f);
  });
  fileInput.addEventListener('change', e => { if (e.target.files[0]) loadFile(e.target.files[0]); e.target.value = ''; });

  function loadFile(file) {
    currentFile = file;
    if (vid.src) URL.revokeObjectURL(vid.src);
    vid.src = URL.createObjectURL(file);
    vid.load();
    // reset extract state
    S.zipData = null; S.extractedFrameCount = 0;
    get('vs-dl-btn').classList.add('vs-hidden');
    get('vs-ext-prog').classList.add('vs-hidden');
    // reset preview mode
    S.previewBlobUrls.forEach(u => URL.revokeObjectURL(u));
    S.previewBlobUrls = []; S.previewImgs = null; S.previewMode = 'canvas';
    get('vs-prev-tabs').classList.add('vs-hidden');
    get('vs-prev-canvas').classList.add('vs-prev-tab--active');
    get('vs-prev-imgseq').classList.remove('vs-prev-tab--active');
    get('vs-prev-dom').classList.remove('vs-prev-tab--active');
    get('vs-prev-imgseq').disabled = true;
    get('vs-prev-dom').disabled = true;
    const oldDom = document.getElementById('vs-dom-frames');
    if (oldDom) oldDom.remove();
  }

  // ─── Video events ─────────────────────────────────────────────────────────────
  vid.addEventListener('loadedmetadata', () => {
    S.loaded = true; S.duration = vid.duration;
    S.vw = vid.videoWidth; S.vh = vid.videoHeight;
    S.seeking = false; S.targetTime = 0;

    const canvas = get('vs-canvas');
    ctx = canvas.getContext('2d');
    canvas.width = S.vw; canvas.height = S.vh;
    fitCanvas();

    get('vs-fname').textContent = currentFile.name;
    get('vs-dur').textContent = S.duration.toFixed(2) + 's';
    get('vs-dim').textContent = S.vw + ' × ' + S.vh;
    get('vs-frames').textContent = '~' + Math.round(S.duration * 30) + ' frames';
    get('vs-tot').textContent = S.duration.toFixed(3) + 's';

    get('vs-info-bar').classList.remove('vs-hidden');
    get('vs-placeholder').classList.add('vs-hidden');
    get('vs-scroll-box').classList.remove('vs-hidden');
    get('vs-progress').classList.remove('vs-hidden');
    get('vs-time-row').classList.remove('vs-hidden');
    get('vs-spacer').style.height = S.sh + 'px';
    get('vs-ext-btn').disabled = false;

    seekTo(0); updateEstimate(); updateCode();
  });

  vid.addEventListener('seeked', () => {
    if (!ctx) return;
    const canvas = get('vs-canvas');
    ctx.drawImage(vid, 0, 0, canvas.width, canvas.height);
    const t = vid.currentTime;
    get('vs-cur').textContent = t.toFixed(3) + 's';
    const range = S.duration * (S.endPct - S.startPct);
    const elapsed = t - S.duration * S.startPct;
    const pct = range > 0 ? Math.max(0, Math.min(100, elapsed / range * 100)) : 0;
    get('vs-pfill').style.width = pct + '%';
    if (Math.abs(vid.currentTime - S.targetTime) > 0.016) vid.currentTime = S.targetTime;
    else S.seeking = false;
  });

  vid.addEventListener('error', () => {
    get('vs-placeholder').classList.remove('vs-hidden');
    get('vs-scroll-box').classList.add('vs-hidden');
    get('vs-placeholder').querySelector('.vs-placeholder-icon').textContent = '⚠️';
    get('vs-placeholder').querySelector('span') || get('vs-placeholder').appendChild(document.createElement('span'));
    get('vs-placeholder').lastElementChild.textContent = '無法讀取此影片格式';
  });

  function seekTo(t) {
    S.targetTime = Math.max(0, Math.min(S.duration, t));
    if (!S.seeking) { S.seeking = true; vid.currentTime = S.targetTime; }
  }

  // ─── Scroll preview ───────────────────────────────────────────────────────────
  const scrollBox = get('vs-scroll-box');
  scrollBox.addEventListener('scroll', () => {
    if (!S.loaded) return;
    const max = S.sh - scrollBox.clientHeight;
    if (max <= 0) return;
    renderAtFrac(Math.max(0, Math.min(1, scrollBox.scrollTop / max)));
  });

  function renderAtFrac(frac) {
    const effFrac = S.reverse ? 1 - frac : frac;

    if (S.previewMode === 'imgseq' && S.previewImgs && S.previewImgs.length > 0) {
      const idx = Math.round(effFrac * (S.previewImgs.length - 1));
      const img = S.previewImgs[idx];
      const canvas = get('vs-canvas');
      if (img.complete && img.naturalWidth > 0) ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      updateScrubUI(effFrac);
    } else if (S.previewMode === 'dom') {
      const container = document.getElementById('vs-dom-frames');
      if (container && container.children.length > 0) {
        const total = container.children.length;
        const idx = Math.round(effFrac * (total - 1));
        const cur = container.querySelector('.active');
        const next = container.children[idx];
        if (cur && next && cur !== next) { cur.classList.remove('active'); next.classList.add('active'); }
      }
      updateScrubUI(effFrac);
    } else {
      const st = S.duration * S.startPct, et = S.duration * S.endPct;
      seekTo(S.reverse ? et - frac * (et - st) : st + frac * (et - st));
    }
  }

  function updateScrubUI(effFrac) {
    get('vs-pfill').style.width = (effFrac * 100) + '%';
    const startT = S.duration * S.startPct, endT = S.duration * S.endPct;
    get('vs-cur').textContent = (startT + effFrac * (endT - startT)).toFixed(3) + 's';
  }

  // ─── Preview mode tabs ────────────────────────────────────────────────────────
  get('vs-prev-canvas').addEventListener('click', () => switchPreviewMode('canvas'));
  get('vs-prev-imgseq').addEventListener('click', () => switchPreviewMode('imgseq'));
  get('vs-prev-dom').addEventListener('click', () => switchPreviewMode('dom'));

  function switchPreviewMode(mode) {
    S.previewMode = mode;
    ['canvas', 'imgseq', 'dom'].forEach(m =>
      get('vs-prev-' + m).classList.toggle('vs-prev-tab--active', m === mode)
    );
    // Show/hide canvas vs DOM container
    const canvas = get('vs-canvas');
    const domContainer = document.getElementById('vs-dom-frames');
    if (mode === 'dom') {
      canvas.style.display = 'none';
      if (domContainer) domContainer.classList.remove('vs-hidden');
    } else {
      canvas.style.display = '';
      if (domContainer) domContainer.classList.add('vs-hidden');
    }
    // Re-render current position
    const max = S.sh - scrollBox.clientHeight;
    const frac = max > 0 ? Math.max(0, Math.min(1, scrollBox.scrollTop / max)) : 0;
    renderAtFrac(frac);
  }

  async function buildPreviewImgs(blobs) {
    // Overlay loading indicator on canvas area
    const wrap = get('vs-scroll-box').querySelector('.vs-canvas-wrap');
    const loader = document.createElement('div');
    loader.className = 'vs-imgseq-loading';
    loader.innerHTML = `<div class="vs-load-bar"><div class="vs-load-fill" id="vs-load-fill"></div></div><span id="vs-load-txt">建立預覽…</span>`;
    wrap.appendChild(loader);

    // Revoke old URLs
    S.previewBlobUrls.forEach(u => URL.revokeObjectURL(u));
    S.previewBlobUrls = blobs.map(b => URL.createObjectURL(b));

    // Preload images in batches so UI stays responsive
    const imgs = S.previewBlobUrls.map(url => { const img = new Image(); img.src = url; return img; });
    let done = 0;
    await Promise.all(imgs.map(img => new Promise(res => {
      if (img.complete) { done++; res(); return; }
      img.onload = img.onerror = () => {
        done++;
        const fill = document.getElementById('vs-load-fill');
        const txt = document.getElementById('vs-load-txt');
        if (fill) fill.style.width = (done / imgs.length * 100) + '%';
        if (txt) txt.textContent = `${done} / ${imgs.length}`;
        res();
      };
    })));

    wrap.removeChild(loader);
    S.previewImgs = imgs;

    // Build DOM frames container for DOM preview mode
    let domContainer = document.getElementById('vs-dom-frames');
    if (domContainer) domContainer.remove();
    domContainer = document.createElement('div');
    domContainer.id = 'vs-dom-frames';
    domContainer.className = 'vs-dom-frames vs-hidden';
    S.previewBlobUrls.forEach((url, i) => {
      const img = document.createElement('img');
      img.src = url;
      img.className = 'vs-dom-frame' + (i === 0 ? ' active' : '');
      domContainer.appendChild(img);
    });
    wrap.appendChild(domContainer);

    get('vs-prev-tabs').classList.remove('vs-hidden');
    get('vs-prev-imgseq').disabled = false;
    get('vs-prev-dom').disabled = false;
    switchPreviewMode('imgseq');
  }

  function fitCanvas() {
    const canvas = get('vs-canvas');
    const maxW = scrollBox.clientWidth || 400, maxH = 320;
    const r = S.vw / S.vh;
    let dw = maxW, dh = maxW / r;
    if (dh > maxH) { dh = maxH; dw = maxH * r; }
    canvas.style.width = Math.round(dw) + 'px';
    canvas.style.height = Math.round(dh) + 'px';
  }
  if (window.ResizeObserver) new ResizeObserver(() => { if (S.loaded) fitCanvas(); }).observe(scrollBox);

  // ─── Settings ─────────────────────────────────────────────────────────────────
  function initRange(id, lblId, fmt, onChange) {
    syncRange(id, lblId, fmt);
    get(id).addEventListener('input', () => { syncRange(id, lblId, fmt); onChange(parseFloat(get(id).value)); });
  }

  initRange('vs-sh', 'vs-sh-lbl', v => v + 'px', v => {
    S.sh = v;
    if (S.loaded) get('vs-spacer').style.height = v + 'px';
    updateCode();
  });
  initRange('vs-st', 'vs-st-lbl', v => v + '%', v => {
    S.startPct = v / 100;
    if (S.endPct <= S.startPct) { S.endPct = Math.min(1, S.startPct + 0.01); get('vs-et').value = Math.round(S.endPct * 100); syncRange('vs-et', 'vs-et-lbl', w => w + '%'); }
    if (S.loaded) seekTo(S.duration * S.startPct);
    updateCode(); updateEstimate();
  });
  initRange('vs-et', 'vs-et-lbl', v => v + '%', v => {
    S.endPct = v / 100;
    if (S.startPct >= S.endPct) { S.startPct = Math.max(0, S.endPct - 0.01); get('vs-st').value = Math.round(S.startPct * 100); syncRange('vs-st', 'vs-st-lbl', w => w + '%'); }
    if (S.loaded) seekTo(S.duration * S.endPct);
    updateCode(); updateEstimate();
  });
  get('vs-rev').addEventListener('change', e => { S.reverse = e.target.checked; updateCode(); });

  initRange('vs-fps', 'vs-fps-lbl', v => v + 'fps', v => { S.extFps = v; updateEstimate(); });
  initRange('vs-qual', 'vs-qual-lbl', v => v + '%', v => { S.extQuality = v / 100; updateEstimate(); });

  [['vs-fmt-jpg', 'image/jpeg'], ['vs-fmt-webp', 'image/webp']].forEach(([id, fmt]) => {
    get(id).addEventListener('click', () => {
      S.extFormat = fmt;
      get('vs-fmt-jpg').classList.toggle('vs-fmt-btn--active', fmt === 'image/jpeg');
      get('vs-fmt-webp').classList.toggle('vs-fmt-btn--active', fmt === 'image/webp');
      updateEstimate();
    });
  });

  function updateEstimate() {
    if (!S.loaded) return;
    const dur = S.duration * (S.endPct - S.startPct);
    const count = Math.max(1, Math.round(dur * S.extFps) + 1);
    const perFrameKB = Math.round((S.vw * S.vh / 921600) * (S.extFormat === 'image/webp' ? 55 : 85) * S.extQuality);
    const totalMB = (count * perFrameKB / 1024).toFixed(1);
    get('vs-estimate').innerHTML = `<strong>${count} 幀</strong> · ~${totalMB} MB`;
  }

  // ─── Frame Extraction ─────────────────────────────────────────────────────────
  const extBtn = get('vs-ext-btn');
  const dlBtn = get('vs-dl-btn');

  extBtn.addEventListener('click', () => {
    if (S.extracting) { S.extracting = false; return; }
    startExtraction();
  });

  dlBtn.addEventListener('click', () => {
    if (!S.zipData) return;
    const blob = new Blob([S.zipData], { type: 'application/zip' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
    const base = (currentFile ? currentFile.name.replace(/\.[^.]+$/, '') : 'frames');
    a.download = base + '_frames.zip'; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  });

  async function startExtraction() {
    if (!S.loaded || !currentFile) return;
    S.extracting = true;
    S.zipData = null;

    const dur = S.duration * (S.endPct - S.startPct);
    const startT = S.duration * S.startPct;
    const total = Math.max(1, Math.round(dur * S.extFps) + 1);
    S.totalFrames = total;

    extBtn.textContent = '取消';
    extBtn.classList.remove('vs-btn--primary'); extBtn.classList.add('vs-btn--cancel');
    dlBtn.classList.add('vs-hidden');
    get('vs-ext-prog').classList.remove('vs-hidden');
    get('vs-ext-fill').style.width = '0%';
    get('vs-ext-status').textContent = '準備中…';

    // Use a separate video element to avoid disturbing the preview
    const ev = document.createElement('video');
    ev.preload = 'auto'; ev.muted = true; ev.playsInline = true;
    ev.src = URL.createObjectURL(currentFile);

    await new Promise(r => { ev.addEventListener('loadedmetadata', r, { once: true }); ev.load(); });

    const ec = document.createElement('canvas');
    ec.width = S.vw; ec.height = S.vh;
    const ectx = ec.getContext('2d');
    const ext = S.extFormat === 'image/webp' ? 'webp' : 'jpg';

    const zip = new MiniZip();
    const collectedBlobs = [];

    for (let i = 0; i < total; i++) {
      if (!S.extracting) break;
      const t = total === 1 ? startT : startT + (i / (total - 1)) * dur;

      const blob = await new Promise(resolve => {
        ev.addEventListener('seeked', function handler() {
          ev.removeEventListener('seeked', handler);
          ectx.drawImage(ev, 0, 0, ec.width, ec.height);
          ec.toBlob(resolve, S.extFormat, S.extQuality);
        }, { once: true });
        ev.currentTime = t;
      });

      if (!S.extracting) break;
      collectedBlobs.push(blob);
      const buf = new Uint8Array(await blob.arrayBuffer());
      const name = `frames/frame_${String(i + 1).padStart(4, '0')}.${ext}`;
      await zip.add(name, buf);

      const pct = Math.round((i + 1) / total * 100);
      get('vs-ext-fill').style.width = pct + '%';
      get('vs-ext-status').textContent = `${i + 1} / ${total}`;
      S.extractedCount = i + 1;
    }

    URL.revokeObjectURL(ev.src);

    if (S.extracting && S.extractedCount === total) {
      S.zipData = zip.generate();
      S.extractedFrameCount = total;
      get('vs-ext-status').textContent = `完成 ${total} 幀，正在建立預覽…`;
      dlBtn.classList.remove('vs-hidden');
      updateCode();
      await buildPreviewImgs(collectedBlobs);
      get('vs-ext-status').textContent = `完成 ${total} 幀`;
    } else {
      get('vs-ext-status').textContent = `已取消（${S.extractedCount} 幀）`;
    }

    S.extracting = false;
    extBtn.textContent = '重新拆幀';
    extBtn.classList.remove('vs-btn--cancel'); extBtn.classList.add('vs-btn--primary');
  }

  // ─── Code Tabs ────────────────────────────────────────────────────────────────
  document.querySelectorAll('#video-scrubbing-app .vs-code-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      S.codeTab = tab.dataset.tab;
      document.querySelectorAll('#video-scrubbing-app .vs-code-tab').forEach(t => t.classList.toggle('vs-code-tab--active', t === tab));
      updateCode();
    });
  });

  // ─── Code Generation ─────────────────────────────────────────────────────────
  function updateCode() {
    const builders = { canvas: buildCanvasCode, imgseq: buildImgSeqCode, dom: buildDomCode };
    get('vs-code').textContent = (builders[S.codeTab] || buildCanvasCode)();
  }

  function buildCanvasCode() {
    const st = S.loaded ? (S.duration * S.startPct).toFixed(3) : '0.000';
    const et = S.loaded ? (S.duration * S.endPct).toFixed(3) : '10.000';
    return `<!-- Video Scrubbing — Canvas 版 (video.currentTime) -->
<!-- 影片 seek 受限於 keyframe 間距，適合快速製作 -->

<!-- 1. HTML -->
<section class="scrub-section" id="scrubSection">
  <div class="scrub-sticky">
    <canvas id="scrubCanvas"></canvas>
  </div>
</section>

<!-- 2. CSS -->
<style>
.scrub-section {
  height: ${S.sh}px; /* 調整此值控制滾動速度 */
  position: relative;
}
.scrub-sticky {
  position: sticky; top: 0; height: 100vh;
  display: flex; align-items: center; justify-content: center;
  background: #000; overflow: hidden;
}
#scrubCanvas { max-width: 100%; max-height: 100vh; display: block; }
</style>

<!-- 3. JS -->
<script>
(function () {
  var video = document.createElement('video');
  video.src = 'your-video.mp4'; /* ← 替換成影片路徑 */
  video.preload = 'auto'; video.muted = true; video.playsInline = true; video.load();

  var canvas = document.getElementById('scrubCanvas');
  var ctx    = canvas.getContext('2d');
  var section = document.getElementById('scrubSection');
  var START = ${st}, END = ${et}, REVERSE = ${S.reverse};
  var target = START, seeking = false;

  video.addEventListener('loadedmetadata', function () {
    canvas.width = video.videoWidth; canvas.height = video.videoHeight;
    video.currentTime = START;
  });
  video.addEventListener('seeked', function () {
    ctx.drawImage(video, 0, 0);
    if (Math.abs(video.currentTime - target) > 0.016) video.currentTime = target;
    else seeking = false;
  });
  window.addEventListener('scroll', function () {
    var rect = section.getBoundingClientRect();
    var total = section.offsetHeight - window.innerHeight;
    if (total <= 0) return;
    var frac = Math.max(0, Math.min(1, -rect.top / total));
    target = REVERSE ? END - frac*(END-START) : START + frac*(END-START);
    if (!seeking) { seeking = true; video.currentTime = target; }
  });
}());
<\/script>`;
  }

  function buildImgSeqCode() {
    const frameCount = S.extractedFrameCount || (S.loaded ? Math.max(1, Math.round(S.duration * (S.endPct - S.startPct) * S.extFps) + 1) : 120);
    const ext = S.extFormat === 'image/webp' ? 'webp' : 'jpg';
    const note = S.extractedFrameCount ? `已拆幀：${S.extractedFrameCount} 幀` : '（先完成拆幀，幀數將自動更新）';
    return `<!-- Video Scrubbing — Image Sequence 版 -->
<!-- ${note} -->
<!-- 適合追求逐格流暢的場景，如 Apple AirPods 頁面效果 -->
<!-- 將 frames/ 資料夾（從 ZIP 解壓）放在同目錄下 -->

<!-- 1. HTML -->
<section class="scrub-section" id="scrubSection">
  <canvas id="scrubCanvas"></canvas>
</section>

<!-- 2. CSS -->
<style>
.scrub-section {
  height: ${S.sh}px;
  position: relative;
}
#scrubCanvas {
  position: sticky; top: 0;
  max-width: 100%; max-height: 100vh; display: block;
  margin: 0 auto; background: #000;
}
</style>

<!-- 3. JS -->
<script>
(function () {
  var TOTAL   = ${frameCount};
  var EXT     = '${ext}';
  var REVERSE = ${S.reverse};
  var PREFIX  = 'frames/frame_'; /* ← 確認路徑 */

  var canvas  = document.getElementById('scrubCanvas');
  var ctx     = canvas.getContext('2d');
  var section = document.getElementById('scrubSection');
  var frames  = new Array(TOTAL);
  var loaded  = 0;
  var ready   = false;

  for (var i = 0; i < TOTAL; i++) {
    (function (idx) {
      var img = new Image();
      img.onload = function () {
        if (++loaded === TOTAL) {
          canvas.width  = frames[0].naturalWidth;
          canvas.height = frames[0].naturalHeight;
          ctx.drawImage(frames[0], 0, 0);
          ready = true;
        }
      };
      img.src = PREFIX + String(idx + 1).padStart(4, '0') + '.' + EXT;
      frames[idx] = img;
    }(i));
  }

  window.addEventListener('scroll', function () {
    if (!ready) return;
    var rect  = section.getBoundingClientRect();
    var total = section.offsetHeight - window.innerHeight;
    if (total <= 0) return;
    var frac = Math.max(0, Math.min(1, -rect.top / total));
    var idx  = Math.round((REVERSE ? 1 - frac : frac) * (TOTAL - 1));
    ctx.drawImage(frames[idx], 0, 0);
  });
}());
<\/script>`;
  }

  function buildDomCode() {
    const frameCount = S.extractedFrameCount || (S.loaded ? Math.max(1, Math.round(S.duration * (S.endPct - S.startPct) * S.extFps) + 1) : 120);
    const ext = S.extFormat === 'image/webp' ? 'webp' : 'jpg';
    const note = S.extractedFrameCount ? `已拆幀：${S.extractedFrameCount} 幀` : '（先完成拆幀，幀數將自動更新）';
    const imgTags = Array.from({ length: Math.min(frameCount, 5) }, (_, i) =>
      `      <img class="scrub-frame${i === 0 ? ' active' : ''}" src="frames/frame_${String(i + 1).padStart(4, '0')}.${ext}" alt="">`
    ).join('\n') + (frameCount > 5 ? `\n      <!-- … 共 ${frameCount} 張，其餘依序補上 -->` : '');

    return `<!-- Video Scrubbing — DOM 版 (純 img 標籤，無 canvas) -->
<!-- ${note} -->
<!-- 優點：CSS compositor 切換，無需 drawImage；不需 JS 即可顯示第一幀 -->
<!-- 注意：幀數多時 DOM 節點較重，建議 FPS ≤ 24、時長 ≤ 15s -->

<!-- 1. HTML：將 frames/ 資料夾（ZIP 解壓）放同目錄 -->
<section class="scrub-section" id="scrubSection">
  <div class="scrub-sticky">
${imgTags}
  </div>
</section>

<!-- 2. CSS -->
<style>
.scrub-section {
  height: ${S.sh}px;
  position: relative;
}
.scrub-sticky {
  position: sticky; top: 0; height: 100vh;
  display: flex; align-items: center; justify-content: center;
  background: #000; overflow: hidden;
}
.scrub-frame {
  position: absolute;
  max-width: 100%; max-height: 100vh;
  opacity: 0;          /* 全部隱藏 */
  will-change: opacity;
}
.scrub-frame.active {
  opacity: 1;          /* 只顯示當前幀 */
}
</style>

<!-- 3. JS：純 class 切換，不碰 canvas -->
<script>
(function () {
  var frames  = document.querySelectorAll('.scrub-frame');
  var section = document.getElementById('scrubSection');
  var REVERSE = ${S.reverse};
  var cur = 0;

  window.addEventListener('scroll', function () {
    var rect  = section.getBoundingClientRect();
    var total = section.offsetHeight - window.innerHeight;
    if (total <= 0) return;

    var frac = Math.max(0, Math.min(1, -rect.top / total));
    var idx  = Math.round((REVERSE ? 1 - frac : frac) * (frames.length - 1));

    if (idx !== cur) {
      frames[cur].classList.remove('active');
      frames[idx].classList.add('active');
      cur = idx;
    }
  });
}());
<\/script>`;
  }

  // ─── Copy ─────────────────────────────────────────────────────────────────────
  const copyBtn = get('vs-copy-btn');
  copyBtn.addEventListener('click', () => {
    copyText(get('vs-code').textContent).then(() => {
      copyBtn.classList.add('vs-btn--ok');
      copyBtn.innerHTML = `<svg width="13" height="13" viewBox="0 0 18 18" fill="none"><path d="M4 9l3.5 3.5L14 5" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg> 已複製!`;
      setTimeout(() => {
        copyBtn.classList.remove('vs-btn--ok');
        copyBtn.innerHTML = `<svg width="13" height="13" viewBox="0 0 18 18" fill="none"><rect x="6" y="6" width="10" height="10" rx="2" stroke="currentColor" stroke-width="1.5"/><path d="M12 6V4a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h2" stroke="currentColor" stroke-width="1.5"/></svg> 複製代碼`;
      }, 2000);
    });
  });

  updateCode();
})();
