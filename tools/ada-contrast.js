(function() {
  const css = `
    /* ADA Contrast Scoped Styles */
    #ada-app {
      padding: 32px;
      color: #f8fafc;
      font-family: 'Inter', sans-serif;
      height: 100%;
      overflow-y: auto;
    }
    #ada-app .ada-container {
      max-width: 1200px;
      margin: 0 auto;
      display: flex;
      flex-direction: column;
      gap: 2rem;
    }
    #ada-app header {
      text-align: center;
      margin-bottom: 1rem;
    }
    #ada-app header h1 {
      font-size: 2.2rem;
      font-weight: 700;
      background: linear-gradient(to right, #60a5fa, #a78bfa);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      margin-bottom: 0.5rem;
    }
    #ada-app header p {
      color: #94a3b8;
      font-size: 1rem;
    }
    #ada-app .main-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 2rem;
      background: rgba(25, 25, 38, 0.5);
      border: 1px solid rgba(255, 255, 255, 0.1);
      border-radius: 20px;
      padding: 2rem;
      backdrop-filter: blur(12px);
    }
    @media (max-width: 1024px) {
      #ada-app .main-grid { grid-template-columns: 1fr; }
    }
    @media (max-width: 640px) {
      #ada-app { padding: 16px; }
      #ada-app .main-grid { padding: 1rem; gap: 1rem; }
    }
    #ada-app .controls { display: flex; flex-direction: column; gap: 1.5rem; }
    #ada-app .control-block {
      background: rgba(0,0,0,0.2);
      border: 1px solid rgba(255, 255, 255, 0.05);
      border-radius: 12px;
      padding: 1.25rem;
      display: flex;
      flex-direction: column;
      gap: 1rem;
    }
    #ada-app .control-block-title {
      font-size: 1rem;
      font-weight: 600;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    #ada-app .color-input-wrapper {
      display: flex;
      align-items: center;
      background: rgba(0, 0, 0, 0.3);
      border: 1px solid rgba(255, 255, 255, 0.1);
      border-radius: 8px;
      padding: 0.5rem;
      gap: 1rem;
    }
    #ada-app .color-input-wrapper:focus-within { border-color: #3b82f6; }
    #ada-app input[type="color"] {
      -webkit-appearance: none;
      border: none;
      width: 36px;
      height: 36px;
      border-radius: 50%;
      cursor: pointer;
      background: none;
      padding: 0;
    }
    #ada-app input[type="color"]::-webkit-color-swatch-wrapper { padding: 0; }
    #ada-app input[type="color"]::-webkit-color-swatch {
      border: 2px solid rgba(255,255,255,0.2);
      border-radius: 50%;
    }
    #ada-app input[type="text"] {
      background: transparent;
      border: none;
      color: #fff;
      font-size: 1rem;
      font-family: 'JetBrains Mono', monospace;
      width: 100%;
      outline: none;
      text-transform: uppercase;
    }
    #ada-app .slider-group { display: flex; flex-direction: column; gap: 0.5rem; }
    #ada-app .slider-group label {
      font-size: 0.85rem;
      color: #94a3b8;
      display: flex;
      justify-content: space-between;
    }
    #ada-app input[type="range"] {
      -webkit-appearance: none;
      width: 100%;
      height: 6px;
      background: rgba(255, 255, 255, 0.1);
      border-radius: 6px;
      outline: none;
      margin-top: 5px;
    }
    #ada-app input[type="range"]::-webkit-slider-thumb {
      -webkit-appearance: none;
      appearance: none;
      width: 20px;
      height: 20px;
      border-radius: 50%;
      background: #fff;
      cursor: pointer;
      box-shadow: 0 2px 4px rgba(0,0,0,0.5);
    }
    #ada-app .slider-gradient {
      height: 10px !important;
      border-radius: 10px !important;
      border: 1px solid rgba(255,255,255,0.1);
    }
    #ada-app .results { display: flex; flex-direction: column; gap: 1.5rem; }
    #ada-app .result-block {
      background: rgba(0,0,0,0.2);
      border: 1px solid rgba(255, 255, 255, 0.05);
      border-radius: 12px;
      padding: 1.25rem;
      display: flex;
      flex-direction: column;
      gap: 1rem;
    }
    #ada-app .result-block-title { font-size: 0.9rem; color: #94a3b8; text-align: center; }
    #ada-app .ratio-value { text-align: center; font-size: 2.5rem; font-weight: 700; line-height: 1; transition: color 0.3s; }
    #ada-app .badges { display: grid; grid-template-columns: 1fr 1fr; gap: 0.5rem; }
    #ada-app .badge {
      display: flex; align-items: center; justify-content: space-between; gap: 4px;
      padding: 0.4rem 0.6rem; white-space: nowrap;
      border-radius: 6px; font-size: 0.8rem; font-weight: 600;
      background: rgba(0,0,0,0.2); border: 1px solid rgba(255,255,255,0.1);
    }
    #ada-app .badge.pass { background: rgba(16, 185, 129, 0.1); border-color: rgba(16, 185, 129, 0.3); color: #34d399; }
    #ada-app .badge.fail { background: rgba(239, 68, 68, 0.1); border-color: rgba(239, 68, 68, 0.3); color: #ef4444; }
    #ada-app .preview-section {
      margin-top: auto; border-radius: 12px; overflow: hidden;
      border: 1px solid rgba(255, 255, 255, 0.05); display: flex; flex-direction: column;
    }
    #ada-app .preview-header {
      background: rgba(0,0,0,0.3); padding: 0.75rem; text-align: center;
      font-size: 0.9rem; font-weight: 600; border-bottom: 1px solid rgba(255, 255, 255, 0.05);
    }
    #ada-app .preview-box {
      padding: 2rem; display: flex; align-items: center; justify-content: center;
      transition: background-color 0.3s; min-height: 200px;
    }
    #ada-app .preview-obj1 {
      padding: 1.5rem; border-radius: 12px; max-width: 320px; width: 100%;
      text-align: center; transition: background-color 0.3s;
    }
    #ada-app .preview-obj2 { transition: color 0.3s; }
    #ada-app .preview-obj2 h2 { font-size: 1.25rem; font-weight: 700; margin-bottom: 0.5rem; }
    #ada-app .preview-obj2 p { font-size: 0.9rem; line-height: 1.5; opacity: 0.9; margin-bottom: 1rem; }
    #ada-app .preview-btn {
      background-color: transparent; border: 2px solid currentColor;
      padding: 0.5rem 1rem; border-radius: 6px; font-weight: 600; cursor: pointer;
      width: 100%; color: inherit; transition: opacity 0.2s;
    }
    #ada-app .preview-btn:hover { opacity: 0.8; }
    #ada-app .ada-info-footer {
      margin-top: 1rem; background: rgba(0,0,0,0.2);
      border: 1px solid rgba(255, 255, 255, 0.05);
      border-radius: 12px; padding: 1.5rem; color: #94a3b8;
    }
    #ada-app .ada-info-footer h3 {
      font-size: 1rem; color: #e2e8f0; margin-bottom: 0.75rem; font-weight: 600;
    }
    #ada-app .ada-info-footer ul {
      list-style: none; display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 0.75rem;
    }
    #ada-app .ada-info-footer li {
      font-size: 0.85rem; background: rgba(255,255,255,0.03);
      padding: 0.5rem 0.75rem; border-radius: 6px; border: 1px solid rgba(255,255,255,0.02);
    }
    #ada-app .ada-info-footer li strong {
      color: #cbd5e1; display: block; margin-bottom: 0.25rem;
    }
    #ada-app .ada-info-footer p {
      font-size: 0.8rem; opacity: 0.7; margin-top: 1rem;
    }
  `;

  const html = `
    <div class="ada-container">
      <header>
        <h1>ADA Contrast</h1>
        <p>Multi-layer contrast checker</p>
      </header>

      <div class="main-grid">
        <div class="controls">
          <div class="control-block">
            <div class="control-block-title">Background (底色)</div>
            <div class="color-input-wrapper">
              <input type="color" id="ada-bg-color" value="#ffffff">
              <input type="text" id="ada-bg-hex" value="#FFFFFF" maxlength="7">
            </div>
            <div class="slider-group">
              <label>Lightness <span id="ada-bg-light-val" style="color:#60a5fa;">100%</span></label>
              <input type="range" id="ada-bg-slider" class="slider-gradient" min="0" max="100" value="100">
            </div>
          </div>

          <div class="control-block">
            <div class="control-block-title">Object 1 (外層)</div>
            <div class="color-input-wrapper">
              <input type="color" id="ada-obj1-color" value="#3b82f6">
              <input type="text" id="ada-obj1-hex" value="#3B82F6" maxlength="7">
            </div>
            <div class="slider-group">
              <label>Lightness <span id="ada-obj1-light-val" style="color:#60a5fa;">50%</span></label>
              <input type="range" id="ada-obj1-slider" class="slider-gradient" min="0" max="100" value="50">
            </div>
          </div>

          <div class="control-block">
            <div class="control-block-title">Object 2 (內層/文字)</div>
            <div class="color-input-wrapper">
              <input type="color" id="ada-obj2-color" value="#000000">
              <input type="text" id="ada-obj2-hex" value="#000000" maxlength="7">
            </div>
            <div class="slider-group">
              <label>Lightness <span id="ada-obj2-light-val" style="color:#60a5fa;">0%</span></label>
              <input type="range" id="ada-obj2-slider" class="slider-gradient" min="0" max="100" value="0">
            </div>
          </div>
        </div>

        <div class="results">
          <div class="two-col-results">
            <div class="result-block">
              <div class="result-block-title">BG vs Obj 1</div>
              <div class="ratio-value" id="ada-ratio1">1.00 : 1</div>
              <div class="badges" id="ada-badges1"></div>
            </div>
            <div class="result-block">
              <div class="result-block-title">Obj 1 vs Obj 2</div>
              <div class="ratio-value" id="ada-ratio2">1.00 : 1</div>
              <div class="badges" id="ada-badges2"></div>
            </div>
          </div>

          <div class="preview-section">
            <div class="preview-header">Live Preview</div>
            <div class="preview-box" id="ada-preview-bg">
              <div class="preview-obj1" id="ada-preview-obj1">
                <div class="preview-obj2" id="ada-preview-obj2">
                  <h2>Contrast Check</h2>
                  <p>Check outer card vs background, and inner text vs card.</p>
                  <button class="preview-btn">Interactive</button>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
      
      <div class="ada-info-footer">
        <h3>WCAG 2.2 對比度標準 (Contrast Ratio)</h3>
        <ul>
          <li><strong>AA Nrm (一般文字)</strong>最少 4.5 : 1</li>
          <li><strong>AA Lrg (大字體 / UI 元件)</strong>最少 3.0 : 1</li>
          <li><strong>AAA Nrm (高標準 一般文字)</strong>最少 7.0 : 1</li>
          <li><strong>AAA Lrg (高標準 大字體)</strong>最少 4.5 : 1</li>
        </ul>
        <p>* 大字體定義：至少 18pt（約 24px）或 14pt 粗體（約 18.5px）。UI 元件泛指按鈕邊界、輸入框等操作物件。</p>
      </div>
    </div>
  `;

  // Inject CSS & HTML
  const appContainer = document.getElementById('ada-app');
  const styleEl = document.createElement('style');
  styleEl.textContent = css;
  appContainer.appendChild(styleEl);
  appContainer.insertAdjacentHTML('beforeend', html);

  // Logic
  const root = document.getElementById('ada-app');
  const $ = id => document.getElementById(id);

  const els = {
    bgC: $('ada-bg-color'), bgH: $('ada-bg-hex'), bgS: $('ada-bg-slider'), bgL: $('ada-bg-light-val'),
    o1C: $('ada-obj1-color'), o1H: $('ada-obj1-hex'), o1S: $('ada-obj1-slider'), o1L: $('ada-obj1-light-val'),
    o2C: $('ada-obj2-color'), o2H: $('ada-obj2-hex'), o2S: $('ada-obj2-slider'), o2L: $('ada-obj2-light-val'),
    r1: $('ada-ratio1'), b1: $('ada-badges1'),
    r2: $('ada-ratio2'), b2: $('ada-badges2'),
    pBg: $('ada-preview-bg'), pO1: $('ada-preview-obj1'), pO2: $('ada-preview-obj2')
  };

  let hsls = {
    bg: { h:0, s:0, l:100 },
    o1: { h:217, s:90, l:60 },
    o2: { h:0, s:0, l:0 }
  };

  function hexToRgb(hex) {
    hex = hex.replace(/^#/, '');
    if (hex.length === 3) hex = hex.split('').map(c=>c+c).join('');
    let n = parseInt(hex, 16);
    return { r:(n>>16)&255, g:(n>>8)&255, b:n&255 };
  }
  function rgbToHex(r,g,b) { return "#"+((1<<24)+(r<<16)+(g<<8)+b).toString(16).slice(1).toUpperCase(); }
  function rgbToHsl(r,g,b) {
    r/=255; g/=255; b/=255;
    let max = Math.max(r,g,b), min = Math.min(r,g,b);
    let h, s, l = (max+min)/2;
    if(max===min){ h=s=0; } else {
      let d = max-min;
      s = l>0.5 ? d/(2-max-min) : d/(max+min);
      switch(max){
        case r: h=(g-b)/d+(g<b?6:0); break;
        case g: h=(b-r)/d+2; break;
        case b: h=(r-g)/d+4; break;
      }
      h/=6;
    }
    return { h:h*360, s:s*100, l:l*100 };
  }
  function hslToRgb(h,s,l) {
    h/=360; s/=100; l/=100;
    let r,g,b;
    if(s===0){ r=g=b=l; } else {
      const h2r = (p,q,t) => {
        if(t<0) t+=1; if(t>1) t-=1;
        if(t<1/6) return p+(q-p)*6*t;
        if(t<1/2) return q;
        if(t<2/3) return p+(q-p)*(2/3-t)*6;
        return p;
      };
      let q = l<0.5 ? l*(1+s) : l+s-l*s;
      let p = 2*l-q;
      r = h2r(p,q, h+1/3); g = h2r(p,q, h); b = h2r(p,q, h-1/3);
    }
    return { r:Math.round(r*255), g:Math.round(g*255), b:Math.round(b*255) };
  }
  function getLum(r,g,b) {
    const a = [r,g,b].map(v=>{ v/=255; return v<=0.03928 ? v/12.92 : Math.pow((v+0.055)/1.055, 2.4); });
    return a[0]*0.2126 + a[1]*0.7152 + a[2]*0.0722;
  }
  function contrast(h1,h2) {
    let r1=hexToRgb(h1), r2=hexToRgb(h2);
    if(!r1||!r2) return 1;
    let l1=getLum(r1.r,r1.g,r1.b), l2=getLum(r2.r,r2.g,r2.b);
    return (Math.max(l1,l2)+0.05)/(Math.min(l1,l2)+0.05);
  }

  function normHex(v) {
    v = v.trim(); if(!v.startsWith('#')) v='#'+v;
    if(/^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/.test(v)) {
      if(v.length===4) v='#'+v[1]+v[1]+v[2]+v[2]+v[3]+v[3];
      return v.toUpperCase();
    }
    return null;
  }

  function mkBadge(lbl, pass) {
    return `<div class="badge ${pass?'pass':'fail'}"><span>${lbl}</span><span>${pass?'Pass':'Fail'}</span></div>`;
  }
  function updateSliderGrad(el, hsl) {
    let s=hslToRgb(hsl.h,hsl.s,0), m=hslToRgb(hsl.h,hsl.s,50), e=hslToRgb(hsl.h,hsl.s,100);
    el.style.background = `linear-gradient(to right, rgb(${s.r},${s.g},${s.b}) 0%, rgb(${m.r},${m.g},${m.b}) 50%, rgb(${e.r},${e.g},${e.b}) 100%)`;
  }

  function sync() {
    let cBg = normHex(els.bgC.value) || '#FFFFFF';
    let cO1 = normHex(els.o1C.value) || '#3B82F6';
    let cO2 = normHex(els.o2C.value) || '#000000';

    els.bgH.value = cBg; els.o1H.value = cO1; els.o2H.value = cO2;
    els.pBg.style.backgroundColor = cBg;
    els.pO1.style.backgroundColor = cO1;
    els.pO2.style.color = cO2;

    updateSliderGrad(els.bgS, hsls.bg);
    updateSliderGrad(els.o1S, hsls.o1);
    updateSliderGrad(els.o2S, hsls.o2);

    let rat1 = contrast(cBg, cO1), rat2 = contrast(cO1, cO2);
    els.r1.textContent = rat1.toFixed(2)+' : 1';
    els.r1.style.color = rat1>=4.5?'#34d399':rat1>=3?'#fbbf24':'#ef4444';
    els.b1.innerHTML = mkBadge('AA Nrm', rat1>=4.5)+mkBadge('AA Lrg', rat1>=3)+mkBadge('AAA Nrm', rat1>=7)+mkBadge('AAA Lrg', rat1>=4.5);

    els.r2.textContent = rat2.toFixed(2)+' : 1';
    els.r2.style.color = rat2>=4.5?'#34d399':rat2>=3?'#fbbf24':'#ef4444';
    els.b2.innerHTML = mkBadge('AA Nrm', rat2>=4.5)+mkBadge('AA Lrg', rat2>=3)+mkBadge('AAA Nrm', rat2>=7)+mkBadge('AAA Lrg', rat2>=4.5);
  }

  function hColor(key, v, us=true) {
    let hx = normHex(v); if(!hx) return;
    els[key+'C'].value = hx;
    let rgb = hexToRgb(hx); hsls[key] = rgbToHsl(rgb.r,rgb.g,rgb.b);
    if(us) { els[key+'S'].value = Math.round(hsls[key].l); els[key+'L'].textContent = Math.round(hsls[key].l)+'%'; }
    sync();
  }
  function hSlider(key, v) {
    els[key+'L'].textContent = v+'%'; hsls[key].l = v;
    let rgb=hslToRgb(hsls[key].h, hsls[key].s, hsls[key].l);
    let hx = rgbToHex(rgb.r,rgb.g,rgb.b);
    els[key+'C'].value=hx; els[key+'H'].value=hx;
    sync();
  }

  ['bg','o1','o2'].forEach(k=>{
    els[k+'C'].addEventListener('input', e=>hColor(k, e.target.value));
    els[k+'H'].addEventListener('input', e=>hColor(k, e.target.value));
    els[k+'S'].addEventListener('input', e=>hSlider(k, e.target.value));
  });

  hColor('bg', els.bgC.value); hColor('o1', els.o1C.value); hColor('o2', els.o2C.value);
})();
