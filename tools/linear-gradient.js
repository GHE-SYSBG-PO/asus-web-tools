(function() {
  const css = `
    #gradient-app {
      --g-bg-primary: #0d0d12;
      --g-bg-card: rgba(25, 25, 38, 0.4);
      --g-bg-glass: rgba(255, 255, 255, 0.04);
      --g-bg-glass-hover: rgba(255, 255, 255, 0.08);
      --g-border-color: rgba(255, 255, 255, 0.08);
      --g-border-active: rgba(139, 92, 246, 0.5);
      --g-text-primary: #f0f0f5;
      --g-text-secondary: #8888a0;
      --g-text-muted: #55556a;
      --g-accent: #8b5cf6;
      --g-accent-glow: rgba(139, 92, 246, 0.25);
      --g-success: #34d399;
      --g-radius-sm: 8px;
      --g-radius-md: 12px;
      --g-radius-lg: 16px;
      --g-radius-xl: 20px;
      --g-font-sans: 'Inter', sans-serif;
      --g-font-mono: 'JetBrains Mono', monospace;
      --g-transition: 0.2s cubic-bezier(0.4, 0, 0.2, 1);
      --g-checker-size: 12px;

      font-family: var(--g-font-sans);
      color: var(--g-text-primary);
      height: 100%;
      overflow-y: auto;
      padding: 32px 20px;
      display: flex;
      justify-content: center;
    }

    #gradient-app .lg-app {
      width: 100%;
      max-width: 560px;
      display: flex;
      flex-direction: column;
      gap: 24px;
    }

    #gradient-app .lg-header { text-align: center; }
    #gradient-app .lg-header__title { font-size: 1.8rem; font-weight: 700; display: flex; align-items: center; justify-content: center; gap: 10px; }
    #gradient-app .lg-header__icon { background: linear-gradient(135deg, #8b5cf6, #3b82f6, #06b6d4); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    #gradient-app .lg-header__subtitle { color: var(--g-text-secondary); font-size: 0.9rem; margin-top: 6px; }

    #gradient-app .lg-tabs {
      display: flex; gap: 4px; background: var(--g-bg-card);
      border: 1px solid var(--g-border-color); border-radius: var(--g-radius-lg);
      padding: 6px; backdrop-filter: blur(12px);
    }
    #gradient-app .lg-tab {
      flex: 1; display: flex; align-items: center; justify-content: center; gap: 8px;
      padding: 10px 8px; border: none; border-radius: var(--g-radius-md);
      background: transparent; color: var(--g-text-secondary);
      font-size: 0.85rem; font-weight: 500; cursor: pointer; transition: all var(--g-transition);
    }
    #gradient-app .lg-tab:hover { color: var(--g-text-primary); background: var(--g-bg-glass-hover); }
    #gradient-app .lg-tab--active { color: var(--g-text-primary); background: var(--g-bg-glass-hover); box-shadow: 0 0 0 1px var(--g-border-active), 0 2px 8px var(--g-accent-glow); }

    #gradient-app .lg-tab__dot { width: 10px; height: 10px; border-radius: 3px; }
    #gradient-app .lg-tab__dot--solid { background: #8b5cf6; }
    #gradient-app .lg-tab__dot--solid-alpha { background: linear-gradient(135deg, rgba(139, 92, 246, 0.5), rgba(139, 92, 246, 0.5)); border: 1px dashed rgba(139, 92, 246, 0.6); }
    #gradient-app .lg-tab__dot--gradient { background: linear-gradient(135deg, #8b5cf6, #3b82f6); }
    #gradient-app .lg-tab__dot--gradient-alpha { background: linear-gradient(135deg, rgba(139, 92, 246, 0.7), rgba(59, 130, 246, 0.4)); border: 1px dashed rgba(139, 92, 246, 0.5); }

    #gradient-app .lg-preview {
      position: relative; height: 200px; border-radius: var(--g-radius-xl);
      overflow: hidden; border: 1px solid var(--g-border-color);
    }
    #gradient-app .lg-preview__checker {
      position: absolute; inset: 0;
      background-image: linear-gradient(45deg, #1a1a2e 25%, transparent 25%), linear-gradient(-45deg, #1a1a2e 25%, transparent 25%), linear-gradient(45deg, transparent 75%, #1a1a2e 75%), linear-gradient(-45deg, transparent 75%, #1a1a2e 75%);
      background-size: calc(var(--g-checker-size) * 2) calc(var(--g-checker-size) * 2);
      background-position: 0 0, 0 var(--g-checker-size), var(--g-checker-size) calc(-1 * var(--g-checker-size)), calc(-1 * var(--g-checker-size)) 0;
      background-color: #12121e;
    }
    #gradient-app .lg-preview__gradient { position: absolute; inset: 0; transition: background-image 0.15s ease; }
    #gradient-app .lg-preview__label { position: absolute; bottom: 10px; right: 14px; font-size: 0.65rem; font-weight: 600; text-transform: uppercase; color: rgba(255,255,255,0.3); }

    #gradient-app .lg-gradient-bar-section { display: flex; flex-direction: column; gap: 8px; margin-top: 1rem; }
    #gradient-app .lg-gradient-bar-wrapper { position: relative; display: flex; align-items: center; gap: 12px; }
    #gradient-app .lg-gradient-bar { position: relative; flex: 1; height: 32px; border-radius: 16px; cursor: crosshair; }
    #gradient-app .lg-gradient-bar__track { position: absolute; inset: 0; border-radius: 16px; border: 1px solid var(--g-border-color); overflow: hidden; pointer-events: none; }
    #gradient-app .lg-gradient-bar__track::before {
      content: ''; position: absolute; inset: 0;
      background-image: linear-gradient(45deg, #1a1a2e 25%, transparent 25%), linear-gradient(-45deg, #1a1a2e 25%, transparent 25%), linear-gradient(45deg, transparent 75%, #1a1a2e 75%), linear-gradient(-45deg, transparent 75%, #1a1a2e 75%);
      background-size: 12px 12px; background-position: 0 0, 0 6px, 6px -6px, -6px 0; background-color: #12121e;
    }
    #gradient-app .lg-gradient-bar__track::after { content: ''; position: absolute; inset: 0; border-radius: 16px; }

    #gradient-app .lg-gradient-bar__add {
      position: absolute;
      right: 0;
      top: 130%;
      display: flex; align-items: center; justify-content: center; width: 32px; height: 32px;
      border-radius: 50%; border: 1px solid var(--g-border-color); background: var(--g-bg-card);
      color: var(--g-text-secondary); cursor: pointer; transition: all var(--g-transition); flex-shrink: 0;
    }
    #gradient-app .lg-gradient-bar__add:hover { color: var(--g-accent); border-color: var(--g-accent); background: var(--g-accent-glow); }
    #gradient-app .lg-gradient-bar__hint { font-size: 0.75rem; color: var(--g-text-muted); text-align: center; }

    #gradient-app .lg-stop-handle {
      position: absolute; top: 50%; width: 22px; height: 22px; border-radius: 50%; border: 3px solid #fff;
      transform: translate(-50%, -50%); cursor: grab; box-shadow: 0 1px 4px rgba(0,0,0,0.5); z-index: 2;
    }
    #gradient-app .lg-stop-handle.dragging { cursor: grabbing; transform: translate(-50%, -50%) scale(1.2); box-shadow: 0 2px 12px rgba(0,0,0,0.7), 0 0 0 2px var(--g-accent); }
    #gradient-app .lg-stop-handle.active { box-shadow: 0 2px 8px rgba(0,0,0,0.6), 0 0 0 3px var(--g-accent); }

    #gradient-app .lg-controls { display: flex; flex-direction: column; gap: 24px; margin-top: 1rem; }
    #gradient-app .lg-control-group { display: flex; flex-direction: column; gap: 12px; }
    #gradient-app .lg-control-label { font-size: 0.8rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: var(--g-text-secondary); }

    #gradient-app .lg-angle-control { display: flex; align-items: center; gap: 20px; }
    #gradient-app .lg-angle-wheel {
      position: relative; width: 64px; height: 64px; border-radius: 50%;
      border: 2px solid var(--g-border-color); background: var(--g-bg-card); cursor: pointer;
    }
    #gradient-app .lg-angle-wheel__indicator {
      position: absolute; top: 50%; left: 50%; width: 2px; height: 45%; background: var(--g-accent);
      transform-origin: bottom center; transform: translate(-50%, -100%) rotate(90deg); pointer-events: none;
    }
    #gradient-app .lg-angle-wheel__indicator::after {
      content: ''; position: absolute; top: -2px; left: 50%; transform: translateX(-50%);
      width: 6px; height: 6px; border-radius: 50%; background: var(--g-accent);
    }
    #gradient-app .lg-angle-input-wrap { display: flex; align-items: center; gap: 4px; background: var(--g-bg-card); border: 1px solid var(--g-border-color); border-radius: var(--g-radius-sm); padding: 10px 16px; }
    #gradient-app .lg-angle-input { width: 48px; border: none; background: transparent; color: var(--g-text-primary); font-family: var(--g-font-mono); font-size: 1rem; outline: none; text-align: right; }
    #gradient-app .lg-angle-unit { color: var(--g-text-muted); font-family: var(--g-font-mono); font-size: 0.9rem; }

    #gradient-app .lg-stops-list { display: flex; flex-direction: column; gap: 10px; }
    #gradient-app .lg-stop-item {
      display: flex; align-items: center; gap: 12px; padding: 12px 16px; background: rgba(0,0,0,0.2);
      border: 1px solid var(--g-border-color); border-radius: var(--g-radius-md); transition: border-color 0.2s;
    }
    #gradient-app .lg-stop-item.active { border-color: var(--g-border-active); box-shadow: 0 0 0 1px var(--g-accent-glow); }
    #gradient-app .lg-stop-item__color-wrap { position: relative; width: 36px; height: 36px; border-radius: var(--g-radius-sm); overflow: hidden; border: 1px solid var(--g-border-color); flex-shrink: 0; }
    #gradient-app .lg-stop-item__color-checker { position: absolute; inset: 0; background-color: #12121e; }
    #gradient-app .lg-stop-item__color-preview { position: absolute; inset: 0; pointer-events: none; }
    #gradient-app .lg-stop-item__color-input { position: absolute; inset: 0; width: 100%; height: 100%; opacity: 0.01; cursor: pointer; }
    #gradient-app .lg-stop-item__hex { width: 80px; border: none; background: transparent; color: var(--g-text-primary); font-family: var(--g-font-mono); font-size: 0.9rem; outline: none; text-transform: uppercase; }

    #gradient-app .lg-stop-item__alpha-group { display: flex; align-items: center; gap: 10px; flex: 1; }
    #gradient-app .lg-stop-item__alpha-slider { flex: 1; -webkit-appearance: none; height: 6px; border-radius: 3px; outline: none; }
    #gradient-app .lg-stop-item__alpha-slider::-webkit-slider-thumb { -webkit-appearance: none; width: 16px; height: 16px; border-radius: 50%; border: 2px solid var(--g-text-primary); background: var(--g-accent); cursor: grab; }
    #gradient-app .lg-stop-item__alpha-value { width: 44px; text-align: right; border: none; background: transparent; color: var(--g-text-secondary); font-family: var(--g-font-mono); font-size: 0.85rem; outline: none; }

    #gradient-app .lg-stop-item__position-group { display: flex; align-items: center; gap: 4px; }
    #gradient-app .lg-stop-item__position { width: 40px; text-align: right; border: none; background: transparent; color: var(--g-text-secondary); font-family: var(--g-font-mono); font-size: 0.85rem; outline: none; }

    #gradient-app .lg-stop-item__delete { display: flex; align-items: center; justify-content: center; width: 28px; height: 28px; border: none; border-radius: 6px; background: transparent; color: var(--g-text-muted); cursor: pointer; transition: all 0.2s; opacity: 0; }
    #gradient-app .lg-stop-item:hover .lg-stop-item__delete { opacity: 1; }
    #gradient-app .lg-stop-item__delete:hover { color: #ef4444; background: rgba(239, 68, 68, 0.15); }

    #gradient-app .lg-output { margin-top: 1rem; }
    #gradient-app .lg-output__box { position: relative; background: rgba(0,0,0,0.3); border: 1px solid var(--g-border-color); border-radius: var(--g-radius-md); padding: 16px 50px 16px 16px; }
    #gradient-app .lg-output__code { font-family: var(--g-font-mono); font-size: 0.85rem; color: var(--g-text-primary); white-space: pre-wrap; word-break: break-all; }
    #gradient-app .lg-output__copy { position: absolute; top: 12px; right: 12px; display: flex; align-items: center; justify-content: center; width: 36px; height: 36px; border: 1px solid var(--g-border-color); border-radius: var(--g-radius-sm); background: var(--g-bg-card); color: var(--g-text-secondary); cursor: pointer; transition: all 0.2s; }
    #gradient-app .lg-output__copy:hover { color: var(--g-accent); border-color: var(--g-accent); background: var(--g-accent-glow); }
    #gradient-app .lg-output__copy.copied { color: var(--g-success); border-color: var(--g-success); background: rgba(52, 211, 153, 0.1); }

    #gradient-app .lg-hidden { display: none !important; }
  `;

  const html = `
    <div class="lg-app">
      <header class="lg-header">
        <h1 class="lg-header__title">
          <span class="lg-header__icon">◆</span>
          CSS Gradient Generator
        </h1>
        <p class="lg-header__subtitle">產生完美線性漸層與透明度</p>
      </header>

      <nav class="lg-tabs" role="tablist">
        <button class="lg-tab lg-tab--active" data-mode="solid">
          <span class="lg-tab__dot lg-tab__dot--solid"></span> 單色
        </button>
        <button class="lg-tab" data-mode="solid-alpha">
          <span class="lg-tab__dot lg-tab__dot--solid-alpha"></span> 單色透明
        </button>
        <button class="lg-tab" data-mode="gradient">
          <span class="lg-tab__dot lg-tab__dot--gradient"></span> 漸層
        </button>
        <button class="lg-tab" data-mode="gradient-alpha">
          <span class="lg-tab__dot lg-tab__dot--gradient-alpha"></span> 漸層透明
        </button>
      </nav>

      <main class="lg-main">
        <section class="lg-preview">
          <div class="lg-preview__checker"></div>
          <div class="lg-preview__gradient" id="lg-preview"></div>
          <div class="lg-preview__label">Preview</div>
        </section>

        <section class="lg-gradient-bar-section" id="lg-bar-section">
          <div class="lg-gradient-bar-wrapper">
            <div class="lg-gradient-bar" id="lg-bar">
              <div class="lg-gradient-bar__track" id="lg-track"></div>
            </div>
            <button class="lg-gradient-bar__add" id="lg-add-btn" title="新增色票">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M8 3v10M3 8h10" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
            </button>
          </div>
          <p class="lg-gradient-bar__hint">點擊漸層條新增色票 ・ 拖曳調整位置 ・ 雙擊刪除</p>
        </section>

        <section class="lg-controls">
          <div class="lg-control-group" id="lg-angle-group">
            <label class="lg-control-label">角度</label>
            <div class="lg-angle-control">
              <div class="lg-angle-wheel" id="lg-wheel">
                <div class="lg-angle-wheel__indicator" id="lg-indicator"></div>
              </div>
              <div class="lg-angle-input-wrap">
                <input type="number" class="lg-angle-input" id="lg-angle-in" value="90">
                <span class="lg-angle-unit">deg</span>
              </div>
            </div>
          </div>

          <div class="lg-control-group" id="lg-stops-group">
            <label class="lg-control-label">色票</label>
            <div class="lg-stops-list" id="lg-stops-list"></div>
          </div>
        </section>

        <section class="lg-output">
          <label class="lg-control-label">CSS 輸出</label>
          <div class="lg-output__box">
            <pre class="lg-output__code" id="lg-out-code"></pre>
            <button class="lg-output__copy" id="lg-copy-btn">
              <svg class="lg-copy-icon" width="18" height="18" viewBox="0 0 18 18" fill="none"><rect x="6" y="6" width="10" height="10" rx="2" stroke="currentColor" stroke-width="1.5"/><path d="M12 6V4a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h2" stroke="currentColor" stroke-width="1.5"/></svg>
              <svg class="lg-check-icon" width="18" height="18" viewBox="0 0 18 18" fill="none" style="display:none"><path d="M4 9l3.5 3.5L14 5" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
            </button>
          </div>
        </section>
      </main>
    </div>
  `;

  const app = document.getElementById('gradient-app');
  const styleEl = document.createElement('style');
  styleEl.textContent = css;
  app.appendChild(styleEl);
  app.insertAdjacentHTML('beforeend', html);

  // -- JS Logic --
  const $ = sel => app.querySelector(sel);
  const $$ = sel => app.querySelectorAll(sel);

  const state = {
    mode: 'solid', angle: 90, activeStopIndex: 0,
    stops: [ { color: '#8b5cf6', alpha: 1, position: 0 }, { color: '#3b82f6', alpha: 1, position: 100 } ]
  };

  const dom = {
    tabs: $$('.lg-tab'), preview: $('#lg-preview'),
    barSec: $('#lg-bar-section'), bar: $('#lg-bar'), track: $('#lg-track'), addBtn: $('#lg-add-btn'),
    angGrp: $('#lg-angle-group'), wheel: $('#lg-wheel'), ind: $('#lg-indicator'), angIn: $('#lg-angle-in'),
    list: $('#lg-stops-list'), out: $('#lg-out-code'), copy: $('#lg-copy-btn')
  };

  function hexToRgb(h) {
    let x = h.replace('#','');
    if(x.length===3) x=x.split('').map(c=>c+c).join('');
    let n=parseInt(x,16); return {r:(n>>16)&255, g:(n>>8)&255, b:n&255};
  }
  function isGrad() { return state.mode.startsWith('gradient'); }
  function hasAlpha() { return state.mode.includes('alpha'); }

  function buildCol(s) {
    let {r,g,b} = hexToRgb(s.color);
    return hasAlpha() ? `rgba(${r},${g},${b},${s.alpha})` : `rgb(${r},${g},${b})`;
  }
  function buildCSS() {
    let st = [...state.stops].sort((a,b)=>a.position-b.position);
    if(!isGrad()) {
      let c = buildCol(st[0]); return `linear-gradient(${c}, ${c})`;
    }
    let p = st.map(s=>`${buildCol(s)} ${s.position}%`).join(', ');
    return `linear-gradient(${state.angle}deg, ${p})`;
  }

  function updateVisuals() {
    dom.preview.style.backgroundImage = buildCSS();

    // Bar Track
    let st = [...state.stops].sort((a,b)=>a.position-b.position);
    let gradStr = isGrad() ? `linear-gradient(90deg, ${st.map(s=>`${buildCol(s)} ${s.position}%`).join(',')})` : `${buildCol(st[0])}`;
    dom.track.setAttribute('style', `background: ${gradStr};`);

    let handles = $$('.lg-stop-handle');
    if(handles.length === state.stops.length) {
      state.stops.forEach((s,i) => {
        handles[i].className = 'lg-stop-handle' + (i===state.activeStopIndex?' active':'');
        handles[i].style.left = `${s.position}%`;
        handles[i].style.backgroundColor = buildCol(s);
      });
    }

    let items = $$('.lg-stop-item');
    let toShow = isGrad() ? state.stops : [state.stops[0]];
    if(items.length === toShow.length) {
      toShow.forEach((s,i) => {
        let {r,g,b} = hexToRgb(s.color);
        let previewColor = hasAlpha() ? `rgba(${r},${g},${b},${s.alpha})` : s.color;
        items[i].className = 'lg-stop-item' + (i===state.activeStopIndex?' active':'');
        items[i].style.order = Math.round(s.position);
         
        let cPrev = items[i].querySelector('.lg-stop-item__color-preview');
        cPrev.style.background = previewColor;
        
        let cIn = items[i].querySelector('.lg-stop-item__color-input');
        if(cIn.value !== s.color) cIn.value = s.color;
        
        let hIn = items[i].querySelector('.lg-stop-item__hex');
        if(document.activeElement !== hIn) hIn.value = s.color.toUpperCase();

        if(hasAlpha()) {
          let aSli = items[i].querySelector('.lg-stop-item__alpha-slider');
          if(aSli.value !== String(Math.round(s.alpha*100))) aSli.value = Math.round(s.alpha*100);
          aSli.style.background = `linear-gradient(to right, rgba(${r},${g},${b},0), rgba(${r},${g},${b},1))`;
          
          let aVal = items[i].querySelector('.lg-stop-item__alpha-value');
          if(document.activeElement !== aVal) aVal.value = Math.round(s.alpha*100)+'%';
        }
        if(isGrad()) {
          let pIn = items[i].querySelector('.lg-stop-item__position');
          if(document.activeElement !== pIn) pIn.value = Math.round(s.position);
        }
      });
    }

    dom.ind.style.transform = `translate(-50%, -100%) rotate(${state.angle}deg)`;
    if(document.activeElement !== dom.angIn) dom.angIn.value = state.angle;
    dom.out.textContent = `${buildCSS()}`;
  }

  function render() {
    dom.preview.style.backgroundImage = buildCSS();

    // Bar Track
    let st = [...state.stops].sort((a,b)=>a.position-b.position);
    let gradStr = isGrad() ? `linear-gradient(90deg, ${st.map(s=>`${buildCol(s)} ${s.position}%`).join(',')})` : `${buildCol(st[0])}`;
    dom.track.setAttribute('style', `background: ${gradStr};`);

    // Handles
    $$('.lg-stop-handle').forEach(h=>h.remove());
    if(isGrad()) {
      state.stops.forEach((s,i) => {
        let h = document.createElement('div');
        h.className = 'lg-stop-handle' + (i===state.activeStopIndex?' active':'');
        h.style.left = `${s.position}%`; h.style.backgroundColor = buildCol(s);
        h.addEventListener('mousedown', e=>onDrag(e,i));
        h.addEventListener('dblclick', ()=>rmStop(i));
        h.addEventListener('click', e=>{e.stopPropagation(); state.activeStopIndex=i; updateVisuals();});
        dom.bar.appendChild(h);
      });
    }

    // List
    dom.list.innerHTML = '';
    let toShow = isGrad() ? state.stops : [state.stops[0]];
    toShow.forEach((s,i) => {
      let {r,g,b} = hexToRgb(s.color);
      let previewColor = hasAlpha() ? `rgba(${r},${g},${b},${s.alpha})` : s.color;

      let el = document.createElement('div');
      el.className = 'lg-stop-item' + (i===state.activeStopIndex?' active':'');
      el.style.order = Math.round(s.position);
      el.innerHTML = `
        <div class="lg-stop-item__color-wrap">
          <div class="lg-stop-item__color-checker"></div>
          <div class="lg-stop-item__color-preview" style="background:${previewColor}"></div>
          <input type="color" class="lg-stop-item__color-input" value="${s.color}">
        </div>
        <input type="text" class="lg-stop-item__hex" value="${s.color.toUpperCase()}" maxlength="7">
        ${hasAlpha() ? `<div class="lg-stop-item__alpha-group">
          <input type="range" class="lg-stop-item__alpha-slider" min="0" max="100" value="${Math.round(s.alpha*100)}" style="background: linear-gradient(to right, rgba(${r},${g},${b},0), rgba(${r},${g},${b},1))">
          <input type="text" class="lg-stop-item__alpha-value" value="${Math.round(s.alpha*100)}%">
        </div>` : ''}
        ${isGrad() ? `<div class="lg-stop-item__position-group">
          <input type="number" class="lg-stop-item__position" value="${Math.round(s.position)}">%
        </div>` : ''}
        <button class="lg-stop-item__delete ${!isGrad()||toShow.length<=2?'lg-hidden':''}"><svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M3 3l8 8M11 3l-8 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg></button>
      `;

      el.querySelector('.lg-stop-item__color-input').addEventListener('input', e => { state.stops[i].color=e.target.value; updateVisuals(); });
      el.querySelector('.lg-stop-item__hex').addEventListener('input', e => {
        let v=e.target.value; if(!v.startsWith('#')) v='#'+v;
        if(/^#[0-9A-Fa-f]{6}$/.test(v)) { state.stops[i].color=v.toLowerCase(); updateVisuals(); }
      });
      if(hasAlpha()) {
        el.querySelector('.lg-stop-item__alpha-slider').addEventListener('input', e=>{state.stops[i].alpha=e.target.value/100; updateVisuals();});
        el.querySelector('.lg-stop-item__alpha-value').addEventListener('input', e=>{let v=parseInt(e.target.value); if(!isNaN(v)){ state.stops[i].alpha=Math.max(0,Math.min(100,v))/100; updateVisuals(); }});
      }
      if(isGrad()) {
        el.querySelector('.lg-stop-item__position').addEventListener('input', e=>{let v=parseInt(e.target.value); if(!isNaN(v)){ state.stops[i].position=Math.max(0,Math.min(100,v)); updateVisuals(); }});
        el.querySelector('.lg-stop-item__delete').addEventListener('click', ()=>rmStop(i));
      }
      el.addEventListener('click', ()=>{state.activeStopIndex=i; updateVisuals();});
      dom.list.appendChild(el);
    });

    // Angle & Output
    dom.ind.style.transform = `translate(-50%, -100%) rotate(${state.angle}deg)`;
    dom.angIn.value = state.angle;
    dom.out.textContent = `${buildCSS()}`;

    dom.angGrp.classList.toggle('lg-hidden', !isGrad());
    dom.addBtn.classList.toggle('lg-hidden', !isGrad());
    app.querySelector('.lg-gradient-bar__hint').classList.toggle('lg-hidden', !isGrad());
  }

  function addStop(pos) {
    state.stops.push({ color:'#000000', alpha:1, position:pos });
    state.activeStopIndex = state.stops.length-1; render();
  }
  function rmStop(idx) {
    if(state.stops.length<=2) return;
    state.stops.splice(idx,1); state.activeStopIndex = Math.max(0, idx-1); render();
  }

  function onDrag(e, i) {
    e.preventDefault(); e.stopPropagation(); state.activeStopIndex = i; updateVisuals();
    const r = dom.bar.getBoundingClientRect();
    const move = ev => {
      let pct = Math.max(0, Math.min(100, ((ev.clientX-r.left)/r.width)*100));
      state.stops[i].position = Math.round(pct); updateVisuals();
    };
    const up = () => { document.removeEventListener('mousemove',move); document.removeEventListener('mouseup',up); };
    document.addEventListener('mousemove',move); document.addEventListener('mouseup',up);
  }

  dom.tabs.forEach(t => t.addEventListener('click', () => {
    state.mode = t.dataset.mode;
    dom.tabs.forEach(btn => btn.classList.toggle('lg-tab--active', btn===t));
    if(!isGrad()) state.activeStopIndex=0;
    render();
  }));

  dom.bar.addEventListener('click', e => {
    if(!isGrad() || e.target.closest('.lg-stop-handle')) return;
    const r = dom.bar.getBoundingClientRect();
    addStop(Math.round(((e.clientX-r.left)/r.width)*100));
  });

  dom.addBtn.addEventListener('click', () => { if(isGrad()) addStop(50); });

  // Angle Wheel
  let angDrag = false;
  const updAng = ev => {
    const r = dom.wheel.getBoundingClientRect();
    let cx=r.left+r.width/2, cy=r.top+r.height/2;
    let ang = Math.round(Math.atan2(ev.clientX-cx, cy-ev.clientY)*(180/Math.PI));
    if(ang<0) ang+=360; state.angle=ang; updateVisuals();
  };
  dom.wheel.addEventListener('mousedown', e=>{angDrag=true; updAng(e);});
  document.addEventListener('mousemove', e=>{if(angDrag) updAng(e);});
  document.addEventListener('mouseup', ()=>{angDrag=false;});
  dom.angIn.addEventListener('input', e=>{state.angle=(e.target.value%360+360)%360; updateVisuals();});

  // Copy
  dom.copy.addEventListener('click', () => {
    navigator.clipboard.writeText(dom.out.textContent).then(()=>{
      dom.copy.classList.add('copied');
      app.querySelector('.lg-copy-icon').style.display='none'; app.querySelector('.lg-check-icon').style.display='';
      setTimeout(()=>{
        dom.copy.classList.remove('copied');
        app.querySelector('.lg-copy-icon').style.display=''; app.querySelector('.lg-check-icon').style.display='none';
      }, 1500);
    });
  });

  render();
})();
