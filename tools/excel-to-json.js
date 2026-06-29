(function() {
  // Excel 轉 JSON 工具專屬 CSS (黑底白字主題)
  const css = `
    #excel-to-json-app {
      padding: 32px;
      color: var(--text-primary);
      font-family: var(--font-sans);
      max-width: 900px;
      margin: 0 auto;
      line-height: 1.6;
    }
    
    #excel-to-json-app h2 {
      color: var(--text-primary);
      font-size: 2rem;
      font-weight: 700;
      margin-bottom: 2rem;
      background: linear-gradient(to right, #60a5fa, #a78bfa);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      text-align: center;
    }
    
    #excel-to-json-app .upload-group {
      margin-bottom: 20px;
      padding: 20px;
      background: var(--bg-card);
      border-radius: var(--radius-md);
      border: 1px dashed var(--border-color);
      backdrop-filter: blur(12px);
    }
    
    #excel-to-json-app label {
      font-weight: bold;
      color: var(--text-primary);
      display: block;
      margin-bottom: 8px;
    }
    
    #excel-to-json-app input[type="file"] {
      width: 100%;
      cursor: pointer;
      background: var(--bg-secondary);
      border: 1px solid var(--border-color);
      border-radius: var(--radius-sm);
      padding: 12px;
      color: var(--text-primary);
      font-family: var(--font-sans);
      transition: var(--transition);
    }
    
    #excel-to-json-app input[type="file"]:hover {
      border-color: var(--accent);
      background: var(--bg-glass-hover);
    }
    
    #excel-to-json-app input[type="file"]::file-selector-button {
      background: var(--accent);
      color: white;
      border: none;
      border-radius: var(--radius-sm);
      padding: 8px 16px;
      margin-right: 12px;
      cursor: pointer;
      font-weight: 500;
      transition: var(--transition);
    }
    
    #excel-to-json-app input[type="file"]::file-selector-button:hover {
      background: var(--accent-2);
    }
    
    #excel-to-json-app .upload-group small {
      color: var(--text-secondary);
      display: block;
      margin-top: 8px;
      font-size: 0.9rem;
    }
    
    #excel-to-json-app .output-group {
      margin-top: 20px;
    }
    
    #excel-to-json-app .output-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 8px;
    }
    
    #excel-to-json-app .output-header label {
      margin-bottom: 0;
      color: var(--text-primary);
    }
    
    #excel-to-json-app .copy-btn {
      padding: 8px 16px;
      border: none;
      border-radius: var(--radius-sm);
      background: var(--accent);
      color: white;
      font-weight: 500;
      cursor: pointer;
      display: none;
      transition: var(--transition);
      font-family: var(--font-sans);
    }
    
    #excel-to-json-app .copy-btn:hover {
      background: var(--accent-2);
    }
    
    #excel-to-json-app .copy-btn.success {
      background: #10b981;
    }
    
    #excel-to-json-app pre {
      background: var(--bg-secondary);
      color: var(--text-primary);
      padding: 20px;
      border-radius: var(--radius-md);
      border: 1px solid var(--border-color);
      overflow-x: auto;
      max-height: 600px;
      font-family: var(--font-mono);
      font-size: 0.9rem;
      line-height: 1.5;
    }
    
    #excel-to-json-app .error {
      color: #ef4444;
      font-weight: bold;
    }

    #excel-to-json-app .json-html-group {
      margin-top: 28px;
      padding: 20px;
      background: var(--bg-card);
      border-radius: var(--radius-md);
      border: 1px solid var(--border-color);
    }

    #excel-to-json-app .json-html-group h3 {
      margin: 0 0 12px;
      font-size: 1.15rem;
      color: var(--text-primary);
    }

    #excel-to-json-app textarea {
      width: 100%;
      min-height: 220px;
      resize: vertical;
      background: var(--bg-secondary);
      border: 1px solid var(--border-color);
      border-radius: var(--radius-sm);
      padding: 12px;
      color: var(--text-primary);
      font-family: var(--font-mono);
      font-size: 0.88rem;
      line-height: 1.45;
    }

    #excel-to-json-app .action-row {
      margin-top: 12px;
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      align-items: center;
    }

    #excel-to-json-app .action-btn {
      padding: 8px 14px;
      border: 1px solid var(--border-color);
      border-radius: var(--radius-sm);
      background: var(--bg-secondary);
      color: var(--text-primary);
      font-family: var(--font-sans);
      cursor: pointer;
      transition: var(--transition);
    }

    #excel-to-json-app .action-btn:hover {
      border-color: var(--accent);
      background: var(--bg-glass-hover);
    }

    #excel-to-json-app .action-btn.primary {
      background: var(--accent);
      color: #fff;
      border-color: transparent;
    }

    #excel-to-json-app .action-btn.primary:hover {
      background: var(--accent-2);
    }

    #excel-to-json-app .status {
      margin-top: 10px;
      font-size: 0.92rem;
      color: var(--text-secondary);
    }

    #excel-to-json-app .status.ok {
      color: #10b981;
    }

    #excel-to-json-app .status.bad {
      color: #ef4444;
    }

    #excel-to-json-app .preview-wrap {
      margin-top: 16px;
    }

    #excel-to-json-app .preview-wrap label {
      margin-bottom: 8px;
    }

    #excel-to-json-app .html-preview {
      width: 100%;
      min-height: 280px;
      border: 1px solid var(--border-color);
      border-radius: var(--radius-md);
      background: #fff;
      padding: 14px;
      overflow-x: auto;
    }

    #excel-to-json-app .html-preview .comparison-toolbar {
      display: flex;
      justify-content: flex-start;
      align-items: flex-end;
      gap: 12px;
      margin-bottom: 14px;
    }

    #excel-to-json-app .html-preview .comparison-spec-spacer {
      flex: 0 0 28%;
      min-width: 220px;
    }

    #excel-to-json-app .html-preview .comparison-selectors {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      flex: 1;
    }

    #excel-to-json-app .html-preview .comparison-add-slot {
      flex: 0 0 auto;
      min-width: 180px;
    }

    #excel-to-json-app .html-preview .comparison-selector-item {
      min-width: 180px;
      flex: 1;
    }

    #excel-to-json-app .html-preview .comparison-selector-item label {
      display: block;
      margin-bottom: 6px;
      font-size: 12px;
      color: #6b7280;
    }

    #excel-to-json-app .html-preview .comparison-product-select {
      width: 100%;
      border: 1px solid #d1d5db;
      border-radius: 8px;
      padding: 8px;
      background: #fff;
      color: #111827;
    }

    #excel-to-json-app .html-preview .comparison-add-btn {
      border: 1px solid #d1d5db;
      background: #111827;
      color: #fff;
      border-radius: 8px;
      padding: 8px 12px;
      cursor: pointer;
      white-space: nowrap;
      height: 40px;
      width: 100%;
    }

    #excel-to-json-app .html-preview .comparison-add-btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }

    #excel-to-json-app .html-preview .comparison-category {
      border: 1px solid #e5e7eb;
      border-radius: 10px;
      margin-bottom: 12px;
      overflow: hidden;
    }

    #excel-to-json-app .html-preview .comparison-category > summary {
      cursor: pointer;
      list-style: none;
      padding: 12px 14px;
      background: #f9fafb;
      color: #111827;
      font-weight: 700;
    }

    #excel-to-json-app .html-preview .comparison-category > summary::-webkit-details-marker {
      display: none;
    }

    #excel-to-json-app .html-preview .comparison-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
      color: #111827;
      background: #fff;
    }

    #excel-to-json-app .html-preview .comparison-table th,
    #excel-to-json-app .html-preview .comparison-table td {
      border: 1px solid #e5e7eb;
      padding: 10px;
      vertical-align: top;
      text-align: left;
    }

    #excel-to-json-app .html-preview .comparison-table thead th {
      background: #f3f4f6;
    }

    #excel-to-json-app .html-preview .comparison-table tbody tr[data-hidden="true"] {
      opacity: 0.55;
    }

    #excel-to-json-app .html-preview .comparison-link-btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 6px 12px;
      border-radius: 999px;
      border: 1px solid #0ea5e9;
      color: #0ea5e9;
      text-decoration: none;
      font-size: 12px;
      font-weight: 600;
      white-space: nowrap;
    }

    #excel-to-json-app .html-preview .comparison-link-btn:hover {
      background: rgba(14, 165, 233, 0.1);
    }

    #excel-to-json-app .html-preview .comparison-link-group {
      display: flex;
      flex-direction: column;
      gap: 8px;
      align-items: flex-start;
    }

    #excel-to-json-app .html-preview .comparison-link-item {
      display: block;
    }

    #excel-to-json-app .html-preview .comparison-link-row-spacer {
      height: 38px;
    }

    #excel-to-json-app .html-preview .custom-select-wrap {
      position: relative;
      width: 100%;
    }

    #excel-to-json-app .html-preview .custom-select-trigger {
      width: 100%;
      background: #111827;
      color: #fff;
      border: 1px solid #374151;
      border-radius: 8px;
      padding: 8px 12px;
      cursor: pointer;
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 14px;
      text-align: left;
      font-family: inherit;
    }

    #excel-to-json-app .html-preview .custom-select-dropdown {
      position: absolute;
      top: calc(100% + 4px);
      left: 0;
      right: 0;
      background: #111827;
      border: 1px solid #374151;
      border-radius: 8px;
      list-style: none;
      padding: 4px 0;
      margin: 0;
      z-index: 100;
      display: none;
      box-shadow: 0 4px 16px rgba(0,0,0,0.4);
    }

    #excel-to-json-app .html-preview .custom-select-wrap.is-open .custom-select-dropdown {
      display: block;
    }

    #excel-to-json-app .html-preview .custom-option {
      padding: 10px 14px;
      cursor: pointer;
      color: #fff;
      font-size: 14px;
      list-style: none;
    }

    #excel-to-json-app .html-preview .custom-option:hover:not(.is-disabled):not(.is-selected) {
      color: #3b82f6;
    }

    #excel-to-json-app .html-preview .custom-option.is-selected {
      background: #2563eb;
      color: #fff;
    }

    #excel-to-json-app .html-preview .custom-option.is-disabled {
      color: #6b7280;
      cursor: default;
    }

    @media (max-width: 640px) {
      #excel-to-json-app {
        padding: 16px;
      }
      #excel-to-json-app h2 {
        font-size: 1.5rem;
      }
      #excel-to-json-app .upload-group {
        padding: 16px;
      }
    }
  `;

  // HTML 模板
  const html = `
    <div class="excel-to-json-container">
      <h2>ID192 : Table_Comparison - Excel 轉 JSON 工具</h2>
      
      <div class="upload-group">
        <label for="excel-upload">上傳規格表 (<a href="./public/excel-to-json/id192-demo.xlsx" target="_blank">.xlsx</a>)</label>
        <input type="file" id="excel-upload" accept=".xlsx, .xls, .csv">
        <small>
          * 第一列的 Hidden 與 Spec 為保留欄位, 不可修改, 後續列為產品型號可自行增減。<br>
          * 若 整條 SPEC 對應的數值皆留空，將自動判斷為分類標題（Category）。<br>
          * 可自行增減 SPEC 數量, 但要注意相同的 Hidden 設定要放在同一群, 穿插會造成未知的跳動問題。<br>
          * 請勿隨便更改第一排與第一列格式, 否則可能導致轉換失敗或資料錯亂。
        </small>
      </div>

      <div class="output-group" aria-live="polite">
        <div class="output-header">
          <label for="json-output">模組專用 JSON：</label>
          <button id="copy-btn" class="copy-btn">複製 JSON</button>
        </div>
        <pre id="json-output">等待上傳資料...</pre>
      </div>

      <div class="json-html-group">
        <h3>JSON 轉 HTML（Category 可收合 + Spec Table）</h3>
        <label for="json-input">貼上 JSON：</label>
        <textarea id="json-input" placeholder="請貼上 { products: [], rows: [] } JSON"></textarea>
        <div class="action-row">
          <button id="convert-json-btn" class="action-btn primary">驗證並轉換 HTML</button>
          <button id="copy-html-btn" class="action-btn" style="display:none;">複製 HTML</button>
        </div>
        <div id="json-status" class="status">等待貼上 JSON...</div>

        <div class="output-group" aria-live="polite">
          <div class="output-header">
            <label for="html-output">輸出 HTML：</label>
          </div>
          <pre id="html-output">尚未轉換</pre>
        </div>

        <div class="preview-wrap">
          <label for="html-preview">HTML 預覽：</label>
          <div id="html-preview" class="html-preview"></div>
        </div>
      </div>
    </div>
  `;

  // 載入 SheetJS 庫
  const script = document.createElement('script');
  script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
  document.head.appendChild(script);

  // 等待 SheetJS 載入完成後初始化工具
  script.onload = function() {
    // 注入 CSS 和 HTML 到掛載點
    const appContainer = document.getElementById('excel-to-json-app');
    const styleEl = document.createElement('style');
    styleEl.textContent = css;
    appContainer.appendChild(styleEl);
    appContainer.insertAdjacentHTML('beforeend', html);

    // 元素選取器（限制在工具容器內）
    const $ = selector => appContainer.querySelector(selector);

    /**
     * 將 SheetJS 產出的原始 JSON 轉換為模組渲染專用的結構
     */
    function formatSpecData(rawData) {
      if (!rawData || rawData.length === 0) return null;

      // 1. 取得所有欄位名稱 (Key)
      const allKeys = Object.keys(rawData[0]);
      
      // 【關鍵修正】：過濾掉 SheetJS 自動產生的 '__EMPTY' 幽靈欄位
      const cleanKeys = allKeys.filter(key => !key.startsWith('__EMPTY'));
      
      // 2. 第一個乾淨的欄位固定為規格標題 (例如 "Spec")
      const specColumnHidden = cleanKeys[0]; 
      const specColumnName = cleanKeys[1]; 
      
      // 3. 剩下欄位為產品欄，需過濾掉整欄都空白的無效欄位
      const allProductKeys = cleanKeys.slice(2);
      const productKeys = allProductKeys.filter((key) => {
        return rawData.some((row) => {
          const val = row[key];
          return val !== undefined && val !== null && String(val).trim() !== "";
        });
      });
      const productNames = productKeys.map((key, idx) => {
        const name = String(key ?? "").trim();
        return name || `Product ${idx + 1}`;
      });

      // 4. 迭代每一列資料
      const formattedRows = rawData.map((row) => {
        const specHidden = row[specColumnHidden]?.trim() || "";
        const specName = row[specColumnName]?.trim() || "";

        // 檢查所有產品欄位是否皆為空值
        const isCategory = productKeys.every(productKey => {
          const val = row[productKey];
          return val === undefined || val === null || String(val).trim() === "";
        });

        // 回傳對應結構
        if (isCategory) {
          // 防呆機制：如果是 Category 列，但 specName 本身也是空的，代表這是一整行純空白列，將其標記為 null 稍後濾除
          if (!specName) return null;

          return {
            type: "category",
            name: specName
          };
        } else {
          const specValues = productKeys.map(productKey => {
            const val = row[productKey];
            return (val !== undefined && val !== null && val !== "") ? String(val).trim() : "";
          });

          return {
            type: "spec",
            name: specName,
            hidden: specHidden,
            values: specValues
          };
        }
      }).filter(row => row !== null); // 【關鍵修正】：濾除完全空白的無效列

      return {
        products: productNames,
        rows: formattedRows
      };
    }

    function escapeHtml(value) {
      return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    function toCellHtml(value) {
      return escapeHtml(String(value ?? '').replace(/<br\s*\/?>(\s*)/gi, '\n')).replace(/\r?\n/g, '<br>');
    }

    function parseAemColorToken(value) {
      const raw = value !== undefined && value !== null ? String(value) : '';
      const match = raw.match(/^\s*\[(\d+)\]\s*/);
      if (!match) return { token: '', text: raw };
      return {
        token: match[1],
        text: raw.slice(match[0].length)
      };
    }

    function getAemTokenStyle(token) {
      // AEM 可覆寫這些值：可設為 "#7069F6" 或 "linear-gradient(...)"
      const tokenStyles = {
        '1': '#7069F6',
        '2': '#16a34a',
        '3': 'linear-gradient(90deg, rgb(139, 92, 246) 0%, rgb(59, 130, 246) 100%)'
      };

      return tokenStyles[String(token || '')] || '';
    }

    function resolveAemTokenStyle(tokenInfo) {
      const token = String(tokenInfo?.token || '').trim();
      const styleValue = getAemTokenStyle(token);
      if (!token || !styleValue) {
        return { token: '', type: 'none', value: '' };
      }

      const isGradient = /^linear-gradient\s*\(/i.test(styleValue);
      return {
        token,
        type: isGradient ? 'gradient' : 'solid',
        value: styleValue
      };
    }

    function applyAemTokenColor(html, tokenInfo, fallbackColor = '') {
      const resolved = resolveAemTokenStyle(tokenInfo);
      if (resolved.type === 'gradient') {
        const className = `aem-color-token aem-gradient-token aem-color-${escapeHtml(resolved.token)}`;
        const style = `background-image:${escapeHtml(resolved.value)};-webkit-background-clip:text;background-clip:text;color:transparent;-webkit-text-fill-color:transparent;`;
        return `<span class="${className}" data-aem-color-token="${escapeHtml(resolved.token)}" data-aem-style-type="gradient" style="${style}">${html}</span>`;
      }

      if (resolved.type === 'solid') {
        const className = `aem-color-token aem-color-${escapeHtml(resolved.token)}`;
        return `<span class="${className}" data-aem-color-token="${escapeHtml(resolved.token)}" data-aem-style-type="solid" style="color:${escapeHtml(resolved.value)};">${html}</span>`;
      }

      const finalColor = fallbackColor || '';
      if (!finalColor) return html;

      return `<span class="aem-color-token" style="color:${escapeHtml(finalColor)};">${html}</span>`;
    }

    function getAemLinkStyle(tokenInfo, fallbackColor = '') {
      const resolved = resolveAemTokenStyle(tokenInfo);
      if (resolved.type === 'gradient') {
        return {
          classPart: ` aem-color-token aem-gradient-token aem-color-${escapeHtml(resolved.token)}`,
          dataAttrs: ` data-aem-color-token="${escapeHtml(resolved.token)}" data-aem-style-type="gradient"`,
          stylePart: ` style="color:#fff; border-color:transparent; background-image:${escapeHtml(resolved.value)};"`
        };
      }

      if (resolved.type === 'solid') {
        return {
          classPart: ` aem-color-token aem-color-${escapeHtml(resolved.token)}`,
          dataAttrs: ` data-aem-color-token="${escapeHtml(resolved.token)}" data-aem-style-type="solid"`,
          stylePart: ` style="color:${escapeHtml(resolved.value)}; border-color:${escapeHtml(resolved.value)};"`
        };
      }

      const stylePart = fallbackColor
        ? ` style="color:${escapeHtml(fallbackColor)}; border-color:${escapeHtml(fallbackColor)};"`
        : '';
      return {
        classPart: '',
        dataAttrs: '',
        stylePart
      };
    }

    function applyTextColor(html, color, tokenInfo = null) {
      return applyAemTokenColor(html, tokenInfo, color);
    }

    function isLinkSpecName(specName) {
      return String(specName || '').trim().startsWith('/');
    }

    function getDisplaySpecName(specName) {
      return String(specName || '').trim().replace(/^\/+/, '') || '-';
    }

    function normalizeLinkUrl(rawValue) {
      const safeValue = rawValue !== undefined && rawValue !== null ? String(rawValue).trim() : '';
      if (!safeValue) return '';
      const matched = safeValue.match(/https?:\/\/\S+/i);
      return matched ? matched[0] : safeValue;
    }

    function getWorkbookFileText(workbook, candidatePaths) {
      const files = workbook?.files;
      if (!files) return '';

      const decoder = new TextDecoder('utf-8');

      function toText(entry) {
        if (!entry) return '';
        if (typeof entry === 'string') return entry;

        const payload = entry.content ?? entry.data ?? entry._data ?? null;
        if (typeof payload === 'string') return payload;

        if (payload instanceof ArrayBuffer) {
          return decoder.decode(new Uint8Array(payload));
        }
        if (payload instanceof Uint8Array) {
          return decoder.decode(payload);
        }
        if (Array.isArray(payload)) {
          return decoder.decode(new Uint8Array(payload));
        }

        if (typeof entry.asText === 'function') {
          try {
            return entry.asText();
          } catch (error) {
            return '';
          }
        }

        return '';
      }

      for (const path of candidatePaths) {
        if (files[path]) {
          const text = toText(files[path]);
          if (text) return text;
        }
      }

      const fileKeys = Object.keys(files);
      for (const path of candidatePaths) {
        const matchedKey = fileKeys.find((key) => key.endsWith(path));
        if (matchedKey) {
          const text = toText(files[matchedKey]);
          if (text) return text;
        }
      }

      return '';
    }

    function parseXml(xmlText) {
      if (!xmlText) return null;
      try {
        return new DOMParser().parseFromString(xmlText, 'application/xml');
      } catch (error) {
        return null;
      }
    }

    function buildWorkbookColorLookup(workbook, worksheet) {
      const styleMap = {
        styleIdToColor: {},
        addressToColor: {}
      };

      const stylesXml = getWorkbookFileText(workbook, ['xl/styles.xml', 'styles.xml']);
      const stylesDoc = parseXml(stylesXml);
      if (!stylesDoc) return styleMap;

      const fontColors = Array.from(stylesDoc.querySelectorAll('fonts > font')).map((fontNode) => {
        const colorNode = fontNode.querySelector('color');
        const rgb = colorNode?.getAttribute('rgb') || colorNode?.getAttribute('argb') || '';
        return normalizeRgbToHex(rgb);
      });

      Array.from(stylesDoc.querySelectorAll('cellXfs > xf')).forEach((xfNode, styleId) => {
        const fontIdRaw = xfNode.getAttribute('fontId');
        const fontId = fontIdRaw !== null ? Number(fontIdRaw) : NaN;
        if (!Number.isFinite(fontId)) return;
        const color = fontColors[fontId] || '';
        if (color) styleMap.styleIdToColor[styleId] = color;
      });

      const worksheetRef = worksheet?.['!ref'] || '';
      const files = workbook?.files ? Object.keys(workbook.files) : [];
      const sheetCandidates = files
        .filter((key) => /xl\/worksheets\/sheet\d+\.xml$/i.test(key))
        .sort();

      let targetSheetXml = '';
      for (const key of sheetCandidates) {
        const xml = getWorkbookFileText(workbook, [key]);
        if (!xml) continue;
        if (!worksheetRef) {
          targetSheetXml = xml;
          break;
        }
        const doc = parseXml(xml);
        const dimRef = doc?.querySelector('dimension')?.getAttribute('ref') || '';
        if (dimRef === worksheetRef) {
          targetSheetXml = xml;
          break;
        }
      }

      if (!targetSheetXml && sheetCandidates.length > 0) {
        targetSheetXml = getWorkbookFileText(workbook, [sheetCandidates[0]]);
      }

      const sheetDoc = parseXml(targetSheetXml);
      if (!sheetDoc) return styleMap;

      Array.from(sheetDoc.querySelectorAll('sheetData c[s][r]')).forEach((cellNode) => {
        const address = cellNode.getAttribute('r') || '';
        const styleId = Number(cellNode.getAttribute('s'));
        if (!address || !Number.isFinite(styleId)) return;
        const color = styleMap.styleIdToColor[styleId] || '';
        if (color) styleMap.addressToColor[address] = color;
      });

      return styleMap;
    }

    function extractColorFromFont(font) {
      const color = font?.color;
      if (!color) return '';

      const byRgb = normalizeRgbToHex(color.rgb || color.argb);
      if (byRgb) return byRgb;

      // theme/indexed 顏色在不同檔案常會還原成預設色，這裡先保留空字串避免錯色
      return '';
    }

    function getStyleArrays(workbook) {
      const styles = workbook?.Styles || workbook?.styles || {};
      return {
        cellXfs: styles.CellXf || styles.CellXF || styles.CellXfs || styles.cellXfs || [],
        fonts: styles.Fonts || styles.fonts || []
      };
    }

    function extractFontColor(cell, workbook, address = '') {
      // Case 1: cell.s 已是完整樣式物件
      const directColor = extractColorFromFont(cell?.s?.font);
      if (directColor) return directColor;

      // Case 1.5: cell.s 是樣式物件，但只帶 fontId
      const inlineFontId = cell?.s?.fontId ?? cell?.s?.FontId;
      if (inlineFontId !== undefined && inlineFontId !== null) {
        const { fonts } = getStyleArrays(workbook);
        const font = fonts[Number(inlineFontId)];
        const byInlineFontId = extractColorFromFont(font);
        if (byInlineFontId) return byInlineFontId;
      }

      // Case 2: cell.s 是樣式索引，需回查 workbook Styles
      const styleId = typeof cell?.s === 'number' || typeof cell?.s === 'string'
        ? Number(cell.s)
        : null;
      if (styleId !== null && Number.isFinite(styleId) && styleId >= 0) {
        const { cellXfs, fonts } = getStyleArrays(workbook);
        const xf = cellXfs[styleId];
        const fontId = xf?.fontId ?? xf?.FontId;
        if (fontId !== undefined && fontId !== null) {
          const font = fonts[Number(fontId)];
          const styleColor = extractColorFromFont(font);
          if (styleColor) return styleColor;
        }
      }

      // Case 2.5: 直接從 XML 解析到的 address 色彩
      if (address && workbook?.__addressToColor && workbook.__addressToColor[address]) {
        return workbook.__addressToColor[address];
      }

      // Fallback: 某些檔案在樣式解析時只保留 HTML rich text
      const html = String(cell?.h || '');
      const m = html.match(/color\s*:\s*(#[0-9a-fA-F]{3,8}|rgb\([^\)]+\))/i);
      return m ? m[1] : '';
    }

    function renderPreviewCellHtml(specName, rawValue) {
      const parsedValue = parseAemColorToken(rawValue);
      const safeValue = String(parsedValue.text || '').trim();
      if (!isLinkSpecName(specName)) {
        return applyTextColor(toCellHtml(safeValue), '', parsedValue);
      }

      if (!safeValue) return '';

      const safeUrl = escapeHtml(normalizeLinkUrl(safeValue));
      const safeLabel = escapeHtml(getDisplaySpecName(specName));
      const linkStyle = getAemLinkStyle(parsedValue, '');
      return `<a class="comparison-link-btn${linkStyle.classPart}"${linkStyle.dataAttrs}${linkStyle.stylePart} href="${safeUrl}" target="_blank" rel="noopener noreferrer">${safeLabel}</a>`;
    }

    function isTrueFlag(value) {
      return String(value ?? '').trim().toUpperCase() === 'TRUE';
    }

    // isTrueFlag 也需在 output script 中定義
    // (已在 convertJsonToHtml 的 script 字串中包含)

    function validateComparisonJson(data) {
      const errors = [];

      if (!data || typeof data !== 'object') {
        errors.push('根節點必須是物件。');
        return errors;
      }

      if (!Array.isArray(data.products) || data.products.length === 0) {
        errors.push('products 必須是非空陣列。');
      } else if (data.products.some(p => typeof p !== 'string')) {
        errors.push('products 內每一項都必須是字串。');
      }

      if (!Array.isArray(data.rows) || data.rows.length === 0) {
        errors.push('rows 必須是非空陣列。');
        return errors;
      }

      data.rows.forEach((row, index) => {
        const rowNo = index + 1;
        if (!row || typeof row !== 'object') {
          errors.push(`rows[${rowNo}] 必須是物件。`);
          return;
        }

        if (row.type !== 'category' && row.type !== 'spec') {
          errors.push(`rows[${rowNo}] 的 type 只能是 category 或 spec。`);
          return;
        }

        if (typeof row.name !== 'string' || !row.name.trim()) {
          errors.push(`rows[${rowNo}] 的 name 必須是非空字串。`);
        }

        if (row.type === 'spec') {
          if (!Array.isArray(row.values)) {
            errors.push(`rows[${rowNo}] 是 spec 時，values 必須是陣列。`);
            return;
          }

          if (Array.isArray(data.products) && row.values.length !== data.products.length) {
            errors.push(`rows[${rowNo}] 的 values 長度 (${row.values.length}) 必須等於 products 長度 (${data.products.length})。`);
          }
        }
      });

      return errors;
    }

    function convertJsonToHtml(data) {
      const sections = buildComparisonSections(data.rows);

      const inlineData = JSON.stringify({
        products: data.products,
        sections
      }).replace(/</g, '\\u003c');

      return [
        '<style>',
        '.comparison-module .custom-select-wrap { position: relative; width: 100%; }',
        '.comparison-module .custom-select-trigger { width: 100%; background: #111827; color: #fff; border: 1px solid #374151; border-radius: 8px; padding: 8px 12px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; font-size: 14px; text-align: left; font-family: inherit; }',
        '.comparison-module .custom-select-dropdown { position: absolute; top: calc(100% + 4px); left: 0; right: 0; background: #111827; border: 1px solid #374151; border-radius: 8px; list-style: none; padding: 4px 0; margin: 0; z-index: 100; display: none; box-shadow: 0 4px 16px rgba(0,0,0,0.4); }',
        '.comparison-module .custom-select-wrap.is-open .custom-select-dropdown { display: block; }',
        '.comparison-module .custom-option { padding: 10px 14px; cursor: pointer; color: #fff; font-size: 14px; list-style: none; }',
        '.comparison-module .custom-option:hover:not(.is-disabled):not(.is-selected) { color: #3b82f6; }',
        '.comparison-module .custom-option.is-selected { background: #2563eb; color: #fff; }',
        '.comparison-module .custom-option.is-disabled { color: #6b7280; cursor: default; }',
        '</style>',
        '<section class="comparison-module">',
        '  <div class="comparison-toolbar">',
        '    <div class="comparison-spec-spacer" aria-hidden="true"></div>',
        '    <div class="comparison-selectors" aria-label="產品選擇"></div>',
        '    <div class="comparison-add-slot"></div>',
        '  </div>',
        '  <div class="comparison-sections"></div>',
        '</section>',
        '<script>',
        '(function () {',
        '  const root = document.currentScript.previousElementSibling;',
        `  const data = ${inlineData};`,
        '  const selectorsEl = root.querySelector(".comparison-selectors");',
        '  const sectionsEl = root.querySelector(".comparison-sections");',
        '  const addSlotEl = root.querySelector(".comparison-add-slot");',
        '  const maxVisible = Math.min(4, data.products.length);',
        '  let visible = data.products.length > 1 ? [0, 1] : [0];',
        '',
        '  function escapeHtml(text) {',
        '    return String(text)',
        '      .replace(/&/g, "&amp;")',
        '      .replace(/</g, "&lt;")',
        '      .replace(/>/g, "&gt;")',
        '      .replace(/\"/g, "&quot;")',
        '      .replace(/\'/g, "&#39;");',
        '  }',
        '',
        '  function getNextProductIndex() {',
        '    for (let i = 0; i < data.products.length; i += 1) {',
        '      if (!visible.includes(i)) return i;',
        '    }',
        '    return -1;',
        '  }',
        '',
        '  function isTrueFlag(value) {',
        '    return String(value !== null && value !== undefined ? value : "").trim().toUpperCase() === "TRUE";',
        '  }',
        '',
        '  function isLinkSpecName(specName) {',
        '    return String(specName || "").trim().startsWith("/");',
        '  }',
        '',
        '  function getDisplaySpecName(specName) {',
        '    return String(specName || "").trim().replace(/^\\/+/, "") || "-";',
        '  }',
        '',
        '  function normalizeLinkUrl(rawValue) {',
        '    const safeValue = rawValue !== undefined && rawValue !== null ? String(rawValue).trim() : "";',
        '    if (!safeValue) return "";',
        '    const matched = safeValue.match(/https?:\\/\\/\\S+/i);',
        '    return matched ? matched[0] : safeValue;',
        '  }',
        '',
        '  function parseAemColorToken(value) {',
        '    const raw = value !== undefined && value !== null ? String(value) : "";',
        '    const match = raw.match(/^\\s*\\[(\\d+)\\]\\s*/);',
        '    if (!match) return { token: "", text: raw };',
        '    return { token: match[1], text: raw.slice(match[0].length) };',
        '  }',
        '',
        '  function getAemTokenStyle(token) {',
        '    const tokenStyles = {',
        '      "1": "#7069F6",',
        '      "2": "#16a34a",',
        '      "3": "linear-gradient(90deg, rgb(139, 92, 246) 0%, rgb(59, 130, 246) 100%)"',
        '    };',
        '    return tokenStyles[String(token || "")] || "";',
        '  }',
        '',
        '  function resolveAemTokenStyle(tokenInfo) {',
        '    const token = String((tokenInfo && tokenInfo.token) || "").trim();',
        '    const styleValue = getAemTokenStyle(token);',
        '    if (!token || !styleValue) return { token: "", type: "none", value: "" };',
        '    const isGradient = /^linear-gradient\\s*\\(/i.test(styleValue);',
        '    return { token, type: isGradient ? "gradient" : "solid", value: styleValue };',
        '  }',
        '',
        '  function applyAemTokenColor(html, tokenInfo, fallbackColor = "") {',
        '    const resolved = resolveAemTokenStyle(tokenInfo);',
        '    if (resolved.type === "gradient") {',
        '      const className = `aem-color-token aem-gradient-token aem-color-${escapeHtml(resolved.token)}`;',
        '      const style = `background-image:${escapeHtml(resolved.value)};-webkit-background-clip:text;background-clip:text;color:transparent;-webkit-text-fill-color:transparent;`;',
        '      return `<span class="${className}" data-aem-color-token="${escapeHtml(resolved.token)}" data-aem-style-type="gradient" style="${style}">${html}</span>`;',
        '    }',
        '    if (resolved.type === "solid") {',
        '      const className = `aem-color-token aem-color-${escapeHtml(resolved.token)}`;',
        '      return `<span class="${className}" data-aem-color-token="${escapeHtml(resolved.token)}" data-aem-style-type="solid" style="color:${escapeHtml(resolved.value)};">${html}</span>`;',
        '    }',
        '    const finalColor = fallbackColor || "";',
        '    if (!finalColor) return html;',
        '    return `<span class="aem-color-token" style="color:${escapeHtml(finalColor)};">${html}</span>`;',
        '  }',
        '',
        '  function getAemLinkStyle(tokenInfo, fallbackColor = "") {',
        '    const resolved = resolveAemTokenStyle(tokenInfo);',
        '    if (resolved.type === "gradient") {',
        '      return {',
        '        classPart: ` aem-color-token aem-gradient-token aem-color-${escapeHtml(resolved.token)}`,',
        '        dataAttrs: ` data-aem-color-token="${escapeHtml(resolved.token)}" data-aem-style-type="gradient"`,',
        '        stylePart: ` style="color:#fff; border-color:transparent; background-image:${escapeHtml(resolved.value)};"`',
        '      };',
        '    }',
        '    if (resolved.type === "solid") {',
        '      return {',
        '        classPart: ` aem-color-token aem-color-${escapeHtml(resolved.token)}`,',
        '        dataAttrs: ` data-aem-color-token="${escapeHtml(resolved.token)}" data-aem-style-type="solid"`,',
        '        stylePart: ` style="color:${escapeHtml(resolved.value)}; border-color:${escapeHtml(resolved.value)};"`',
        '      };',
        '    }',
        '    const stylePart = fallbackColor',
        '      ? ` style="color:${escapeHtml(fallbackColor)}; border-color:${escapeHtml(fallbackColor)};"`',
        '      : "";',
        '    return { classPart: "", dataAttrs: "", stylePart };',
        '  }',
        '',
        '  function renderCellHtml(specName, rawValue, color = "") {',
        '    const parsedValue = parseAemColorToken(rawValue);',
        '    const safeValue = String(parsedValue.text || "").trim();',
        '    if (!isLinkSpecName(specName)) {',
        '      const html = escapeHtml(safeValue.replace(/<br\\s*\\/?>(\\s*)/gi, "\\n")).replace(/\\r?\\n/g, "<br>");',
        '      return applyAemTokenColor(html, parsedValue, color);',
        '    }',
        '    if (!safeValue) return "";',
        '    const safeUrl = escapeHtml(normalizeLinkUrl(safeValue));',
        '    const safeLabel = escapeHtml(getDisplaySpecName(specName));',
        '    const linkStyle = getAemLinkStyle(parsedValue, color);',
        '    return `<a class="comparison-link-btn${linkStyle.classPart}"${linkStyle.dataAttrs}${linkStyle.stylePart} href="${safeUrl}" target="_blank" rel="noopener noreferrer">${safeLabel}</a>`;',
        '  }',
        '',
        '  function applyTextColor(html, color, tokenInfo = null) {',
        '    return applyAemTokenColor(html, tokenInfo, color);',
        '  }',
        '',
        '  function buildSelectors() {',
        '    selectorsEl.innerHTML = "";',
        '    visible.forEach((productIdx, colIdx) => {',
        '      const wrap = document.createElement("div");',
        '      wrap.className = "comparison-selector-item";',
        '',
        '      const label = document.createElement("label");',
        '      label.textContent = `Product ${colIdx + 1}`;',
        '',
        '      const selectWrap = document.createElement("div");',
        '      selectWrap.className = "custom-select-wrap";',
        '',
        '      const trigger = document.createElement("button");',
        '      trigger.type = "button";',
        '      trigger.className = "custom-select-trigger";',
        '      const triggerText = document.createElement("span");',
        '      triggerText.textContent = data.products[productIdx];',
        '      const triggerArrow = document.createElement("span");',
        '      triggerArrow.textContent = "⌄";',
        '      trigger.appendChild(triggerText);',
        '      trigger.appendChild(triggerArrow);',
        '',
        '      const dropdown = document.createElement("ul");',
        '      dropdown.className = "custom-select-dropdown";',
        '',
        '      data.products.forEach((name, idx) => {',
        '        const li = document.createElement("li");',
        '        li.className = "custom-option";',
        '        li.textContent = name;',
        '',
        '        if (idx === productIdx) {',
        '          li.classList.add("is-selected");',
        '        } else if (visible.includes(idx)) {',
        '          li.classList.add("is-disabled");',
        '        } else {',
        '          li.addEventListener("click", () => {',
        '            visible[colIdx] = idx;',
        '            selectWrap.classList.remove("is-open");',
        '            buildSelectors();',
        '            buildTables();',
        '          });',
        '        }',
        '',
        '        dropdown.appendChild(li);',
        '      });',
        '',
        '      trigger.addEventListener("click", (e) => {',
        '        e.stopPropagation();',
        '        const isOpen = selectWrap.classList.contains("is-open");',
        '        selectorsEl.querySelectorAll(".custom-select-wrap.is-open").forEach((el) => el.classList.remove("is-open"));',
        '        if (!isOpen) {',
        '          selectWrap.classList.add("is-open");',
        '        }',
        '      });',
        '',
        '      selectWrap.appendChild(trigger);',
        '      selectWrap.appendChild(dropdown);',
        '      wrap.appendChild(label);',
        '      wrap.appendChild(selectWrap);',
        '      selectorsEl.appendChild(wrap);',
        '    });',
        '    renderAddControl();',
        '  }',

        '  function renderAddControl() {',
        '    addSlotEl.innerHTML = "";',
        '    const nextIndices = data.products',
        '      .map((_, idx) => idx)',
        '      .filter((idx) => !visible.includes(idx));',
        '',
        '    if (visible.length >= maxVisible || nextIndices.length === 0) {',
        '      return;',
        '    }',
        '',
        '    if (visible.length === maxVisible - 1) {',
        '      const select = document.createElement("select");',
        '      select.className = "comparison-add-btn";',
        '',
        '      const placeholder = document.createElement("option");',
        '      placeholder.value = "";',
        '      placeholder.textContent = "Select 4th Product";',
        '      placeholder.selected = true;',
        '      placeholder.disabled = true;',
        '      select.appendChild(placeholder);',
        '',
        '      nextIndices.forEach((idx) => {',
        '        const option = document.createElement("option");',
        '        option.value = String(idx);',
        '        option.textContent = data.products[idx];',
        '        select.appendChild(option);',
        '      });',
        '',
        '      select.addEventListener("change", (event) => {',
        '        const nextIdx = Number(event.target.value);',
        '        if (Number.isNaN(nextIdx)) return;',
        '        visible.push(nextIdx);',
        '        buildSelectors();',
        '        buildTables();',
        '      });',
        '',
        '      addSlotEl.appendChild(select);',
        '      return;',
        '    }',
        '',
        '    const addBtn = document.createElement("button");',
        '    addBtn.type = "button";',
        '    addBtn.className = "comparison-add-btn";',
        '    addBtn.textContent = "Add Product";',
        '    addBtn.addEventListener("click", () => {',
        '      const next = getNextProductIndex();',
        '      if (next === -1) return;',
        '      visible.push(next);',
        '      buildSelectors();',
        '      buildTables();',
        '    });',
        '    addSlotEl.appendChild(addBtn);',
        '  }',
        '',
        '  function buildTables() {',
        '    const valueWidth = visible.length > 0 ? (72 / visible.length).toFixed(4) : 72;',
        '    const colgroupHtml = [',
        '      `<col style="width: 28%;">`,',
        '      ...visible.map(() => `<col style="width: ${valueWidth}%;">`)',
        '    ].join("");',
        '',
        '    const sectionsHtml = data.sections.map((section) => {',
        '      const rows = [];',
        '      for (let i = 0; i < section.specs.length; i += 1) {',
        '        const spec = section.specs[i];',
        '        const hiddenAttr = String(spec.hidden || "").toUpperCase() === "TRUE" ? " data-hidden=\"true\"" : "";',
        '',
        '        if (isLinkSpecName(spec.name)) {',
        '          const linkGroup = [spec];',
        '          let j = i + 1;',
        '          while (j < section.specs.length && isLinkSpecName(section.specs[j].name)) {',
        '            linkGroup.push(section.specs[j]);',
        '            j += 1;',
        '          }',
        '',
        '          const span = linkGroup.length;',
        '          const linkCols = visible',
        '            .map((idx) => {',
        '              const buttons = linkGroup',
        '                .map((linkSpec) => {',
        '                  const rawValue = linkSpec.values && linkSpec.values[idx] !== undefined ? linkSpec.values[idx] : "";',
        '                  const anchor = renderCellHtml(linkSpec.name, rawValue, "");',
        '                  return anchor ? `<div class="comparison-link-item">${anchor}</div>` : "";',
        '                })',
        '                .filter(Boolean)',
        '                .join("");',
        '',
        '              return `<td rowspan="${span}"><div class="comparison-link-group">${buttons}</div></td>`;',
        '            })',
        '            .join("");',
        '',
        '          rows.push(`<tr${hiddenAttr}><th scope="row" rowspan="${span}"></th>${linkCols}</tr>`);',
        '          for (let r = 1; r < span; r += 1) {',
        '            rows.push("<tr class=\"comparison-link-row-spacer\"></tr>");',
        '          }',
        '',
        '          i = j - 1;',
        '          continue;',
        '        }',
        '',
        '        const valueCols = visible',
        '          .map((idx) => {',
        '            const rawValue = spec.values && spec.values[idx] !== undefined ? spec.values[idx] : "";',
        '            return `<td>${renderCellHtml(spec.name, rawValue, "")}</td>`;',
        '          })',
        '          .join("");',
        '',
        '        const parsedSpecName = parseAemColorToken(spec.name);',
        '        const cleanSpecName = getDisplaySpecName(parsedSpecName.text);',
        '        const headerText = applyTextColor(escapeHtml(cleanSpecName), "", parsedSpecName);',
        '        rows.push(`<tr${hiddenAttr}><th scope="row">${headerText}</th>${valueCols}</tr>`);',
        '      }',
        '',
        '      const rowsHtml = rows.join("");',
        '',
        '      return [',
        '        "<details class=\"comparison-category\" open>",',
        '        `  <summary>${escapeHtml(section.name)}</summary>`,',
        '        "  <table class=\"comparison-table\">",',
        '        `    <colgroup>${colgroupHtml}</colgroup>`,',
        '        `    <tbody>${rowsHtml}</tbody>`,',
        '        "  </table>",',
        '        "</details>"',
        '      ].join("\\n");',
        '    }).join("\\n\\n");',
        '',
        '    sectionsEl.innerHTML = sectionsHtml;',
        '  }',
        '',
        '  document.addEventListener("click", () => {',
        '    root.querySelectorAll(".custom-select-wrap.is-open").forEach((el) => el.classList.remove("is-open"));',
        '  });',
        '',
        '  buildSelectors();',
        '  buildTables();',
        '})();',
        '</script>'
      ].join('\n');
    }

    function buildComparisonSections(rows) {
      const sections = [];
      let current = null;

      rows.forEach((row) => {
        if (row.type === 'category') {
          current = { name: row.name, specs: [] };
          sections.push(current);
          return;
        }

        if (!current) {
          current = { name: 'General', specs: [] };
          sections.push(current);
        }

        current.specs.push({ ...row });
      });

      return sections;
    }

    function renderInteractivePreview(data, mountEl) {
      const sections = buildComparisonSections(data.rows);
      let visible = data.products.length > 1 ? [0, 1] : [0];

      mountEl.innerHTML = [
        '<section class="comparison-module">',
        '  <div class="comparison-toolbar">',
        '    <div class="comparison-spec-spacer" aria-hidden="true"></div>',
        '    <div class="comparison-selectors" aria-label="產品選擇"></div>',
        '    <div class="comparison-add-slot"></div>',
        '  </div>',
        '  <div class="comparison-sections"></div>',
        '</section>'
      ].join('');

      if (!mountEl._hasClickOutside) {
        mountEl._hasClickOutside = true;
        document.addEventListener('click', () => {
          mountEl.querySelectorAll('.custom-select-wrap.is-open').forEach((el) => el.classList.remove('is-open'));
        });
      }

      const selectorsEl = mountEl.querySelector('.comparison-selectors');
      const sectionsEl = mountEl.querySelector('.comparison-sections');
      const addSlotEl = mountEl.querySelector('.comparison-add-slot');
      const maxVisible = Math.min(4, data.products.length);

      function getNextProductIndex() {
        for (let i = 0; i < data.products.length; i += 1) {
          if (!visible.includes(i)) return i;
        }
        return -1;
      }

      function buildSelectors() {
        selectorsEl.innerHTML = '';

        visible.forEach((productIdx, colIdx) => {
          const wrap = document.createElement('div');
          wrap.className = 'comparison-selector-item';

          const label = document.createElement('label');
          label.textContent = `Product ${colIdx + 1}`;

          const selectWrap = document.createElement('div');
          selectWrap.className = 'custom-select-wrap';

          const trigger = document.createElement('button');
          trigger.type = 'button';
          trigger.className = 'custom-select-trigger';
          const triggerText = document.createElement('span');
          triggerText.textContent = data.products[productIdx];
          const triggerArrow = document.createElement('span');
          triggerArrow.textContent = '⌄';
          trigger.appendChild(triggerText);
          trigger.appendChild(triggerArrow);

          const dropdown = document.createElement('ul');
          dropdown.className = 'custom-select-dropdown';

          data.products.forEach((name, idx) => {
            const li = document.createElement('li');
            li.className = 'custom-option';
            li.textContent = name;

            if (idx === productIdx) {
              li.classList.add('is-selected');
            } else if (visible.includes(idx)) {
              li.classList.add('is-disabled');
            } else {
              li.addEventListener('click', () => {
                visible[colIdx] = idx;
                selectWrap.classList.remove('is-open');
                buildSelectors();
                buildTables();
              });
            }

            dropdown.appendChild(li);
          });

          trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            const isOpen = selectWrap.classList.contains('is-open');
            selectorsEl.querySelectorAll('.custom-select-wrap.is-open').forEach((el) => el.classList.remove('is-open'));
            if (!isOpen) {
              selectWrap.classList.add('is-open');
            }
          });

          selectWrap.appendChild(trigger);
          selectWrap.appendChild(dropdown);
          wrap.appendChild(label);
          wrap.appendChild(selectWrap);
          selectorsEl.appendChild(wrap);
        });

        renderAddControl();
      }

      function renderAddControl() {
        addSlotEl.innerHTML = '';
        const nextIndices = data.products
          .map((_, idx) => idx)
          .filter((idx) => !visible.includes(idx));

        if (visible.length >= maxVisible || nextIndices.length === 0) {
          return;
        }

        if (visible.length === maxVisible - 1) {
          const select = document.createElement('select');
          select.className = 'comparison-add-btn';

          const placeholder = document.createElement('option');
          placeholder.value = '';
          placeholder.textContent = 'Select 4th Product';
          placeholder.selected = true;
          placeholder.disabled = true;
          select.appendChild(placeholder);

          nextIndices.forEach((idx) => {
            const option = document.createElement('option');
            option.value = String(idx);
            option.textContent = data.products[idx];
            select.appendChild(option);
          });

          select.addEventListener('change', (event) => {
            const nextIdx = Number(event.target.value);
            if (Number.isNaN(nextIdx)) return;
            visible.push(nextIdx);
            buildSelectors();
            buildTables();
          });

          addSlotEl.appendChild(select);
          return;
        }

        const addBtn = document.createElement('button');
        addBtn.type = 'button';
        addBtn.className = 'comparison-add-btn';
        addBtn.textContent = 'Add Product';
        addBtn.addEventListener('click', () => {
          const next = getNextProductIndex();
          if (next === -1) return;
          visible.push(next);
          buildSelectors();
          buildTables();
        });
        addSlotEl.appendChild(addBtn);
      }

      function buildTables() {
        const valueWidth = visible.length > 0 ? (72 / visible.length).toFixed(4) : 72;
        const colgroupHtml = [
          `<col style="width: 28%;">`,
          ...visible.map(() => `<col style="width: ${valueWidth}%;">`)
        ].join('');

        const sectionHtml = sections.map((section) => {
          const rows = [];
          for (let i = 0; i < section.specs.length; i += 1) {
            const spec = section.specs[i];
            if (isLinkSpecName(spec.name)) {
              const allLinkSpecs = [spec];
              let j = i + 1;
              while (j < section.specs.length && isLinkSpecName(section.specs[j].name)) {
                allLinkSpecs.push(section.specs[j]);
                j += 1;
              }
              i = j - 1;

              const linkGroup = allLinkSpecs.filter(s => !isTrueFlag(s.hidden));
              if (linkGroup.length === 0) continue;

              const span = linkGroup.length;
              const linkCols = visible
                .map((idx) => {
                  const buttons = linkGroup
                    .map((linkSpec) => {
                      const rawValue = linkSpec.values && linkSpec.values[idx] !== undefined ? linkSpec.values[idx] : '';
                      const anchor = renderPreviewCellHtml(linkSpec.name, rawValue);
                      return anchor ? `<div class="comparison-link-item">${anchor}</div>` : '';
                    })
                    .filter(Boolean)
                    .join('');

                  return `<td rowspan="${span}"><div class="comparison-link-group">${buttons}</div></td>`;
                })
                .join('');

              rows.push(`<tr><th scope="row" rowspan="${span}"></th>${linkCols}</tr>`);
              for (let r = 1; r < span; r += 1) {
                rows.push('<tr class="comparison-link-row-spacer"></tr>');
              }

              continue;
            }

            const valueCols = visible
              .map((idx) => {
                const rawValue = spec.values && spec.values[idx] !== undefined ? spec.values[idx] : '';
                return `<td>${renderPreviewCellHtml(spec.name, rawValue)}</td>`;
              })
              .join('');

            const parsedSpecName = parseAemColorToken(spec.name);
            const cleanSpecName = getDisplaySpecName(parsedSpecName.text);
            const headerText = applyTextColor(escapeHtml(cleanSpecName), '', parsedSpecName);
            rows.push(`<tr><th scope="row">${headerText}</th>${valueCols}</tr>`);
          }

          const rowsHtml = rows.join('');

          return [
            '<details class="comparison-category" open>',
            `  <summary>${escapeHtml(section.name)}</summary>`,
            '  <table class="comparison-table">',
            `    <colgroup>${colgroupHtml}</colgroup>`,
            `    <tbody>${rowsHtml}</tbody>`,
            '  </table>',
            '</details>'
          ].join('\n');
        }).join('\n\n');

        sectionsEl.innerHTML = sectionHtml;
      }

      buildSelectors();
      buildTables();
    }

    function countAemTokens(data) {
      if (!data || !Array.isArray(data.rows)) return 0;
      let count = 0;
      data.rows.forEach((row) => {
        if (row.type !== 'spec') return;
        const nameParsed = parseAemColorToken(row.name || '');
        if (nameParsed.token) count += 1;
        if (Array.isArray(row.values)) {
          row.values.forEach((v) => {
            const valueParsed = parseAemColorToken(v || '');
            if (valueParsed.token) count += 1;
          });
        }
      });
      return count;
    }

    // 綁定上傳事件
    $('#excel-upload').addEventListener('change', function(e) {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      const outputEl = $('#json-output');
      const copyBtn = $('#copy-btn');
      const statusEl = $('#json-status');

      reader.onload = function(event) {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: 'array', raw: true });
          
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];

          // 轉換 JSON 時加入 raw: false，強制輸出畫面上看到的原始格式
          const rawJson = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false });
          
          if (rawJson.length === 0) {
            outputEl.innerHTML = '<span class="error">表格沒有內容或格式不正確。</span>';
            copyBtn.style.display = 'none'; // 發生錯誤時隱藏按鈕
            return;
          }

          // 執行資料清洗
          const finalOutputJson = formatSpecData(rawJson);

          outputEl.textContent = JSON.stringify(finalOutputJson, null, 2);
          const htmlText = convertJsonToHtml(finalOutputJson);
          $('#html-output').textContent = htmlText;
          renderInteractivePreview(finalOutputJson, $('#html-preview'));

          const tokenCount = countAemTokens(finalOutputJson);
          if (tokenCount > 0) {
            statusEl.textContent = `Excel 轉換完成，已偵測 ${tokenCount} 個 AEM 色碼標記（[1]/[2]/[3]，每個 token 可在 AEM 設為單色或漸層）。`;
            statusEl.className = 'status ok';
          } else {
            statusEl.textContent = 'Excel 轉換完成，尚未偵測到 AEM 色碼標記（可在文字前加上 [1]、[2] 或 [3]）。';
            statusEl.className = 'status bad';
          }
          
          // 資料成功產出，顯示複製按鈕
          copyBtn.style.display = 'block';

        } catch (error) {
          console.error("轉換失敗：", error);
          outputEl.innerHTML = '<span class="error">檔案解析失敗，請確保上傳的是有效的 Excel/CSV 檔案。</span>';
          statusEl.textContent = `Excel 轉換失敗：${error.message}`;
          statusEl.className = 'status bad';
          copyBtn.style.display = 'none'; // 發生錯誤時隱藏按鈕
        }
      };

      reader.readAsArrayBuffer(file);
    });

    // 綁定複製按鈕事件
    $('#copy-btn').addEventListener('click', function() {
      const jsonText = $('#json-output').textContent;
      const btn = this;

      navigator.clipboard.writeText(jsonText).then(() => {
        // 複製成功的視覺回饋
        btn.textContent = '已複製！';
        btn.classList.add('success');
        
        // 2秒後恢復原狀
        setTimeout(() => {
          btn.textContent = '複製 JSON';
          btn.classList.remove('success');
        }, 2000);
      }).catch(err => {
        console.error("複製失敗：", err);
        alert("抱歉，複製失敗，請手動選取文字。");
      });
    });

    $('#convert-json-btn').addEventListener('click', function() {
      const inputText = $('#json-input').value.trim();
      const statusEl = $('#json-status');
      const htmlOutputEl = $('#html-output');
      const copyHtmlBtn = $('#copy-html-btn');
      const previewEl = $('#html-preview');

      if (!inputText) {
        statusEl.textContent = '請先貼上 JSON 內容。';
        statusEl.className = 'status bad';
        htmlOutputEl.textContent = '尚未轉換';
        previewEl.innerHTML = '';
        copyHtmlBtn.style.display = 'none';
        return;
      }

      let parsed;
      try {
        parsed = JSON.parse(inputText);
      } catch (error) {
        statusEl.textContent = `JSON 格式錯誤：${error.message}`;
        statusEl.className = 'status bad';
        htmlOutputEl.textContent = '尚未轉換';
        previewEl.innerHTML = '';
        copyHtmlBtn.style.display = 'none';
        return;
      }

      const errors = validateComparisonJson(parsed);
      if (errors.length > 0) {
        statusEl.textContent = `驗證失敗：${errors.join(' | ')}`;
        statusEl.className = 'status bad';
        htmlOutputEl.textContent = '尚未轉換';
        previewEl.innerHTML = '';
        copyHtmlBtn.style.display = 'none';
        return;
      }

      const htmlText = convertJsonToHtml(parsed);
      htmlOutputEl.textContent = htmlText;
      renderInteractivePreview(parsed, previewEl);
      statusEl.textContent = '驗證成功，已產生可收合 Category 的 HTML Table。';
      statusEl.className = 'status ok';
      copyHtmlBtn.style.display = 'inline-block';
    });

    $('#copy-html-btn').addEventListener('click', function() {
      const htmlText = $('#html-output').textContent;
      const btn = this;

      navigator.clipboard.writeText(htmlText).then(() => {
        btn.textContent = '已複製！';
        setTimeout(() => {
          btn.textContent = '複製 HTML';
        }, 2000);
      }).catch((err) => {
        console.error('HTML 複製失敗：', err);
        alert('抱歉，複製失敗，請手動選取文字。');
      });
    });
  };
})();