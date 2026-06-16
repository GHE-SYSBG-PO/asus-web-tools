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
      
      // 3. 剩下的欄位才是真正的產品型號
      const productNames = cleanKeys.slice(2);

      // 4. 迭代每一列資料
      const formattedRows = rawData.map(row => {
        const specHidden = row[specColumnHidden]?.trim() || "";
        const specName = row[specColumnName]?.trim() || "";

        // 檢查所有產品欄位是否皆為空值
        const isCategory = productNames.every(product => {
          const val = row[product];
          return val === undefined || val === null || String(val).trim() === "";
        });

        // 回傳對應結構
        if (isCategory) {
          // 防呆機制：如果是 Category 列，但 specName 本身也是空的，代表這是一整行純空白列，將其標記為 null 稍後濾除
          if (!specName) return null;

          return {
            type: "category",
            name: specName,
            label: specHidden
          };
        } else {
          const specValues = productNames.map(product => {
            const val = row[product];
            return (val !== undefined && val !== null && val !== "") ? String(val).trim() : "";
          });

          return {
            type: "spec",
            hidden: specHidden,
            name: specName,
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
      return escapeHtml(value).replace(/\r?\n/g, '<br>');
    }

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
      const sections = [];
      let current = null;

      data.rows.forEach((row) => {
        if (row.type === 'category') {
          current = { name: row.name, specs: [] };
          sections.push(current);
          return;
        }

        if (!current) {
          current = { name: 'General', specs: [] };
          sections.push(current);
        }

        current.specs.push(row);
      });

      const inlineData = JSON.stringify({
        products: data.products,
        sections
      }).replace(/</g, '\\u003c');

      return [
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
        '  function buildSelectors() {',
        '    selectorsEl.innerHTML = "";',
        '    visible.forEach((productIdx, colIdx) => {',
        '      const wrap = document.createElement("div");',
        '      wrap.className = "comparison-selector-item";',
        '',
        '      const label = document.createElement("label");',
        '      label.textContent = `Product ${colIdx + 1}`;',
        '',
        '      const select = document.createElement("select");',
        '      select.className = "comparison-product-select";',
        '',
        '      data.products.forEach((name, idx) => {',
        '        const option = document.createElement("option");',
        '        option.value = String(idx);',
        '        option.textContent = name;',
        '        option.selected = idx === productIdx;',
        '        select.appendChild(option);',
        '      });',
        '',
        '      select.addEventListener("change", (event) => {',
        '        const nextIdx = Number(event.target.value);',
        '        if (Number.isNaN(nextIdx)) return;',
        '        if (visible.includes(nextIdx)) {',
        '          event.target.value = String(visible[colIdx]);',
        '          return;',
        '        }',
        '        visible[colIdx] = nextIdx;',
        '        buildSelectors();',
        '        buildTables();',
        '      });',
        '',
        '      wrap.appendChild(label);',
        '      wrap.appendChild(select);',
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
        '      const rowsHtml = section.specs.map((spec) => {',
        '        const hiddenAttr = String(spec.hidden || "").toUpperCase() === "TRUE" ? " data-hidden=\"true\"" : "";',
        '        const valueCols = visible',
        '          .map((idx) => {',
        '            const rawValue = spec.values && spec.values[idx] !== undefined ? spec.values[idx] : "";',
        '            return `<td>${escapeHtml(rawValue).replace(/\\r?\\n/g, "<br>")}</td>`;',
        '          })',
        '          .join("");',
        '',
        '        return `<tr${hiddenAttr}><th scope="row">${escapeHtml(spec.name)}</th>${valueCols}</tr>`;',
        '      }).join("");',
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

        current.specs.push(row);
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

          const select = document.createElement('select');
          select.className = 'comparison-product-select';

          data.products.forEach((name, idx) => {
            const option = document.createElement('option');
            option.value = String(idx);
            option.textContent = name;
            option.selected = idx === productIdx;
            select.appendChild(option);
          });

          select.addEventListener('change', (event) => {
            const nextIdx = Number(event.target.value);
            if (Number.isNaN(nextIdx)) return;
            if (visible.includes(nextIdx)) {
              event.target.value = String(visible[colIdx]);
              return;
            }
            visible[colIdx] = nextIdx;
            buildSelectors();
            buildTables();
          });

          wrap.appendChild(label);
          wrap.appendChild(select);
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
          const rowsHtml = section.specs.map((spec) => {
            const hiddenAttr = String(spec.hidden || '').toUpperCase() === 'TRUE' ? ' data-hidden="true"' : '';
            const valueCols = visible
              .map((idx) => {
                const rawValue = spec.values && spec.values[idx] !== undefined ? spec.values[idx] : '';
                return `<td>${toCellHtml(rawValue)}</td>`;
              })
              .join('');

            return `<tr${hiddenAttr}><th scope="row">${escapeHtml(spec.name)}</th>${valueCols}</tr>`;
          }).join('');

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

    // 綁定上傳事件
    $('#excel-upload').addEventListener('change', function(e) {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      const outputEl = $('#json-output');
      const copyBtn = $('#copy-btn');

      reader.onload = function(event) {
        try {
          const data = new Uint8Array(event.target.result);
          
          // 讀取時加入 raw: true，避免 CSV 的 16:9 被誤判為時間
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
          const finalJson = formatSpecData(rawJson);
          outputEl.textContent = JSON.stringify(finalJson, null, 2);
          
          // 資料成功產出，顯示複製按鈕
          copyBtn.style.display = 'block';

        } catch (error) {
          console.error("轉換失敗：", error);
          outputEl.innerHTML = '<span class="error">檔案解析失敗，請確保上傳的是有效的 Excel/CSV 檔案。</span>';
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