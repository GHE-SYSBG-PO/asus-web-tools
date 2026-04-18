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
  };
})();