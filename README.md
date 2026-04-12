# Dev Tools Hub

這是一個前端開發輔助工具整合平台，採用模組化設計（Single Page Application 概念）。旨在將多個零散的小工具（如：ADA 對比度檢查、漸層產生器等）集中於一個頁面內，提供一致的 UI 體驗並減少切換網頁的時間。

## 目錄結構
```text
web-tools/
├── index.html        # 主網頁（包含側邊欄導覽與掛載點）
├── style.css         # 全域樣式（主題設計、導覽列、版面配置）
├── app.js            # 主程式邏輯（處理工具切換、localStorage 狀態儲存）
├── tools/            # 各式工具模組資料夾
│   ├── ada-contrast.js
│   └── linear-gradient.js
└── README.md         # 開發/AI 輔助開發指南
```

---

## 🤖 AI 開發指南：如何新增一個工具

當你需要（或要求 AI）將新的工具整合進此平台時，請嚴格遵循以下四個步驟：

### 步驟 1：新增側邊欄選單 (index.html)
在 `index.html` 的 `<nav class="sidebar__nav">` 中新增一個 `<button>` 節點。
* **屬性要求**：必須包含 `class="nav-item"` 以及自定義的 `data-tool="your-tool-id"`。
* 內容依照現有格式，配置 icon 與名稱。

```html
<!-- 新增按鈕範例 -->
<button class="nav-item" data-tool="your-tool-id" role="tab" aria-selected="false">
  <span class="nav-item__icon">✦</span>
  <span class="nav-item__label">New Tool</span>
  <span class="nav-item__badge">Tag</span>
</button>
```

### 步驟 2：新增工具掛載面板 (index.html)
在 `<main class="main-content">` 中新增一個對應的 `<section>`。
* **ID 命名**：`id="panel-your-tool-id"`（`panel-` 加上剛剛定義的 `data-tool`）。
* **初始隱藏**：加上 `class="tool-panel hidden"`。
* 內部建立一個 `div` 提供為該工具的渲染節點。

```html
<!-- 新增面板範例 -->
<section class="tool-panel hidden" id="panel-your-tool-id" role="tabpanel">
  <div id="your-tool-app"></div>
</section>
```

### 步驟 3：引入工具腳本 (index.html)
在文件底部的 `<body>` 關閉標籤前，將新的 js 模組引入：

```html
<script src="tools/your-tool.js"></script>
<!-- 確保在 app.js 前載入或是本身封裝好 -->
<script src="app.js"></script>
```

### 步驟 4：開發工具模組檔案 (tools/your-tool.js)
為了防止與其他工具發生變數或是 CSS 污染，每個工具**必須使用 IIFE (Immediately Invoked Function Expression)** 包覆，並將專屬的 HTML 與 CSS 注入到專屬的 div 容器內。

**封裝範例模板：**

```javascript
(function() {
  // 1. 限定 Scope 的 CSS (所有的 CSS 選擇器都必須限定在 #your-tool-app 下)
  const css = `
    #your-tool-app {
      padding: 32px;
      color: var(--text-primary);
    }
    #your-tool-app .my-custom-btn {
      background: var(--accent);
      /* ... */
    }
  `;

  // 2. 工具的 HTML 模板
  const html = `
    <div class="tool-wrapper">
      <h2>新工具標題</h2>
      <button class="my-custom-btn" id="yt-btn">點擊我</button>
    </div>
  `;

  // 3. 將 CSS 和 HTML 注入掛載點
  const appContainer = document.getElementById('your-tool-app');
  const styleEl = document.createElement('style');
  styleEl.textContent = css;
  appContainer.appendChild(styleEl);
  appContainer.insertAdjacentHTML('beforeend', html);

  // 4. JS 元素選取與邏輯操作
  // 提示：請避免使用全域的 document.querySelector，改用 appContainer.querySelector 確保隔離性
  const $ = selector => appContainer.querySelector(selector);
  
  const btn = $('#yt-btn');
  btn.addEventListener('click', () => {
    console.log('Button clicked!');
  });
})();
```

---

## 🎨 設計系統指南 (CSS 規範)

開發新工具時，請盡量**複用全域宣告的 CSS 變數 (`style.css`)** 來保持視覺一致性（深色質感、毛玻璃）。

| 用途 | CSS 變數 | 備註說明 |
| :--- | :--- | :--- |
| **主要背景** | `var(--bg-primary)` | 應用程式最底層顏色 |
| **卡片/區塊底色**| `var(--bg-card)` | 具備透明度的深色卡片 |
| **主要文字** | `var(--text-primary)` | 亮白色字體 |
| **次要文字** | `var(--text-secondary)`| 灰白色字體 |
| **主題強化色** | `var(--accent)` | 預設為紫色 |
| **圓角** | `var(--radius-sm/md/lg)`| 統一的邊角設計 |

### 開發守則與注意事項：
1. **🚫 絕對禁止**：在工具模組內覆寫全域標籤（如 `body {}`, `h1 {}`, `div {}` 等等）。
2. **✅ 最佳實踐**：永遠將你的工具 CSS 開頭綁定指定的 ID 容器，例如 `#your-tool-app .button { ... }`，或是為你的 Class 加上特有前綴（例如 `yt-button`）。
3. **無縫切換**：`app.js` 會自動透過讀取選單按鈕的 `data-tool` 來控制特定 `#panel-{tool-id}` 的顯示與隱藏（加上 `.hidden` class），不需要在新工具模組內重複處理切換邏輯。
