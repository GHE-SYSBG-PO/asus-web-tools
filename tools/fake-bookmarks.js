(function () {
  // 1. 限定 Scope 的 CSS
  const css = `
    #fake-bookmarks-app {
      padding: 40px 32px;
      background: #111;
      min-height: 100%;
      color: #fff;
      font-family: inherit;
    }
    #fake-bookmarks-app .fb-title {
      font-size: 1.875rem;
      font-weight: 700;
      text-align: center;
      margin: 0 0 8px;
    }
    #fake-bookmarks-app .fb-subtitle {
      font-size: 1.125rem;
      font-weight: 700;
      text-align: center;
      margin: 0 auto 48px;
      display: block;
    }
    #fake-bookmarks-app .fb-categories {
      display: flex;
      flex-direction: column;
      gap: 64px;
    }
    #fake-bookmarks-app .fb-category {
      /* nothing special */
    }
    #fake-bookmarks-app .fb-category-title {
      font-size: 1.5rem;
      font-weight: 600;
      margin: 0 0 16px;
    }
    #fake-bookmarks-app .fb-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
      gap: 24px;
    }
    #fake-bookmarks-app .fb-card {
      background: #fff;
      color: #111;
      border-radius: 8px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.3);
      overflow: hidden;
      transition: box-shadow 0.3s;
    }
    #fake-bookmarks-app .fb-card:hover {
      box-shadow: 0 8px 24px rgba(0,0,0,0.5);
    }
    #fake-bookmarks-app .fb-card__img-wrap {
      position: relative;
    }
    #fake-bookmarks-app .fb-card__img {
      width: 100%;
      height: 192px;
      object-fit: cover;
      background: #e5e7eb;
      display: block;
    }
    #fake-bookmarks-app .fb-card__zoom {
      position: absolute;
      top: 8px;
      right: 8px;
      background: #fff;
      border: none;
      border-radius: 50%;
      padding: 4px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.25);
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    #fake-bookmarks-app .fb-card__zoom:hover svg {
      color: #1e40af;
    }
    #fake-bookmarks-app .fb-card__body {
      padding: 16px;
    }
    #fake-bookmarks-app .fb-card__name {
      font-size: 1.125rem;
      font-weight: 700;
      margin: 0 0 4px;
      color: #111;
    }
    #fake-bookmarks-app .fb-card__meta {
      font-size: 0.75rem;
      color: #6b7280;
      margin: 0 0 16px;
    }
    #fake-bookmarks-app .fb-card__bookmark {
      display: inline-block;
      border: 1px solid #d9d29b;
      padding: 6px 12px;
      border-radius: 18px;
      background: #ffec48;
      color: #111;
      text-decoration: none;
      font-size: 0.875rem;
      cursor: grab;
    }
    #fake-bookmarks-app .fb-popup-overlay {
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.75);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 50;
    }
    #fake-bookmarks-app .fb-popup-inner {
      position: relative;
    }
    #fake-bookmarks-app .fb-popup-img {
      max-height: 90vh;
      max-width: 90vw;
    }
    #fake-bookmarks-app .fb-popup-close {
      position: absolute;
      top: 8px;
      right: 8px;
      color: #fff;
      font-size: 1.5rem;
      background: none;
      border: none;
      cursor: pointer;
      line-height: 1;
    }
    #fake-bookmarks-app .fb-popup-close:hover {
      color: #f87171;
    }
    #fake-bookmarks-app .fb-error {
      color: #f87171;
      text-align: center;
      padding: 32px;
    }
  `;

  // 2. HTML 骨架
  const html = `
    <h1 class="fb-title">Fake Bookmarks</h1>
    <span class="fb-subtitle">直接把黃色按鈕拖曳到書籤, 即可點擊使用</span>
    <div id="fb-categories" class="fb-categories"></div>
  `;

  // 3. 注入 CSS 與 HTML
  const appContainer = document.getElementById('fake-bookmarks-app');
  const styleEl = document.createElement('style');
  styleEl.textContent = css;
  appContainer.appendChild(styleEl);
  appContainer.insertAdjacentHTML('beforeend', html);

  // 4. JS 邏輯
  const $ = selector => appContainer.querySelector(selector);

  const categoriesContainer = $('#fb-categories');

  // 開啟圖片 Popup
  function openPopup(imgSrc) {
    const overlay = document.createElement('div');
    overlay.className = 'fb-popup-overlay';
    overlay.innerHTML = `
      <div class="fb-popup-inner">
        <img src="${imgSrc}" alt="Zoomed Image" class="fb-popup-img">
        <button class="fb-popup-close">&times;</button>
      </div>
    `;
    appContainer.appendChild(overlay);

    overlay.querySelector('.fb-popup-close').addEventListener('click', () => {
      appContainer.removeChild(overlay);
    });
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) appContainer.removeChild(overlay);
    });
  }

  // 渲染資料
  function render(data) {
    let index = 0;
    Object.keys(data).forEach(category => {
      const categoryDiv = document.createElement('div');
      categoryDiv.className = 'fb-category';

      const title = document.createElement('h2');
      title.className = 'fb-category-title';
      title.textContent = `${category} (${data[category].length})`;
      categoryDiv.appendChild(title);

      const grid = document.createElement('div');
      grid.className = 'fb-grid';

      data[category].forEach(item => {
        index++;
        const card = document.createElement('div');
        card.className = 'fb-card';
        card.innerHTML = `
          <div class="fb-card__img-wrap">
            <img src="${item.imgSrc}" alt="${item.name}" class="fb-card__img">
            <button class="fb-card__zoom" data-img="${item.imgSrc}" title="放大圖片">
              <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="color:#4b5563">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/>
              </svg>
            </button>
          </div>
          <div class="fb-card__body">
            <h3 class="fb-card__name">${index}. ${item.name}</h3>
            <p class="fb-card__meta">${item.type} - ${item.category}</p>
            <a href="${item.url}" class="fb-card__bookmark" target="_blank">${item.name}</a>
          </div>
        `;

        card.querySelector('.fb-card__zoom').addEventListener('click', () => {
          openPopup(item.imgSrc);
        });

        grid.appendChild(card);
      });

      categoryDiv.appendChild(grid);
      categoriesContainer.appendChild(categoryDiv);
    });
  }

  // 載入 JSON
  fetch('./modules.json')
    .then(res => {
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return res.json();
    })
    .then(data => render(data))
    .catch(err => {
      console.error('[fake-bookmarks] Error loading modules.json:', err);
      categoriesContainer.innerHTML = `<p class="fb-error">無法載入書籤資料：${err.message}</p>`;
    });
})();
