// ==UserScript==
// @name         豆瓣标记导出（支持NeoDB导入）
// @name:en      Douban to NeoDB
// @name:zh-CN   豆瓣标记导出（支持NeoDB导入）
// @namespace    https://github.com/zhucebuliaomax/douban2neodb
// @version      1.1.0
// @description  导出豆瓣电影、读书、音乐和游戏收藏为 Excel/JSON；Excel 兼容豆坟格式，可导入 NeoDB。
// @description:en Export Douban movies, books, music and games to Excel/JSON; the Excel file is compatible with the Doufen format and can be imported into NeoDB.
// @author       ming; Modified by Max
// @match        https://*.douban.com/*
// @match        https://douban.com/*
// @match        https://www.douban.com/people/*
// @match        https://movie.douban.com/mine*
// @match        https://movie.douban.com/people/*/collect*
// @match        https://movie.douban.com/people/*/wish*
// @match        https://movie.douban.com/people/*/do*
// @match        https://book.douban.com/mine*
// @match        https://book.douban.com/people/*/collect*
// @match        https://book.douban.com/people/*/wish*
// @match        https://book.douban.com/people/*/do*
// @match        https://music.douban.com/mine*
// @match        https://music.douban.com/people/*/collect*
// @match        https://music.douban.com/people/*/wish*
// @match        https://music.douban.com/people/*/do*
// @match        https://www.douban.com/people/*/games*
// @require      https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js
// @grant        GM_addStyle
// @license      MIT
// @homepage     https://github.com/zhucebuliaomax/douban2neodb
// @supportURL   https://github.com/zhucebuliaomax/douban2neodb/issues
// ==/UserScript==

(function () {
    'use strict';

    const CONFIG = {
        minDelay: 1200,
        maxDelay: 3000,
        stateKey: 'db_export_state_v2',
        dataKey: 'db_export_data_v2'
    };

    const CATEGORIES = {
        movie: { label: '电影', icon: '🎬', sheet: '电影收藏', file: 'Movie', pageSize: 15 },
        book: { label: '读书', icon: '📚', sheet: '读书收藏', file: 'Book', pageSize: 15 },
        music: { label: '音乐', icon: '🎵', sheet: '音乐收藏', file: 'Music', pageSize: 15 },
        game: { label: '游戏', icon: '🎮', sheet: '游戏收藏', file: 'Game', pageSize: 15 }
    };

    const STATUS_LABELS = {
        movie: { wish: '想看', do: '在看', collect: '看过' },
        book: { wish: '想读', do: '在读', collect: '读过' },
        music: { wish: '想听', do: '在听', collect: '听过' },
        game: { wish: '想玩', do: '在玩', collect: '玩过' }
    };
    const STATUS_ORDER = ['wish', 'do', 'collect'];

    // 豆坟（NeoDB 可导入）模板：固定全部 sheet 顺序、表头与列宽；当前分类状态的数据写入对应 sheet，其余 sheet 仅保留表头。
    const DOUFEN_SHEETS = ['看过', '在看', '想看', '听过', '在听', '想听', '读过', '在读', '想读', '玩过', '在玩', '想玩', '看过的舞台剧', '想看的舞台剧'];
    const DOUFEN_HEADERS = ['标题', '简介', '豆瓣评分', '链接', '创建时间', '我的评分', '标签', '评论', '可见性'];
    const DOUFEN_COL_WIDTHS = [40, 50, 10, 50, 20, 10, 30, 50, 10];
    const DOUFEN_STATUS_SHEETS = {
        movie: { collect: '看过', do: '在看', wish: '想看' },
        music: { collect: '听过', do: '在听', wish: '想听' },
        book: { collect: '读过', do: '在读', wish: '想读' },
        game: { collect: '玩过', do: '在玩', wish: '想玩' }
    };

    const styleText = `
        #db-export-btn-bar {
            position: fixed; top: 110px; right: 20px; z-index: 9999;
            display: flex; align-items: center; gap: 8px;
        }
        #db-export-summary-btn {
            padding: 10px 18px; border: 0; border-radius: 24px; cursor: pointer;
            background: #3eaf7c; color: #fff; font-size: 14px; font-weight: 700;
            box-shadow: 0 4px 12px rgba(62,175,124,.35); transition: .2s;
        }
        #db-export-summary-btn:hover { background: #339268; transform: translateY(-1px); }
        #db-export-abort-btn {
            display: inline-flex; align-items: center;
            padding: 10px 18px; border: 0; border-radius: 24px; cursor: pointer;
            background: #e74c3c; color: #fff; font-size: 14px; font-weight: 700;
            box-shadow: 0 4px 12px rgba(231,76,60,.35); transition: .2s;
        }
        #db-export-abort-btn .db-abort-extra {
            display: grid;
            grid-template-columns: 0fr;
            transition: grid-template-columns .25s ease;
        }
        #db-export-abort-btn .db-abort-extra-inner {
            overflow: hidden;
            white-space: nowrap;
        }
        #db-export-abort-btn:hover { background: #c0392b; transform: translateY(-1px); }
        #db-export-abort-btn:hover .db-abort-extra { grid-template-columns: 1fr; }
        #db-export-summary-overlay, #db-export-modal-overlay {
            position: fixed; inset: 0; z-index: 10000; background: rgba(0,0,0,.52);
            display: flex; align-items: center; justify-content: center;
        }
        #db-export-summary-panel, #db-export-modal {
            box-sizing: border-box; width: min(520px, 92vw); max-height: 86vh; overflow-y: auto;
            padding: 24px; border-radius: 10px; background: #fff; color: #333;
            font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
            box-shadow: 0 12px 36px rgba(0,0,0,.22); animation: dbFadeIn .2s ease-out;
        }
        #db-export-summary-panel h3, #db-export-modal h3 { margin: 0; padding-bottom: 12px; border-bottom: 2px solid #3eaf7c; font-size: 18px; }
        .db-summary-help { margin: 12px 0 16px; color:#666; font-size: 13px; line-height: 1.6; }
        .db-summary-list { display:grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }
        .db-summary-card { display:flex; align-items:center; gap: 10px; padding: 12px; border:1px solid #e8e8e8; border-radius: 8px; }
        .db-summary-cover { width: 40px; height: 54px; flex: 0 0 40px; object-fit: cover; border-radius: 3px; background:#f2f2f2; }
        .db-summary-main { min-width:0; flex:1; }
        .db-summary-title { font-weight:700; font-size:14px; }
        .db-summary-meta { margin-top:4px; color:#888; font-size:12px; line-height:1.4; }
        .db-summary-action { margin-top:8px; padding: 6px 10px; border:0; border-radius:5px; cursor:pointer; color:#fff; background:#3eaf7c; font-size:12px; }
        .db-summary-action:hover { background:#339268; }
        .db-btn { padding: 8px 14px; border: 0; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 700; }
        .db-btn-primary { background:#3eaf7c; color:#fff; }
        .db-btn-primary:hover { background:#339268; }
        .db-btn-secondary { background:#f0f0f0; color:#666; }
        .db-btn-secondary:hover { background:#e2e2e2; }
        .db-checkbox-label { display:flex; align-items:center; gap:9px; cursor:pointer; color:#444; font-size:14px; user-select:none; }
        .db-checkbox-label input { width:16px; height:16px; accent-color:#3eaf7c; }
        .db-page-range { display:flex; align-items:center; gap:8px; margin:9px 0 0 25px; padding:10px; border-radius:7px; background:#f7f9f8; }
        .db-page-range[hidden] { display:none; }
        .db-page-range input { box-sizing:border-box; width:72px; padding:5px 7px; border:1px solid #d7dedb; border-radius:5px; }
        .db-download-section { padding:12px; border:1px solid #e6ebe8; border-radius:8px; }
        .db-download-title { margin-bottom:8px; color:#333; font-weight:700; font-size:13px; }
        .db-download-actions { display:flex; flex-direction:column; gap:8px; }
        .db-btn-group { display:flex; justify-content:flex-end; gap:9px; margin-top:20px; }
        .db-note { color:#777; font-size:12px; line-height:1.6; }
        @keyframes dbFadeIn { from { opacity:0; transform:translateY(-12px); } to { opacity:1; transform:translateY(0); } }
        @media (max-width: 560px) { .db-summary-list { grid-template-columns: 1fr; } }
    `;

    function addStyle(css) {
        if (typeof GM_addStyle === 'function') GM_addStyle(css);
        else {
            const style = document.createElement('style');
            style.textContent = css;
            document.head.appendChild(style);
        }
    }

    function textOf(el) {
        return el ? (el.innerText || el.textContent || '').replace(/\s+/g, ' ').trim() : '';
    }

    function parseJson(value, fallback) {
        try { return value ? JSON.parse(value) : fallback; } catch (e) { return fallback; }
    }

    function getProfileSlug() {
        const match = location.pathname.match(/^\/people\/([^/]+)/);
        if (match) return match[1];
        return window._GLOBAL_NAV && window._GLOBAL_NAV.USER_ID ? String(window._GLOBAL_NAV.USER_ID) : '';
    }

    function detectContext() {
        const host = location.hostname;
        const path = location.pathname;
        if (host === 'movie.douban.com' && (/\/mine/.test(path) || /\/people\/[^/]+\/(?:collect|wish|do)/.test(path))) return 'movie';
        if (host === 'book.douban.com' && (/\/mine/.test(path) || /\/people\/[^/]+\/(?:collect|wish|do)/.test(path))) return 'book';
        if (host === 'music.douban.com' && (/\/mine/.test(path) || /\/people\/[^/]+\/(?:collect|wish|do)/.test(path))) return 'music';
        if (host === 'www.douban.com' && /\/people\/[^/]+\/games/.test(path)) return 'game';
        if (host === 'www.douban.com' && /^\/people\/[^/]+\/?$/.test(path)) return 'profile';
        if ((host === 'douban.com' || host.endsWith('.douban.com')) && host !== 'accounts.douban.com') return 'generic';
        return '';
    }

    function storageKey(key) {
        // /mine 首页的下一页通常会跳到 /people/<id>/collect，按 host 存储可跨分页保持状态。
        return `${key}:${location.hostname}`;
    }

    function getState() {
        return parseJson(localStorage.getItem(storageKey(CONFIG.stateKey)), { status: 'idle' });
    }

    function setState(state) {
        localStorage.setItem(storageKey(CONFIG.stateKey), JSON.stringify(state));
    }

    function getStoredData() {
        return parseJson(localStorage.getItem(storageKey(CONFIG.dataKey)), []);
    }

    function setStoredData(data) {
        localStorage.setItem(storageKey(CONFIG.dataKey), JSON.stringify(data));
    }

    function cleanTitle(value) {
        return value.replace(/^\[.*?\]\s*/, '').replace(/\s+/g, ' ').trim();
    }

    function extractDate(value) {
        const match = value.match(/\d{4}-\d{1,2}-\d{1,2}/);
        return match ? match[0] : value.trim();
    }

    function parseRating(el) {
        if (!el) return '';
        const attr = el.getAttribute('data-rating');
        if (attr && /^\d+(?:\.\d+)?$/.test(attr)) return Number(attr);
        const cls = el.className || '';
        const rating = cls.match(/rating(\d)-t/);
        if (rating) return Number(rating[1]);
        const stars = cls.match(/allstar(\d+)/);
        if (stars) return Number(stars[1]) / 10;
        return '';
    }

    function getRating(item) {
        return parseRating(item.querySelector('[class^="rating"][class$="-t"], [class*="allstar"], [data-rating]'));
    }

    function getId(link) {
        const match = (link || '').match(/\/(?:subject|game)\/(\d+)/);
        return match ? match[1] : '';
    }

    function getStatusFromUrl() {
        const url = new URL(location.href);
        const param = url.searchParams.get('status') || url.searchParams.get('action');
        if (param) return param;
        const match = url.pathname.match(/\/(collect|wish|do)\/?$/);
        return match ? match[1] : 'collect';
    }

    function baseRecord(category, link) {
        return {
            category,
            id: getId(link),
            title: '',
            rating: '',
            date: '',
            status: getStatusFromUrl(),
            tags: '',
            comment: '',
            intro: '',
            link
        };
    }

    function parseMoviePage() {
        const items = [...document.querySelectorAll('.grid-view .item, .list-view .item')]
            .filter(item => item.querySelector('a[href*="/subject/"]'));
        return items.map(item => {
            const titleLink = item.querySelector('.title a[href*="/subject/"]') || item.querySelector('a[href*="/subject/"]');
            const record = baseRecord('movie', titleLink ? titleLink.href : '');
            record.title = cleanTitle(textOf(titleLink));
            record.rating = getRating(item);
            record.date = extractDate(textOf(item.querySelector('.date')));
            record.tags = textOf(item.querySelector('.tags')).replace(/^标签[:：]\s*/, '');
            record.comment = textOf(item.querySelector('.comment'));
            record.intro = textOf(item.querySelector('.intro'));
            return record;
        });
    }

    function parseBookPage() {
        return [...document.querySelectorAll('.subject-item')].map(item => {
            const titleLink = item.querySelector('.info h2 a[href*="/subject/"]') || item.querySelector('a[href*="/subject/"]');
            const record = baseRecord('book', titleLink ? titleLink.href : '');
            record.title = cleanTitle(textOf(titleLink));
            record.rating = getRating(item);
            record.date = extractDate(textOf(item.querySelector('.date')));
            record.comment = textOf(item.querySelector('.comment'));
            record.intro = textOf(item.querySelector('.pub'));
            return record;
        }).filter(record => record.link);
    }

    function parseMusicPage() {
        return [...document.querySelectorAll('.item.comment-item, .item')]
            .filter(item => item.querySelector('a[href*="/subject/"]'))
            .map(item => {
                const titleLink = item.querySelector('.title a[href*="/subject/"]') || item.querySelector('a[href*="/subject/"]');
                const record = baseRecord('music', titleLink ? titleLink.href : '');
                record.title = cleanTitle(textOf(titleLink));
                record.rating = getRating(item);
                record.date = extractDate(textOf(item.querySelector('.date')));
                record.comment = textOf(item.querySelector('.comment'));
                record.intro = textOf(item.querySelector('.intro'));
                return record;
            }).filter(record => record.link);
    }

    function parseGamePage() {
        return [...document.querySelectorAll('.game-list .common-item')].map(item => {
            const titleLink = item.querySelector('.title a[href*="/game/"]') || item.querySelector('a[href*="/game/"]');
            const record = baseRecord('game', titleLink ? titleLink.href : '');
            record.title = cleanTitle(textOf(titleLink));
            record.rating = getRating(item);
            record.date = extractDate(textOf(item.querySelector('.date')));
            const desc = item.querySelector('.desc');
            if (desc) {
                const clone = desc.cloneNode(true);
                clone.querySelector('.rating-info')?.remove();
                record.intro = textOf(clone);
            }
            const comment = [...item.querySelectorAll('.content > div')]
                .find(el => !el.classList.contains('title') && !el.classList.contains('desc') && !el.classList.contains('user-operation'));
            record.comment = textOf(comment);
            return record;
        }).filter(record => record.link);
    }

    async function enrichMovieFromList(records) {
        try {
            const listUrl = new URL(location.href);
            listUrl.searchParams.set('mode', 'list');
            const response = await fetch(listUrl.href, { credentials: 'include' });
            if (!response.ok) return records;
            const html = await response.text();
            const doc = new DOMParser().parseFromString(html, 'text/html');
            const detailMap = new Map();
            [...doc.querySelectorAll('.list-view .item')].forEach(item => {
                const link = item.querySelector('.title a[href*="/subject/"]');
                if (!link) return;
                detailMap.set(getId(link.href), {
                    tags: textOf(item.querySelector('.tags')).replace(/^标签[:：]\s*/, ''),
                    comment: textOf(item.querySelector('.comment'))
                });
            });
            records.forEach(record => Object.assign(record, detailMap.get(record.id) || {}));
        } catch (e) {
            console.warn('[Douban Export] 无法补充电影列表字段:', e);
        }
        return records;
    }

    async function scrapeCurrentPage(category) {
        let records;
        if (category === 'movie') records = parseMoviePage();
        else if (category === 'book') records = parseBookPage();
        else if (category === 'music') records = parseMusicPage();
        else records = parseGamePage();
        if (category === 'movie') records = await enrichMovieFromList(records);
        return records;
    }

    function getNextPage() {
        const next = document.querySelector('.paginator .next a[href]');
        return next && next.href && !next.href.startsWith('javascript:') ? next.href : '';
    }

    function getCurrentPageNumber(category) {
        const current = Number.parseInt(textOf(document.querySelector('.paginator .thispage')), 10);
        if (Number.isInteger(current) && current > 0) return current;
        const start = Number.parseInt(new URL(location.href).searchParams.get('start') || '0', 10);
        return Math.floor((Number.isFinite(start) ? start : 0) / CATEGORIES[category].pageSize) + 1;
    }

    function getTotalPageCount() {
        const pages = [...document.querySelectorAll('.paginator .thispage, .paginator a')]
            .map(element => Number.parseInt(textOf(element), 10))
            .filter(page => Number.isInteger(page) && page > 0);
        return pages.length ? Math.max(...pages) : 1;
    }

    function formatPageRange(pageRange) {
        return pageRange ? `第 ${pageRange.startPage}～${pageRange.endPage} 页` : '全部页（从第 1 页开始）';
    }

    function escapeHtml(value) {
        return String(value || '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
    }

    function getCategoryUrl(category, slug, section) {
        const link = section && section.querySelector('a[href*="status=collect"], a[href*="games?action=collect"]');
        if (link) return new URL(link.href, location.href).href;
        if (category === 'movie') return slug ? `https://movie.douban.com/people/${slug}/collect?mode=grid` : 'https://movie.douban.com/mine?status=collect&mode=grid';
        if (category === 'book') return slug ? `https://book.douban.com/people/${slug}/collect?mode=grid` : 'https://book.douban.com/mine?status=collect&mode=grid';
        if (category === 'music') return slug ? `https://music.douban.com/people/${slug}/collect?mode=grid` : 'https://music.douban.com/mine?status=collect&mode=grid';
        return slug ? `https://www.douban.com/people/${slug}/games?action=collect` : 'https://www.douban.com/mine/';
    }

    function getCategoryStatusLabel(category) {
        const labels = STATUS_LABELS[category] || STATUS_LABELS.movie;
        return labels[getStatusFromUrl()] || labels.collect;
    }

    function getCategoryStatusUrl(category, slug, status) {
        if (category === 'game') {
            return slug
                ? `https://www.douban.com/people/${slug}/games?action=${status}`
                : 'https://www.douban.com/mine/';
        }
        const host = category === 'book' ? 'book' : category === 'music' ? 'music' : 'movie';
        return slug
            ? `https://${host}.douban.com/people/${slug}/${status}?mode=grid`
            : `https://${host}.douban.com/mine?status=${status}&mode=grid`;
    }

    function getCategoryStatusFilePart() {
        const parts = { collect: 'Collect', wish: 'Wish', do: 'Do' };
        return parts[getStatusFromUrl()] || 'Collect';
    }

    function withAutoExport(url) {
        const next = new URL(url, location.href);
        next.searchParams.set('db_export', '1');
        return next.href;
    }

    function getSummaryEntries(context) {
        const slug = getProfileSlug();
        const sections = [...document.querySelectorAll('.sort[id]')]
            .filter(section => Object.prototype.hasOwnProperty.call(CATEGORIES, section.id));
        if (context === 'profile' && sections.length) {
            return sections.map(section => {
                const category = section.id;
                const heading = section.querySelector('h2');
                const image = section.querySelector('img.climg');
                return {
                    category,
                    label: CATEGORIES[category].label,
                    icon: CATEGORIES[category].icon,
                    summary: textOf(heading).replace(/·/g, '').replace(/\s+/g, ' ').trim() || '打开收藏页查看全部',
                    cover: image ? image.src : '',
                    current: false,
                    url: withAutoExport(getCategoryUrl(category, slug, section))
                };
            });
        }
        return Object.keys(CATEGORIES).map(category => ({
            category,
            label: CATEGORIES[category].label,
            icon: CATEGORIES[category].icon,
            summary: category === context ? `当前页面：${document.title}` : '打开对应收藏页开始导出',
            cover: '',
            current: category === context,
            url: withAutoExport(getCategoryUrl(category, slug))
        }));
    }

    function showSummaryPanel(context) {
        if (document.getElementById('db-export-summary-overlay')) return;
        const entries = getSummaryEntries(context);
        const overlay = document.createElement('div');
        overlay.id = 'db-export-summary-overlay';
        overlay.innerHTML = `<div id="db-export-summary-panel" role="dialog" aria-label="书影音游戏数据汇总">
            <h3>📊 书影音游戏数据汇总</h3>
            <p class="db-summary-help">从这里选择分类。当前收藏页会直接开始导出全部字段，其他分类会在新标签页打开并自动开始导出。个人主页的栏目顺序沿用豆瓣原生页面顺序。</p>
            <div class="db-summary-list">${entries.map(entry => `<div class="db-summary-card">
                ${entry.cover ? `<img class="db-summary-cover" src="${escapeHtml(entry.cover)}" alt="${escapeHtml(entry.label)}封面">` : '<div class="db-summary-cover"></div>'}
                <div class="db-summary-main"><div class="db-summary-title">${escapeHtml(entry.icon)} ${escapeHtml(entry.label)}</div><div class="db-summary-meta">${escapeHtml(entry.summary)}</div><button class="db-summary-action" data-current="${entry.current ? '1' : '0'}" data-category="${escapeHtml(entry.category)}" data-url="${escapeHtml(entry.url)}">${entry.current ? '导出当前分类' : '去导出'}</button></div>
            </div>`).join('')}</div>
            <div class="db-btn-group"><button class="db-btn db-btn-secondary" id="db-close-summary">关闭</button></div>
        </div>`;
        document.body.appendChild(overlay);
        overlay.addEventListener('click', event => {
            if (event.target === overlay || event.target.id === 'db-close-summary') { overlay.remove(); return; }
            const action = event.target.closest('.db-summary-action');
            if (!action) return;
            if (action.dataset.current === '1') { overlay.remove(); showConfigPanel(action.dataset.category); return; }
            overlay.remove();
            showStatusChooser(action.dataset.category);
        });
    }

    function showStatusChooser(category) {
        if (document.getElementById('db-export-modal-overlay')) return;
        const slug = getProfileSlug();
        const labels = STATUS_LABELS[category];
        const overlay = document.createElement('div');
        overlay.id = 'db-export-modal-overlay';
        overlay.innerHTML = `<div id="db-export-modal">
            <h3>${CATEGORIES[category].icon} 选择要导出的${CATEGORIES[category].label}收藏</h3>
            <p class="db-note">请选择要导出的收藏状态，将在新标签页打开对应收藏页并自动开始导出。</p>
            <div class="db-download-actions">${STATUS_ORDER.map(status => `<button class="db-btn db-btn-primary" data-status="${status}" style="text-align:left">${labels[status]}</button>`).join('')}</div>
            <div class="db-btn-group"><button class="db-btn db-btn-secondary" id="db-status-cancel">取消</button></div>
        </div>`;
        document.body.appendChild(overlay);
        overlay.querySelector('#db-status-cancel').onclick = () => overlay.remove();
        overlay.querySelectorAll('button[data-status]').forEach(button => {
            button.onclick = () => {
                overlay.remove();
                window.open(withAutoExport(getCategoryStatusUrl(category, slug, button.dataset.status)), '_blank', 'noopener');
            };
        });
    }

    function showConfigPanel(category) {
        if (document.getElementById('db-export-modal-overlay')) return;
        if (!CATEGORIES[category]) return;
        const totalPages = getTotalPageCount();
        const overlay = document.createElement('div');
        overlay.id = 'db-export-modal-overlay';
        overlay.innerHTML = `<div id="db-export-modal">
            <h3>${CATEGORIES[category].icon} 导出${CATEGORIES[category].label}（${getCategoryStatusLabel(category)}）</h3>
            <p class="db-note">将导出全部数据字段，完成后可下载 JSON 和 Excel 文件。</p>
            <label class="db-checkbox-label" style="margin-top:12px"><input id="db-limit-pages" type="checkbox"><span><b>仅导出指定页码范围</b><br><small style="color:#888">默认不勾选，将从第 1 页导出到最后一页</small></span></label>
            <div class="db-page-range" id="db-page-range" hidden><label>从第 <input id="db-start-page" type="number" min="1" max="${totalPages}" value="1"> 页</label><span>至</span><label>第 <input id="db-end-page" type="number" min="1" max="${totalPages}" value="${totalPages}"> 页</label></div>
            <p class="db-note">当前共识别到 ${totalPages} 页，每页最多 ${CATEGORIES[category].pageSize} 条。无论从哪一页打开导出，未限制范围时都会先返回第 1 页。</p>
            <div class="db-btn-group"><button class="db-btn db-btn-secondary" id="db-cancel-btn">取消</button><button class="db-btn db-btn-primary" id="db-start-btn">开始抓取</button></div>
        </div>`;
        document.body.appendChild(overlay);
        const rangeToggle = overlay.querySelector('#db-limit-pages');
        const rangeFields = overlay.querySelector('#db-page-range');
        rangeToggle.onchange = () => { rangeFields.hidden = !rangeToggle.checked; };
        overlay.querySelector('#db-cancel-btn').onclick = () => overlay.remove();
        overlay.querySelector('#db-start-btn').onclick = () => {
            let pageRange = null;
            if (rangeToggle.checked) {
                const startPage = Number.parseInt(overlay.querySelector('#db-start-page').value, 10);
                const endPage = Number.parseInt(overlay.querySelector('#db-end-page').value, 10);
                if (!Number.isInteger(startPage) || !Number.isInteger(endPage) || startPage < 1 || endPage < startPage || endPage > totalPages) {
                    alert(`请输入 1～${totalPages} 之间的有效页码，且结束页不能小于起始页。`);
                    return;
                }
                pageRange = { startPage, endPage };
            }
            overlay.remove();
            startScraping(category, pageRange);
        };
    }

    function renderSummaryButton(context) {
        if (document.getElementById('db-export-summary-btn')) return;
        addStyle(styleText);
        const state = getState();
        const bar = document.createElement('div');
        bar.id = 'db-export-btn-bar';
        const button = document.createElement('button');
        button.id = 'db-export-summary-btn';
        button.type = 'button';
        button.textContent = state.status === 'running' && state.category === context ? '⏳ 抓取中 · 汇总' : '📊 书影音游戏汇总';
        button.title = '汇总并导航到具体分类导出';
        button.onclick = () => showSummaryPanel(context);
        bar.appendChild(button);
        if (state.status === 'running' && state.category === context) {
            const abortButton = document.createElement('button');
            abortButton.id = 'db-export-abort-btn';
            abortButton.type = 'button';
            abortButton.innerHTML = '终止抓取<span class="db-abort-extra"><span class="db-abort-extra-inner">，这会清除已抓取数据</span></span>';
            abortButton.onclick = () => {
                localStorage.removeItem(storageKey(CONFIG.dataKey));
                setState({ status: 'idle' });
                location.reload();
            };
            bar.insertBefore(abortButton, button);
        }
        document.body.appendChild(bar);
    }

    function setGridMode(category, url) {
        if (!['movie', 'book', 'music'].includes(category)) return url;
        const next = new URL(url);
        next.searchParams.set('mode', 'grid');
        return next.href;
    }

    function getPageStartUrl(category, pageNumber) {
        const target = new URL(setGridMode(category, location.href));
        target.searchParams.set('start', String(Math.max(0, (pageNumber - 1) * CATEGORIES[category].pageSize)));
        target.searchParams.delete('db_export');
        return target.href;
    }

    function startScraping(category, pageRange) {
        const current = new URL(location.href);
        const target = getPageStartUrl(category, pageRange ? pageRange.startPage : 1);
        setState({ status: 'running', category, pageRange: pageRange || null, startedAt: new Date().toISOString() });
        setStoredData([]);
        if (target !== current.href) {
            location.href = target;
            return;
        }
        processPage(category);
    }

    async function processPage(category) {
        const state = getState();
        if (state.status !== 'running') return;
        const delay = Math.floor(Math.random() * (CONFIG.maxDelay - CONFIG.minDelay) + CONFIG.minDelay);
        setTimeout(async () => {
            if (getState().status !== 'running') return;
            try {
                const pageData = await scrapeCurrentPage(category);
                if (getState().status !== 'running') return;
                const merged = new Map(getStoredData().map(item => [item.link || item.id, item]));
                pageData.forEach(item => merged.set(item.link || item.id, item));
                setStoredData([...merged.values()]);
                const currentPage = getCurrentPageNumber(category);
                const reachedRangeEnd = state.pageRange && currentPage >= state.pageRange.endPage;
                const next = reachedRangeEnd ? '' : getNextPage();
                if (next) location.href = next;
                else {
                    setState({ status: 'paused_for_download', category, pageRange: state.pageRange || null, startedAt: state.startedAt, finishedAt: new Date().toISOString() });
                    showDownloadPanel(category);
                }
            } catch (error) {
                console.error('[Douban Export] 页面解析失败:', error);
                setState({ status: 'error', category, pageRange: state.pageRange || null, message: String(error) });
                alert('本页解析失败，请打开控制台查看错误后重试。');
            }
        }, delay);
    }

    function showDownloadPanel(category) {
        if (document.getElementById('db-export-modal-overlay')) return;
        const data = getStoredData();
        const state = getState();
        const overlay = document.createElement('div');
        overlay.id = 'db-export-modal-overlay';
        overlay.innerHTML = `<div id="db-export-modal">
            <h3>✅ 抓取完成</h3><p style="font-size:16px;text-align:center">共收集到 <b>${data.length}</b> 条${CATEGORIES[category].label}数据（${formatPageRange(state.pageRange)}）</p>
            <div class="db-download-section"><div class="db-download-title">导出数据文件</div><div class="db-download-actions"><button class="db-btn db-btn-primary" id="db-dl-xlsx">📊 导出 Excel (.xlsx)</button><button class="db-btn db-btn-primary" style="background:#2c3e50" id="db-dl-json">🤖 导出 JSON</button></div></div>
            <div class="db-btn-group"><button class="db-btn db-btn-secondary" id="db-close-finish">关闭并清理</button></div>
        </div>`;
        document.body.appendChild(overlay);
        overlay.querySelector('#db-dl-xlsx').onclick = () => generateFile(category, 'xlsx');
        overlay.querySelector('#db-dl-json').onclick = () => generateFile(category, 'json');
        overlay.querySelector('#db-close-finish').onclick = () => {
            localStorage.removeItem(storageKey(CONFIG.dataKey));
            setState({ status: 'idle' });
            location.reload();
        };
    }

    function buildExportItem(item) {
        return {
            title: item.title,
            douban_id: item.id,
            user_rating: item.rating === '' ? null : item.rating,
            mark_date: item.date,
            status: item.status,
            tags: item.tags ? item.tags.split(/\s+/).filter(Boolean) : [],
            comment: item.comment,
            intro: item.intro,
            douban_url: item.link
        };
    }

    function getExportBaseName(category) {
        const statusPart = getCategoryStatusFilePart();
        return `Douban_${CATEGORIES[category].file}${statusPart ? `_${statusPart}` : ''}_Export_${new Date().toISOString().slice(0, 10)}`;
    }

    function buildJsonOutput(category) {
        const data = getStoredData();
        return {
            meta: {
                category,
                category_name: `${CATEGORIES[category].label}（${getCategoryStatusLabel(category)}）`,
                export_date: new Date().toISOString(),
                total_count: data.length,
                source: 'Douban Media Export Tool'
            },
            items: data.map(buildExportItem)
        };
    }

    function buildDoufenRow(item) {
        // 豆坟模板列序：标题 / 简介 / 豆瓣评分 / 链接 / 创建时间 / 我的评分 / 标签 / 评论 / 可见性
        return [
            item.title || null,
            item.intro || null,
            null, // 豆瓣评分（社区评分）脚本未采集，保持为空
            item.link || null,
            item.date || null,
            item.rating === '' ? null : item.rating,
            item.tags ? String(item.tags).split(/\s+/).filter(Boolean).join(',') : null, // 模板标签为逗号分隔
            item.comment || null,
            'public'
        ];
    }

    function buildWorkbook(category) {
        if (typeof XLSX === 'undefined') throw new Error('Excel 组件加载失败，请刷新页面后重试。');
        const data = getStoredData();
        // 与豆坟模板格式一致：保留全部 sheet 与表头，当前分类状态的数据写入对应 sheet，其余 sheet 仅表头。
        const targetSheet = (DOUFEN_STATUS_SHEETS[category] || {})[getStatusFromUrl()] || '';
        const wb = XLSX.utils.book_new();
        DOUFEN_SHEETS.forEach(sheetName => {
            const rows = [DOUFEN_HEADERS.slice()];
            if (sheetName === targetSheet) data.forEach(item => rows.push(buildDoufenRow(item)));
            const ws = XLSX.utils.aoa_to_sheet(rows);
            ws['!cols'] = DOUFEN_COL_WIDTHS.map(width => ({ width }));
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
        });
        return wb;
    }

    function buildXlsxBytes(category) {
        const output = XLSX.write(buildWorkbook(category), { bookType: 'xlsx', type: 'array' });
        return output instanceof Uint8Array ? output : new Uint8Array(output);
    }

    function utf8Bytes(value) {
        return new TextEncoder().encode(value);
    }

    function generateFile(category, format) {
        const data = getStoredData();
        if (!data.length) { alert('无数据'); return; }
        const name = getExportBaseName(category);
        try {
            if (format === 'json') {
                const bytes = utf8Bytes(JSON.stringify(buildJsonOutput(category), null, 2));
                triggerDownload(new Blob([bytes], { type: 'application/json;charset=utf-8' }), `${name}.json`);
                return;
            }
            const bytes = buildXlsxBytes(category);
            triggerDownload(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${name}.xlsx`);
        } catch (error) {
            alert(error.message || error);
        }
    }

    function triggerDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    function init() {
        const context = detectContext();
        if (!context) return;
        if (context === 'profile') {
            renderSummaryButton(context);
            return;
        }
        renderSummaryButton(context);
        const state = getState();
        if (state.status === 'paused_for_download' && state.category === context) {
            showDownloadPanel(context);
        } else if (state.status === 'running' && state.category === context) {
            setTimeout(() => processPage(context), 800);
        } else if (context !== 'generic' && new URL(location.href).searchParams.get('db_export') === '1') {
            setTimeout(() => showConfigPanel(context), 500);
        }
    }

    init();
})();
