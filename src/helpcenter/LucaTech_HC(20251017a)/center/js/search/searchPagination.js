/**
 * 検索結果ページネーション機能モジュール（独自実装版）
 * - ページネーション表示制御
 * - ページ切り替え処理
 * - 複数ページネーションコンテナの連動
 * - pagination.min.jsを使用しない軽量実装
 */

// 定数定義
const PAGINATION_PAGE_SIZE = 10;
const MAX_VISIBLE_PAGES = 7; // ページネーション表示の最大ページ数
const INITIAL_VISIBLE_PAGES = 5; // 最初の方で表示するページ数
const END_OFFSET = 4; // 最後の方の判定オフセット
const SIDE_PAGES = 1; // 中間位置で現在ページの前後に表示するページ数

// セレクタ定数
const SELECTORS = {
    SEARCH_RESULTS: '.searchresults',
    PAGINATION: '#pagination',
    PAGINATION_EXT: '#pagination-ext',
    RESULT_ITEM: '.wSearchResultItem'
};

// モジュール内でのみ使用するページネーション用のソース配列
let sourcesForPaging = [];

// ページネーション状態管理
let paginationState = {
    currentPage: 1,
    totalPages: 1,
    pageSize: PAGINATION_PAGE_SIZE
};

/**
 * 検索結果からデータソースを取得
 * @returns {Array} データソース配列
 */
function collectDataSource() {
    const result = [];
    $(SELECTORS.SEARCH_RESULTS).find(SELECTORS.RESULT_ITEM).each(function () {
        result.push($(this).html());
    });
    return result;
}

/**
 * ページネーション状態を初期化
 * @param {Array} sources - データソース
 * @param {number} initialPage - 初期ページ番号
 */
function initializePaginationState(sources, initialPage = 1) {
    paginationState.totalPages = Math.ceil(sources.length / PAGINATION_PAGE_SIZE);
    paginationState.currentPage = initialPage;
}

/**
 * ページネーションコンテナの表示/非表示を制御
 * @param {boolean} shouldShow - 表示するかどうか
 */
function togglePaginationVisibility(shouldShow) {
    const $pagination = $(SELECTORS.PAGINATION);
    const $paginationExt = $(SELECTORS.PAGINATION_EXT);
    
    if (shouldShow) {
        $pagination.show();
        $paginationExt.show();
    } else {
        $pagination.hide();
        $paginationExt.hide();
    }
}

/**
 * ページネーションを設定する
 */
function setupPagination() {
    sourcesForPaging = [];
    paginationState.currentPage = 1;
    pagination(1);
}

/**
 * ページネーション結果
 * @param {number} [page=1] - 初期ページ番号
 */
function pagination(page = 1) {
    // データソースの取得または再利用
    if (sourcesForPaging.length === 0) {
        sourcesForPaging = collectDataSource();
    }
    const sources = sourcesForPaging;

    // データが存在しない場合は非表示にして終了
    if (sources.length === 0) {
        togglePaginationVisibility(false);
        return;
    }

    // 状態の初期化
    initializePaginationState(sources, page);

    // 初回レンダリング
    renderPage(paginationState.currentPage, sources);
    updateAllPaginationControls(sources);

    // 表示制御
    const shouldShowPagination = sources.length > PAGINATION_PAGE_SIZE;
    togglePaginationVisibility(shouldShowPagination);
}

/**
 * ページボタンのHTMLを生成
 * @param {number} pageNum - ページ番号
 * @param {boolean} isActive - アクティブ状態
 * @returns {string} ページボタンのHTML
 */
function createPageButton(pageNum, isActive = false) {
    const activeClass = isActive ? ' active' : '';
    return `<li class="paginationjs-page J-paginationjs-page${activeClass}" data-num="${pageNum}">
                <a>${pageNum}</a>
             </li>`;
}

/**
 * エリプシス（...）のHTMLを生成
 * @returns {string} エリプシスのHTML
 */
function createEllipsis() {
    return '<li class="paginationjs-ellipsis disabled"><a>...</a></li>';
}

/**
 * 前へ/次へボタンのHTMLを生成
 * @param {string} type - 'prev' または 'next'
 * @param {number|null} pageNum - ページ番号（nullの場合は無効状態）
 * @returns {string} ナビゲーションボタンのHTML
 */
function createNavigationButton(type, pageNum) {
    const className = type === 'prev' ? 'paginationjs-prev' : 'paginationjs-next';
    const jsClassName = type === 'prev' ? 'J-paginationjs-previous' : 'J-paginationjs-next';
    
    if (pageNum === null) {
        return `<li class="${className} disabled"><a></a></li>`;
    }
    
    return `<li class="${className} ${jsClassName}" data-num="${pageNum}">
                <a></a>
             </li>`;
}

/**
 * 全てのページネーションコンテナを更新
 * @param {Array} sources - データソース
 */
function updateAllPaginationControls(sources) {
    renderPaginationControls($(SELECTORS.PAGINATION), sources);
    renderPaginationControls($(SELECTORS.PAGINATION_EXT), sources);
}

/**
 * 表示するページ番号のリストを計算
 * @param {number} currentPage - 現在のページ番号
 * @param {number} totalPages - 総ページ数
 * @returns {Array} 表示するページ番号の配列（エリプシスは null）
 */
function calculatePageNumbers(currentPage, totalPages) {
    if (totalPages <= MAX_VISIBLE_PAGES) {
        // 総ページ数が少ない場合は全て表示
        return Array.from({ length: totalPages }, (_, i) => i + 1);
    }

    if (currentPage <= END_OFFSET) {
        // 最初の方: [1, 2, 3, 4, 5, null, last]
        return [1, 2, 3, 4, 5, null, totalPages];
    }

    if (currentPage >= totalPages - (END_OFFSET - 1)) {
        // 最後の方: [1, null, last-4, last-3, last-2, last-1, last]
        const startPage = totalPages - END_OFFSET;
        return [1, null, ...Array.from({ length: 5 }, (_, i) => startPage + i)];
    }

    // 中間: [1, null, current-1, current, current+1, null, last]
    return [
        1,
        null,
        currentPage - SIDE_PAGES,
        currentPage,
        currentPage + SIDE_PAGES,
        null,
        totalPages
    ];
}

/**
 * 指定ページのデータを表示
 * @param {number} pageNum - ページ番号
 * @param {Array} sources - データソース
 */
function renderPage(pageNum, sources) {
    const startIdx = (pageNum - 1) * PAGINATION_PAGE_SIZE;
    const endIdx = Math.min(startIdx + PAGINATION_PAGE_SIZE, sources.length);
    const pageData = sources.slice(startIdx, endIdx);

    let dataHtml = '<ul>';
    $.each(pageData, function (index, item) {
        dataHtml += '<li class="wSearchResultItem nd-content-search">' + item + '</li>';
    });
    dataHtml += '</ul>';
    $(SELECTORS.SEARCH_RESULTS).html(dataHtml);
}

/**
 * ページネーションコントロールを描画
 * @param {jQuery} container - ページネーションコンテナ
 * @param {Array} sources - データソース
 */
function renderPaginationControls(container, sources) {
    const totalPages = paginationState.totalPages;
    const currentPage = paginationState.currentPage;

    let html = '<div class="paginationjs"><div class="paginationjs-pages"><ul>';

    // 前へボタン
    const prevPage = currentPage > 1 ? currentPage - 1 : null;
    html += createNavigationButton('prev', prevPage);

    // ページ番号を計算して表示
    const pageNumbers = calculatePageNumbers(currentPage, totalPages);
    pageNumbers.forEach(pageNum => {
        if (pageNum === null) {
            html += createEllipsis();
        } else {
            html += createPageButton(pageNum, pageNum === currentPage);
        }
    });

    // 次へボタン
    const nextPage = currentPage < totalPages ? currentPage + 1 : null;
    html += createNavigationButton('next', nextPage);

    html += '</ul></div></div>';

    container.html(html);

    // イベントハンドラーをバインド
    bindPaginationEvents(container, sources);
}

/**
 * ページネーションイベントをバインド
 * @param {jQuery} container - ページネーションコンテナ
 * @param {Array} sources - データソース
 */
function bindPaginationEvents(container, sources) {
    // イベント委譲で全てのクリックイベントを統合処理
    container.off('click').on('click', '.J-paginationjs-page, .J-paginationjs-previous, .J-paginationjs-next', function (e) {
        e.preventDefault();
        
        const $target = $(this);
        
        // 無効状態またはアクティブなページはスキップ
        if ($target.hasClass('disabled') || $target.hasClass('active')) {
            return;
        }

        const pageNum = parseInt($target.attr('data-num'));
        if (!isNaN(pageNum)) {
            goToPage(pageNum, sources);
        }
    });
}

/**
 * 指定ページに移動
 * @param {number} pageNum - ページ番号
 * @param {Array} sources - データソース
 */
function goToPage(pageNum, sources) {
    if (pageNum < 1 || pageNum > paginationState.totalPages) {
        return;
    }

    paginationState.currentPage = pageNum;
    renderPage(pageNum, sources);
    updateAllPaginationControls(sources);
}

/**
 * ページネーションを非表示
 */
function hidePagination() {
    $(SELECTORS.PAGINATION).hide();
    $(SELECTORS.PAGINATION_EXT).hide();
}

/**
 * ページネーション用ソースをリセット
 */
function resetPaginationSource() {
    sourcesForPaging = [];
    paginationState.currentPage = 1;
    paginationState.totalPages = 1;
}
