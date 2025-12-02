/**
 * 検索結果ページネーション機能モジュール (Vue.js対応版)
 * jQuery依存を減らし、Vue.jsのリアクティブシステムと互換性を持たせています
 * 
 * 注意: このファイルはjQuery paginationプラグインに依存しています
 * 将来的にはネイティブ実装またはVueコンポーネントへの移行を推奨します
 * 
 * - ページネーション表示制御
 * - ページ切り替え処理
 * - 複数ページネーションコンテナの連動
 */

// 定数定義
const PAGINATION_PAGE_SIZE = 10;

// モジュール内でのみ使用するページネーション用のソース配列
let sourcesForPagging = [];

/**
 * ページネーション状態管理オブジェクト
 * Vue.jsのリアクティブシステムで使用可能
 */
const paginationState = {
    currentPage: 1,
    totalPages: 0,
    pageSize: PAGINATION_PAGE_SIZE,
    totalItems: 0,
    isVisible: false
};

/**
 * ページネーションを設定する (Vue.js対応版)
 * @returns {void}
 */
function setupPagination() {
    sourcesForPagging = [];
    paginationState.currentPage = 1;
    
    // ネイティブDOM APIでコンテナを取得
    const paginationContainer = document.getElementById('pagination');
    const paginationExtContainer = document.getElementById('pagination-ext');
    
    // jQueryプラグインが利用可能な場合は使用
    if (typeof $ !== 'undefined' && paginationContainer && paginationExtContainer) {
        pagination($('#pagination'), $('#pagination-ext'), 1);
    } else {
        console.warn('setupPagination: jQuery pagination plugin not available');
    }
}

/**
 * ページネーション結果 (jQuery依存あり)
 * 
 * 注意: この関数はjQuery paginationプラグインに依存しています
 * 将来的にはネイティブ実装またはVueコンポーネントへの置き換えを推奨
 * 
 * @param {jQuery} container - ページネーションコンテナ
 * @param {jQuery} container2 - 追加のページネーションコンテナ
 * @param {number} [page] - 初期ページ番号
 * @returns {void}
 */
function pagination(container, container2, page) {
    // ソースデータを準備（初回のみDOM から取得）
    const sources = sourcesForPagging.length === 0 ? function () {
        const result = [];
        const searchResults = document.querySelectorAll('.searchresults .wSearchResultItem');
        searchResults.forEach(item => {
            result.push(item.innerHTML);
        });
        sourcesForPagging = result;
        
        // ページネーション状態を更新
        paginationState.totalItems = result.length;
        paginationState.totalPages = Math.ceil(result.length / PAGINATION_PAGE_SIZE);
        
        return result;
    }() : sourcesForPagging;

    if (sources.length) {
        const options = {
            dataSource: sources,
            pageSize: PAGINATION_PAGE_SIZE,
            prevText: "",
            nextText: "",
            callback: function (response, pagination) {
                // ページ変更時のコールバック（ネイティブDOM APIを使用）
                let dataHtml = '<ul>';
                response.forEach((item, index) => {
                    dataHtml += '<li class="wSearchResultItem nd-content-search">' + item + '</li>';
                });
                dataHtml += '</ul>';
                
                const searchResults = document.querySelector('.searchresults');
                if (searchResults) {
                    searchResults.innerHTML = dataHtml;
                }
                
                // ページネーション状態を更新
                paginationState.currentPage = pagination.pageNumber;
            }
        };
        // 共通のフック処理とページネーション初期化
        const initializePagination = (paginationContainer) => {
            paginationContainer.addHook('beforeInit', function () {
                window.console && console.log('beforeInit...');
            });
            paginationContainer.addHook('beforePageOnClick', function () {
                window.console && console.log('beforePageOnClick...');
                //return false
            });
            paginationContainer.pagination(options);
        };

        initializePagination(container);
        initializePagination(container2);
        
        // afterPagingフックを追加して2つのページネーションを連動
        container.addHook('afterPaging', function () {
            if (container.pagination('getSelectedPageNum') != container2.pagination('getSelectedPageNum')) {
                container2.pagination('go', container.pagination('getSelectedPageNum'));
            }
        });
        
        container2.addHook('afterPaging', function () {
            if (container2.pagination('getSelectedPageNum') != container.pagination('getSelectedPageNum')) {
                container.pagination('go', container2.pagination('getSelectedPageNum'));
            }
        });

        // 表示/非表示の制御（ネイティブDOM API）
        if (sources.length <= PAGINATION_PAGE_SIZE) {
            const paginationContainer = document.getElementById('pagination');
            const paginationExtContainer = document.getElementById('pagination-ext');
            if (paginationContainer) paginationContainer.style.display = 'none';
            if (paginationExtContainer) paginationExtContainer.style.display = 'none';
            paginationState.isVisible = false;
        } else {
            const paginationContainer = document.getElementById('pagination');
            const paginationExtContainer = document.getElementById('pagination-ext');
            if (paginationContainer) paginationContainer.style.display = 'block';
            if (paginationExtContainer) paginationExtContainer.style.display = 'block';
            paginationState.isVisible = true;
        }
    } else {
        const paginationContainer = document.getElementById('pagination');
        const paginationExtContainer = document.getElementById('pagination-ext');
        if (paginationContainer) paginationContainer.style.display = 'none';
        if (paginationExtContainer) paginationExtContainer.style.display = 'none';
        paginationState.isVisible = false;
    }
}

/**
 * ページネーションを非表示 (Vue.js対応版)
 * jQuery依存を除去し、ネイティブDOM APIを使用
 * @returns {void}
 */
function hidePagination() {
    const paginationContainer = document.getElementById('pagination');
    const paginationExtContainer = document.getElementById('pagination-ext');
    
    if (paginationContainer) {
        paginationContainer.style.display = 'none';
    }
    if (paginationExtContainer) {
        paginationExtContainer.style.display = 'none';
    }
    
    paginationState.isVisible = false;
}

/**
 * ページネーションを表示 (Vue.js対応版)
 * @returns {void}
 */
function showPagination() {
    const paginationContainer = document.getElementById('pagination');
    const paginationExtContainer = document.getElementById('pagination-ext');
    
    if (paginationContainer) {
        paginationContainer.style.display = 'block';
    }
    if (paginationExtContainer) {
        paginationExtContainer.style.display = 'block';
    }
    
    paginationState.isVisible = true;
}

/**
 * ページネーション用ソースをリセット
 * @returns {void}
 */
function resetPaginationSource() {
    sourcesForPagging = [];
    paginationState.currentPage = 1;
    paginationState.totalPages = 0;
    paginationState.totalItems = 0;
}

/**
 * ページネーション状態を取得
 * @returns {Object} ページネーション状態オブジェクト
 */
function getPaginationState() {
    return paginationState;
}
