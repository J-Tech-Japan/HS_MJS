/**
 * 検索結果ページネーション機能モジュール
 * - ページネーション表示制御
 * - ページ切り替え処理
 * - 複数ページネーションコンテナの連動
 */

// 定数定義
const PAGINATION_PAGE_SIZE = 10;

// モジュール内でのみ使用するページネーション用のソース配列
let sourcesForPagging = [];

/**
 * ページネーションを設定する
 * @returns {void}
 */
function setupPagination() {
    sourcesForPagging = [];
    pagination($('#pagination'), $('#pagination-ext'), 1);
}

/**
 * ページネーション結果
 * @param {jQuery} container - ページネーションコンテナ
 * @param {jQuery} container2 - 追加のページネーションコンテナ
 * @param {number} [page] - 初期ページ番号
 * @returns {void}
 */
function pagination(container, container2, page) {
    // ページネーション
    // const container = $('.pagination-container');
    const sources = sourcesForPagging.length==0?function () {
        const result = [];
        $('.searchresults').find('.wSearchResultItem').each(function () {
            result.push($(this).html());
        });
        sourcesForPagging=result;
        return result;
    }():sourcesForPagging;

    if (sources.length) {
        const options = {
            dataSource: sources,
            pageSize: PAGINATION_PAGE_SIZE,
            prevText: "",
            nextText: "",
            callback: function (response, pagination) {
                let dataHtml = '<ul>';
                $.each(response, function (index, item) {
                    dataHtml += '<li class="wSearchResultItem nd-content-search">' + item + '</li>';
                });
                dataHtml += '</ul>';
                $('.searchresults').html(dataHtml);
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

        if (sources.length <= PAGINATION_PAGE_SIZE) {
            container.hide();
            container2.hide();
        } else {
            container.show();
            container2.show();
        }
    } else {
        container.hide();
        container2.hide();
    }
}

/**
 * ページネーションを非表示
 * @returns {void}
 */
function hidePagination() {
    $('#pagination, #pagination-ext').hide();
}

/**
 * ページネーション用ソースをリセット
 * @returns {void}
 */
function resetPaginationSource() {
    sourcesForPagging = [];
}
