/**
 * 検索結果ページネーション機能モジュール（独自実装版）
 * - ページネーション表示制御
 * - ページ切り替え処理
 * - 複数ページネーションコンテナの連動
 * - pagination.min.jsを使用しない軽量実装
 */

// 定数定義
const PAGINATION_PAGE_SIZE = 10;

// モジュール内でのみ使用するページネーション用のソース配列
let sourcesForPagging = [];

// ページネーション状態管理
let paginationState = {
    currentPage: 1,
    totalPages: 1,
    pageSize: PAGINATION_PAGE_SIZE
};

/**
 * ページネーションを設定する
 * @returns {void}
 */
function setupPagination() {
    sourcesForPagging = [];
    paginationState.currentPage = 1;
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
    const sources = sourcesForPagging.length == 0 ? function () {
        const result = [];
        $('.searchresults').find('.wSearchResultItem').each(function () {
            result.push($(this).html());
        });
        sourcesForPagging = result;
        return result;
    }() : sourcesForPagging;

    if (sources.length) {
        paginationState.totalPages = Math.ceil(sources.length / PAGINATION_PAGE_SIZE);
        paginationState.currentPage = page || 1;

        // 初回レンダリング
        renderPage(paginationState.currentPage, sources);
        renderPaginationControls(container, sources);
        renderPaginationControls(container2, sources);

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
    $('.searchresults').html(dataHtml);
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
    if (currentPage > 1) {
        html += `<li class="paginationjs-prev J-paginationjs-previous" data-num="${currentPage - 1}">
                    <a></a>
                 </li>`;
    } else {
        html += `<li class="paginationjs-prev disabled"><a></a></li>`;
    }

    // ページ番号（常に7個の要素を表示）
    if (totalPages <= 7) {
        // 総ページ数が7以下の場合は全て表示
        for (let i = 1; i <= totalPages; i++) {
            if (i === currentPage) {
                html += `<li class="paginationjs-page J-paginationjs-page active" data-num="${i}">
                            <a>${i}</a>
                         </li>`;
            } else {
                html += `<li class="paginationjs-page J-paginationjs-page" data-num="${i}">
                            <a>${i}</a>
                         </li>`;
            }
        }
    } else {
        // 総ページ数が8以上の場合
        if (currentPage <= 4) {
            // 最初の方: 1 2 3 4 5 ... last
            for (let i = 1; i <= 5; i++) {
                if (i === currentPage) {
                    html += `<li class="paginationjs-page J-paginationjs-page active" data-num="${i}">
                                <a>${i}</a>
                             </li>`;
                } else {
                    html += `<li class="paginationjs-page J-paginationjs-page" data-num="${i}">
                                <a>${i}</a>
                             </li>`;
                }
            }
            html += `<li class="paginationjs-ellipsis disabled"><a>...</a></li>`;
            html += `<li class="paginationjs-page J-paginationjs-page" data-num="${totalPages}">
                        <a>${totalPages}</a>
                     </li>`;
        } else if (currentPage >= totalPages - 3) {
            // 最後の方: 1 ... last-4 last-3 last-2 last-1 last
            html += `<li class="paginationjs-page J-paginationjs-page" data-num="1">
                        <a>1</a>
                     </li>`;
            html += `<li class="paginationjs-ellipsis disabled"><a>...</a></li>`;
            for (let i = totalPages - 4; i <= totalPages; i++) {
                if (i === currentPage) {
                    html += `<li class="paginationjs-page J-paginationjs-page active" data-num="${i}">
                                <a>${i}</a>
                             </li>`;
                } else {
                    html += `<li class="paginationjs-page J-paginationjs-page" data-num="${i}">
                                <a>${i}</a>
                             </li>`;
                }
            }
        } else {
            // 中間: 1 ... current-1 current current+1 ... last
            html += `<li class="paginationjs-page J-paginationjs-page" data-num="1">
                        <a>1</a>
                     </li>`;
            html += `<li class="paginationjs-ellipsis disabled"><a>...</a></li>`;
            for (let i = currentPage - 1; i <= currentPage + 1; i++) {
                if (i === currentPage) {
                    html += `<li class="paginationjs-page J-paginationjs-page active" data-num="${i}">
                                <a>${i}</a>
                             </li>`;
                } else {
                    html += `<li class="paginationjs-page J-paginationjs-page" data-num="${i}">
                                <a>${i}</a>
                             </li>`;
                }
            }
            html += `<li class="paginationjs-ellipsis disabled"><a>...</a></li>`;
            html += `<li class="paginationjs-page J-paginationjs-page" data-num="${totalPages}">
                        <a>${totalPages}</a>
                     </li>`;
        }
    }

    // 次へボタン
    if (currentPage < totalPages) {
        html += `<li class="paginationjs-next J-paginationjs-next" data-num="${currentPage + 1}">
                    <a></a>
                 </li>`;
    } else {
        html += `<li class="paginationjs-next disabled"><a></a></li>`;
    }

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
    // ページ番号クリック
    container.off('click').on('click', '.J-paginationjs-page', function (e) {
        e.preventDefault();
        const $target = $(this);
        if ($target.hasClass('disabled') || $target.hasClass('active')) {
            return;
        }

        const pageNum = parseInt($target.attr('data-num'));
        goToPage(pageNum, sources);
    });

    // 前へボタン
    container.on('click', '.J-paginationjs-previous', function (e) {
        e.preventDefault();
        const $target = $(this);
        if ($target.hasClass('disabled')) {
            return;
        }

        const pageNum = parseInt($target.attr('data-num'));
        goToPage(pageNum, sources);
    });

    // 次へボタン
    container.on('click', '.J-paginationjs-next', function (e) {
        e.preventDefault();
        const $target = $(this);
        if ($target.hasClass('disabled')) {
            return;
        }

        const pageNum = parseInt($target.attr('data-num'));
        goToPage(pageNum, sources);
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
    renderPaginationControls($('#pagination'), sources);
    renderPaginationControls($('#pagination-ext'), sources);
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
    paginationState.currentPage = 1;
    paginationState.totalPages = 1;
}
