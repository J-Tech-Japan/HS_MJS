/**
 * 検索結果表示とハイライト機能モジュール
 * - 検索結果の表示処理
 * - キーワードハイライト機能
 * - 結果カウント表示
 * 
 * 依存関係:
 * - utils.js: normalizeSearchKeyword, escapeHtml, selectorEscape, highlight配列
 * - searchBreadcrumb.js: buildBreadCrum
 * - searchPagination.js: setupPagination, resetPaginationSource
 */

/**
 * 検索単語を準備する
 * utils.jsのnormalizeSearchKeyword関数を使用して正規化を行う
 * @returns {Array} 検索単語配列
 */
function prepareSearchWord() {
    return normalizeSearchKeyword($("#searchkeyword").val());
}

/**
 * 検索を実行し結果をレンダリングする
 * @param {Array} searchWord - 検索単語配列
 * @returns {number} 総結果数（レンダリングされた結果の件数）
 */
function performSearchAndRender(searchWord) {
    let countAllResult = 0;
    const $searchResults = $(".searchresults");
    $searchResults.empty();

    const catalogues = getSearchCatalogueJs();
    catalogues.forEach(catalogue => {
        if (!$(`#search-in-${catalogue.id}`).is(":checked")) return;
        
        catalogue.findItems.each(function () {
            const $parent = $(this).parent();
            const displayText = searchKeywordsInString($parent.text(), searchWord) || $parent.find(".displayText").text();
            const safeTitle = escapeHtml($parent.find(".search_title").text());
            const safeDisplayText = escapeHtml(displayText);
            const baseUrl = catalogue.baseUrl.replace(/\/$/, "");
            const url = `${baseUrl}/index.html#t=${$parent.attr("id")}.html`;
            
            const resultItem = `
                <div class='wSearchResultItem'>
                    <div class='wSearchResultTitle title-s'>
                        <a class='nolink search-result-link' href='#' data-url="${escapeHtml(url)}">${safeTitle}</a>
                    </div>
                    <div class='wSearchResultBreadCrum'>${buildBreadCrum(catalogue.breadCrum)}</div>
                    <div class='wSearchContent'><span class='wSearchContext nd-p'>${safeDisplayText}</span></div>
                </div>`;
            $searchResults.append(resultItem);
        });
        countAllResult += catalogue.findItems.length;
    });
    
    return countAllResult;
}

/**
 * 検索結果を表示
 * @returns {void}
 */
function displayResult() {
    const searchWord = prepareSearchWord();
    const countAllResult = performSearchAndRender(searchWord);
    const keyword = $("#searchkeyword").val();

    // 結果件数の更新
    updateResultsUI(countAllResult, keyword);

    // 検索結果リンクのイベントハンドラーを設定（XSS対策）
    setupSearchResultLinkHandlers();

    // ページネーション（ハイライトの前に実行）
    setupPagination();
    
    // 検索単語をハイライト（ページネーション後に実行）
    highlightSearchWord(searchWord,$(".wSearchContent"),"font-weight:bold");
}

/**
 * 検索単語をハイライト表示する
 * HTMLエンティティを考慮しながら、テキストノードのみをハイライトする
 * @param {Array} searchWord - 検索単語配列（正規化済み）
 * @param {jQuery} content - ハイライト対象のjQueryオブジェクト
 * @param {string} style - ハイライトスタイル（CSSインラインスタイル）
 * @returns {void}
 */
function highlightSearchWord(searchWord, content, style) {
    content.each(function() {
        let html = $(this).html();
        
        // 既存のハイライトタグを削除してテキストのみを抽出
        html = html.replace(/<font class='keyword'[^>]*>(.*?)<\/font>/gi, '$1');
        
        // 各検索ワードをハイライト
        searchWord.forEach(word => {
            // 正規表現の特殊文字をエスケープ
            const escapedWord = word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            
            // HTMLエンティティ化されたパターンを作成（& を &amp; にもマッチ）
            let pattern = escapedWord.replace(/&/g, '(?:&|&amp;)');
            
            // 全角英数字と半角英数字の両方にマッチ
            highlight.forEach(h => {
                pattern = pattern.replace(new RegExp(h, "gm"), h);
            });
            
            // 大文字小文字を区別せず、HTMLタグの外側のテキストのみにマッチ
            const regex = new RegExp(`(${pattern})(?![^<]*>)`, 'gi');
            html = html.replace(regex, `<font class='keyword' style='${style}'>$1</font>`);
        });
        
        $(this).html(html);
    });
}

/**
 * 左メニューで検索カウントを表示
 * @param {Object} node - カウント表示対象のノード
 * @returns {void}
 */
function displayCount(node) {
    if (!node) return;
    
    node.countItem = node.findItems ? node.findItems.length : 0;
    
    if (node.childs) {
        node.childs.forEach(child => {
            displayCount(child);
            node.countItem += child.countItem;
        });
    }

    $(`#count-${node.id}`).html(`(<span class=countnumber>${node.countItem}</span>)`);
    
    const method = node.countItem === 0 ? "addClass" : "removeClass";
    $(`label[for='search-in-${node.id}'], label[for='search-in-all-${node.id}']`)[method]("emptyresult");

    setupCountClickHandler();
}

/**
 * カウント数字クリック時のイベントハンドラーを設定
 * @returns {void}
 */
function setupCountClickHandler() {
    $('span>span.countnumber').off('click').click(function(e){
        e.preventDefault();
        e.stopImmediatePropagation();
        e.stopPropagation();
        
        // すべてのチェックを外す
        $(".search-in[type='checkbox'], .search-in-all[type='checkbox']")
            .prop("checked", false)
            .closest("div").removeClass("check-new");

        const id = $(this).closest("span.count").attr("id").replace("count-", "");
        $(`#search-in-all-${id}, #search-in-${id}`).prop("checked", true);
        $(`#search-in-${id}`).trigger("change");
    });
}

/**
 * 検索結果エリアをクリア
 * @returns {void}
 */
function clearSearchResults() {
    $(".searchresults").empty();
    resetPaginationSource();
}

/**
 * 検索結果表示のUI要素を更新（searchUI.jsから移行）
 * @param {number} resultCount - 検索結果件数
 * @param {string} keyword - 検索キーワード
 * @returns {void}
 */
function updateResultsUI(resultCount, keyword) {
    const hasResults = resultCount > 0;
    $(".hasresult").toggleClass("hidden", !hasResults);
    $(".noresult").toggleClass("hidden", hasResults);
    if (hasResults) {
        $("#count-all").text(resultCount);
        $("#resultkeyword").text(keyword);
    }
}

/**
 * ローディング表示を設定（searchUI.jsから移行）
 * @returns {void}
 */
function showLoading() {
    $('.box-click-search').html('<div class="loading"><i class="fas fa-spinner fa-spin"></i></div>');
}

/**
 * ヘルプリンクを開く（searchUI.jsから移行）
 * @param {string} url - 開くURL
 * @param {Event} event - クリックイベント
 * @returns {void}
 */
function openhelplink(url, event) {
    // 検索キーワードを取得・正規化
    const searchKeywordRaw = $("#searchkeyword").val();
    const searchWord = normalizeSearchKeyword(searchKeywordRaw);
    
    // パンくずリスト情報を取得
    const breadcrumbTexts = getBreadcrumbTexts(event);
    const breadcrumb = createBreadcrumbData(url, breadcrumbTexts);
    
    // 検索キーワードをブレッドクラムに追加
    breadcrumb.searchKeyword = searchKeywordRaw;
    
    // ローカルストレージに保存
    localStorage.setItem("breadcrumb", JSON.stringify(breadcrumb));
    
    // 新しいウィンドウでヘルプを開く
    window.open(url, "_blank");
}

/**
 * 検索結果リンクのイベント委譲ハンドラーを設定
 * XSS脆弱性を防ぐため、onclick属性の代わりにイベント委譲を使用
 * @returns {void}
 */
function setupSearchResultLinkHandlers() {
    // 既存のハンドラーを削除して重複を防ぐ
    $(document).off('click', '.search-result-link');
    
    // イベント委譲でクリックイベントを処理
    $(document).on('click', '.search-result-link', function(event) {
        event.preventDefault();
        const url = $(this).data('url');
        if (url) {
            openhelplink(url, event);
        }
    });
}

/**
 * 現在の検索キーワードで表示中のコンテンツを再ハイライト
 * ページネーション切り替え時などに使用
 * @returns {void}
 */
function reapplyHighlight() {
    const searchWord = prepareSearchWord();
    if (searchWord && searchWord.length > 0 && searchWord[0]) {
        highlightSearchWord(searchWord, $(".wSearchContent"), "font-weight:bold");
    }
}