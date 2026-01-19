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

    // 検索単語をハイライト
    highlightSearchWord(searchWord,$(".wSearchContent"),"font-weight:bold");

    // 検索結果リンクのイベントハンドラーを設定（XSS対策）
    setupSearchResultLinkHandlers();

    // ページネーション
    setupPagination();
}

/**
 * 検索単語をハイライト表示する
 * HTMLエンティティを考慮しながら、テキストノードのみをハイライトする
 * @param {Array} searchWord - 検索単語配列
 * @param {jQuery} content - ハイライト対象のjQueryオブジェクト
 * @param {string} style - ハイライトスタイル（CSSインラインスタイル）
 * @returns {void}
 */
function highlightSearchWord(searchWord, content, style) {
    const escapedWords = searchWord.map(word => 
        selectorEscape(word.replace(/[<>]/g, m => m === '<' ? '&lt;' : '&gt;'))
    );

    let highlightWord = escapedWords.join("|");
    highlight.forEach(h => {
        highlightWord = highlightWord.replace(new RegExp(h, "gm"), h);
    });

    content.each(function() {
        let html = $(this).html();
        
        // テキストノードのみをハイライト（HTMLタグ内を除外）
        // HTMLタグの外側のテキストのみにマッチする正規表現を使用
        const regex = new RegExp(`(${highlightWord})(?![^<]*>)`, 'gi');
        html = html.replace(regex, `<font class='keyword' style='${style}'>$1</font>`);
        
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
        $("#resultkeyword").text(escapeHtml(keyword));
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