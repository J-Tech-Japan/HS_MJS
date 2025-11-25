/**
 * 検索結果表示とハイライト機能モジュール
 * - 検索結果の表示処理
 * - キーワードハイライト機能
 * - 結果カウント表示
 */

// 定数定義
const CLICK_HANDLER_DELAY = 0;

/**
 * 検索単語を準備する
 * @returns {Array} 検索単語配列
 */
function prepareSearchWord() {
    let searchWordTmp = escapeHtml($("#searchkeyword").val())
        .replace(/(.*?)(?:　| )+(.*?)/g, "$1 $2")
        .trim()
        .toLowerCase();
    
    wide.forEach((w, i) => {
        searchWordTmp = searchWordTmp.split(w).join(narrow[i]);
    });
    
    return searchWordTmp.split(" ");
}

/**
 * 検索を実行し結果カウントを取得する
 * @param {Array} searchWord - 検索単語配列
 * @returns {number} 総結果数
 */
function performSearchAndRender(searchWord) {
    let countAllResult = 0;
    const $searchResults = $(".searchresults");
    $searchResults.empty();

    const searchCatalogueJs = getSearchCatalogueJs();
    for (let searchCatalogueJsCount = 0; searchCatalogueJsCount < searchCatalogueJs.length; searchCatalogueJsCount++) {
        const searchCatalogueItemChild = searchCatalogueJs[searchCatalogueJsCount];
        if ($("#search-in-" + searchCatalogueItemChild.id).is(":checked")) {
            searchCatalogueItemChild.findItems.each(function () {
                const breadCrum = searchCatalogueItemChild.breadCrum;
                let displayText = $(this).parent().text();

                displayText = searchKeywordsInString(displayText, searchWord); 
                if (displayText=="") {
                    displayText=$(this).parent().find(".displayText").text();
                }

                // 安全にHTMLを構築してエスケープされたコンテンツを使用
                const safeTitle = escapeHtml($(this).parent().find(".search_title").text());
                const safeDisplayText = escapeHtml(displayText);
                const safeUrl = escapeHtml(searchCatalogueItemChild.baseUrl.replace(/\/$/, "") + "/index.html#t=" + $(this).parent().attr("id") + ".html");
                
                $searchResults.append($("<div class='wSearchResultItem'><div class='wSearchResultTitle title-s'><a class='nolink' href='#' onclick='openhelplink(\"" + safeUrl + "\", event);return false;'>" + safeTitle + "</a></div><div class='wSearchResultBreadCrum'>" + buildBreadCrum(breadCrum) + "</div><div class='wSearchContent'><span class='wSearchContext nd-p'>" + safeDisplayText + "</span></div></div>"));
            });
            countAllResult += searchCatalogueItemChild.findItems.length;
        }
    }
    
    return countAllResult;
}



/**
 * 検索結果を表示
 * @returns {void}
 */
function displayResult() {
    const searchWord = prepareSearchWord();
    const countAllResult = performSearchAndRender(searchWord);

    // 結果件数の更新
    updateResultCount(countAllResult);

    // 検索単語をハイライト
    highlightSearchWord(searchWord,$(".wSearchContent"),"font-weight:bold");

    // ページネーション
    setupPagination();
}

/**
 * 検索結果件数を更新
 * @param {number} count - 結果件数
 * @returns {void}
 */
function updateResultCount(count) {
    const hasResults = count > 0;
    
    $("#count-all").text(count);
    $(".hasresult").toggleClass("hidden", !hasResults);
    $(".noresult").toggleClass("hidden", hasResults);
    
    if (hasResults) {
        $("#resultkeyword").text($("#searchkeyword").val());
    }
}

/**
 * ハイライト結果
 * @param {Array} searchWord - 検索単語配列
 * @param {jQuery} content - ハイライト対象のjQueryオブジェクト
 * @param {string} style - ハイライトスタイル
 * @returns {void}
 */
function highlightSearchWord(searchWord, content, style) {
    const escapedWords = searchWord.map(word => 
        selectorEscape(word.replace(">", "&gt;").replace("<", "&lt;"))
    );

    let hilightWord = escapedWords.join("|");
    hilight.forEach(h => {
        hilightWord = hilightWord.replace(new RegExp(h, "gm"), h);
    });

    const replacements = {
        regnbsp: [/&nbsp;(?=[^<>]*<)/gm, "　"],
        reggt: [/&gt;(?=[^<>]*<)/gm, ">"],
        reglt: [/&lt;(?=[^<>]*<)/gm, "<"],
        regquot: [/&quot;(?=[^<>]*<)/gm, '"'],
        regamp: [/&amp;(?=[^<>]*<)/gm, "&"],
        reghighlight: [new RegExp(`(${hilightWord})`, 'g'), `<font class='keyword' style='${style}'>$1</font>`]
    };

    content.each(function() {
        let html = $(this).html();
        Object.values(replacements).forEach(([pattern, replacement]) => {
            html = html.replace(pattern, replacement);
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
    if (node) {
        if (node.findItems) {
            node.countItem = node.findItems.length;
        } else {
            node.countItem = 0;
        }
        if (node.childs) {
            for (let i = 0; i < node.childs.length; i++) {
                displayCount(node.childs[i]);
                node.countItem += node.childs[i].countItem;
            }
        }

        $("#count-" + node.id).html("(<span class=countnumber>" + node.countItem + "</span>)");
        if(node.countItem==0){
            $("label[for='search-in-" + node.id+"']").addClass("emptyresult");
            $("label[for='search-in-all-" + node.id+"']").addClass("emptyresult");
        }else{
            $("label[for='search-in-" + node.id+"']").removeClass("emptyresult");
            $("label[for='search-in-all-" + node.id+"']").removeClass("emptyresult");
        }

        // カウント数字クリック時のイベント処理
        setupCountClickHandler();
    }
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
        $(".search-in[type='checkbox']").prop("checked", false);
        $(".search-in-all[type='checkbox']").prop("checked", false);
        $(".search-in[type='checkbox']").closest("div").removeClass("check-new");
        $(".search-in-all[type='checkbox']").closest("div").removeClass("check-new");

        const self = $(this);
        // 現在のものをチェック
        setTimeout(function(){
            const id = self.closest("span.count").attr("id").replace("count-", "");
            $("#search-in-all-" + id).prop("checked", true);
            $("#search-in-" + id).prop("checked", true);
            $("#search-in-" + id).trigger("change");
        }, CLICK_HANDLER_DELAY);
    });
}

/**
 * 検索結果エリアをクリア
 * @returns {void}
 */
function clearSearchResults() {
    const $searchResults = $(".searchresults");
    $searchResults.empty();
    resetPaginationSource();
}



/**
 * 検索結果の表示状態を更新
 * @param {boolean} hasResults - 検索結果があるかどうか
 * @returns {void}
 */
function updateSearchResultsVisibility(hasResults) {
    $(".hasresult").toggleClass("hidden", !hasResults);
    $(".noresult").toggleClass("hidden", hasResults);
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
    
    // ローカルストレージに保存
    localStorage.setItem("breadcrumb", JSON.stringify(breadcrumb));
    
    // 新しいウィンドウでヘルプを開く
    const newWindow = window.open(url, "_blank");
    
    // 将来的な検索ワードハイライト機能のための準備
    if (newWindow) {
        newWindow.onload = function() {
            // TODO: 検索ワードのハイライト機能を実装
            // hightLightSearchWord(searchWord, $(".wSearchContent"));
        };
    }
}