/**
 * 検索結果表示とハイライト機能モジュール (Vue.js対応版)
 * jQuery依存を除去し、Vue.jsのリアクティブシステムと互換性を持たせています
 * 
 * - 検索結果の表示処理
 * - キーワードハイライト機能
 * - 結果カウント表示
 */

// 定数定義
const CLICK_HANDLER_DELAY = 0;

/**
 * 表示状態管理オブジェクト
 * Vue.jsのリアクティブシステムで使用可能
 */
const displayState = {
    results: [],
    displayedResults: [],
    hasResults: false,
    resultCount: 0
};

/**
 * 検索単語を準備する (Vue.js対応版)
 * jQuery依存を除去し、ネイティブDOM APIまたは引数から取得
 * 
 * @param {string} keyword - 検索キーワード（省略時はDOM要素から取得）
 * @returns {Array} 検索単語配列
 */
function prepareSearchWord(keyword) {
    // キーワードを取得
    if (keyword === undefined) {
        const searchInput = document.getElementById('searchkeyword');
        keyword = searchInput ? searchInput.value : '';
    }
    
    let searchWordTmp = escapeHtml(keyword)
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
 * 検索結果を表示 (Vue.js対応版)
 * @param {string} keyword - 検索キーワード（省略時はDOM要素から取得）
 * @returns {Object} 表示結果情報
 */
function displayResult(keyword) {
    const searchWord = prepareSearchWord(keyword);
    const countAllResult = performSearchAndRender(searchWord);

    // 表示状態を更新
    displayState.resultCount = countAllResult;
    displayState.hasResults = countAllResult > 0;

    // 結果件数の更新
    updateResultCount(countAllResult);

    // 検索単語をハイライト（ネイティブDOM要素を取得）
    const contentElements = document.querySelectorAll(".wSearchContent");
    highlightSearchWord(searchWord, contentElements, "font-weight:bold");

    // ページネーション
    setupPagination();
    
    // 結果表示完了イベントを発火
    window.dispatchEvent(new CustomEvent('displayResultCompleted', {
        detail: {
            resultCount: countAllResult,
            searchWords: searchWord
        }
    }));
    
    return {
        resultCount: countAllResult,
        hasResults: countAllResult > 0,
        searchWords: searchWord
    };
}

/**
 * 検索結果件数を更新 (Vue.js対応版)
 * jQuery依存を除去し、ネイティブDOM APIを使用
 * 
 * @param {number} count - 結果件数
 * @returns {void}
 */
function updateResultCount(count) {
    const hasResults = count > 0;
    
    // count-all要素を更新
    const countAllElement = document.getElementById('count-all');
    if (countAllElement) {
        countAllElement.textContent = count;
    }
    
    // hasresult/noresultの表示切り替え
    const hasResultElements = document.querySelectorAll('.hasresult');
    const noResultElements = document.querySelectorAll('.noresult');
    
    hasResultElements.forEach(el => {
        if (!hasResults) {
            el.classList.add('hidden');
        } else {
            el.classList.remove('hidden');
        }
    });
    
    noResultElements.forEach(el => {
        if (hasResults) {
            el.classList.add('hidden');
        } else {
            el.classList.remove('hidden');
        }
    });
    
    if (hasResults) {
        const resultKeywordElement = document.getElementById('resultkeyword');
        const searchInput = document.getElementById('searchkeyword');
        if (resultKeywordElement && searchInput) {
            resultKeywordElement.textContent = searchInput.value;
        }
    }
}

/**
 * ハイライト結果 (Vue.js対応版)
 * jQuery依存を除去し、ネイティブDOM APIを使用
 * 
 * @param {Array} searchWord - 検索単語配列
 * @param {NodeList|Array} content - ハイライト対象の要素コレクション
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

    // ネイティブDOM APIを使用
    const elements = content instanceof NodeList ? Array.from(content) : 
                     Array.isArray(content) ? content : [content];
    
    elements.forEach(element => {
        if (element && element.innerHTML !== undefined) {
            let html = element.innerHTML;
            Object.values(replacements).forEach(([pattern, replacement]) => {
                html = html.replace(pattern, replacement);
            });
            element.innerHTML = html;
        }
    });
}



/**
 * 左メニューで検索カウントを表示 (Vue.js対応版)
 * jQuery依存を除去し、ネイティブDOM APIを使用
 * 
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

        // ネイティブDOM APIを使用
        const countElement = document.getElementById('count-' + node.id);
        if (countElement) {
            countElement.innerHTML = "(<span class=countnumber>" + node.countItem + "</span>)";
        }
        
        const hasResults = node.countItem > 0;
        const labelSearchIn = document.querySelector("label[for='search-in-" + node.id + "']");
        const labelSearchInAll = document.querySelector("label[for='search-in-all-" + node.id + "']");
        
        if (labelSearchIn) {
            if (hasResults) {
                labelSearchIn.classList.remove('emptyresult');
            } else {
                labelSearchIn.classList.add('emptyresult');
            }
        }
        
        if (labelSearchInAll) {
            if (hasResults) {
                labelSearchInAll.classList.remove('emptyresult');
            } else {
                labelSearchInAll.classList.add('emptyresult');
            }
        }

        // カウント数字クリック時のイベント処理
        setupCountClickHandler();
    }
}

/**
 * カウント数字クリック時のイベントハンドラーを設定 (Vue.js対応版)
 * jQuery依存を除去し、ネイティブDOM APIとイベントリスナーを使用
 * 
 * @returns {void}
 */
function setupCountClickHandler() {
    // 既存のリスナーを削除（重複登録防止）
    const countNumbers = document.querySelectorAll('span>span.countnumber');
    
    countNumbers.forEach(countNumber => {
        // 新しいイベントリスナーを追加（古いものは自動的に置き換え）
        const newCountNumber = countNumber.cloneNode(true);
        countNumber.parentNode.replaceChild(newCountNumber, countNumber);
        
        newCountNumber.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopImmediatePropagation();
            e.stopPropagation();
            
            // すべてのチェックを外す
            const searchInCheckboxes = document.querySelectorAll(".search-in[type='checkbox']");
            const searchInAllCheckboxes = document.querySelectorAll(".search-in-all[type='checkbox']");
            
            searchInCheckboxes.forEach(cb => {
                cb.checked = false;
                const closestDiv = cb.closest('div');
                if (closestDiv) closestDiv.classList.remove('check-new');
            });
            
            searchInAllCheckboxes.forEach(cb => {
                cb.checked = false;
                const closestDiv = cb.closest('div');
                if (closestDiv) closestDiv.classList.remove('check-new');
            });

            // 現在のものをチェック
            setTimeout(() => {
                const countSpan = newCountNumber.closest('span.count');
                if (countSpan) {
                    const id = countSpan.id.replace('count-', '');
                    const searchInAll = document.getElementById('search-in-all-' + id);
                    const searchIn = document.getElementById('search-in-' + id);
                    
                    if (searchInAll) searchInAll.checked = true;
                    if (searchIn) {
                        searchIn.checked = true;
                        // changeイベントを発火
                        searchIn.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                }
            }, CLICK_HANDLER_DELAY);
        });
    });
}

/**
 * 検索結果エリアをクリア (Vue.js対応版)
 * @returns {void}
 */
function clearSearchResults() {
    const searchResults = document.querySelector('.searchresults');
    if (searchResults) {
        searchResults.innerHTML = '';
    }
    resetPaginationSource();
    
    // 表示状態をリセット
    displayState.results = [];
    displayState.displayedResults = [];
    displayState.hasResults = false;
    displayState.resultCount = 0;
}

/**
 * 検索結果の表示状態を更新 (Vue.js対応版)
 * @param {boolean} hasResults - 検索結果があるかどうか
 * @returns {void}
 */
function updateSearchResultsVisibility(hasResults) {
    const hasResultElements = document.querySelectorAll('.hasresult');
    const noResultElements = document.querySelectorAll('.noresult');
    
    hasResultElements.forEach(el => {
        if (!hasResults) {
            el.classList.add('hidden');
        } else {
            el.classList.remove('hidden');
        }
    });
    
    noResultElements.forEach(el => {
        if (hasResults) {
            el.classList.add('hidden');
        } else {
            el.classList.remove('hidden');
        }
    });
}

/**
 * 検索結果表示のUI要素を更新 (Vue.js対応版)
 * @param {number} resultCount - 検索結果件数
 * @param {string} keyword - 検索キーワード
 * @returns {void}
 */
function updateResultsUI(resultCount, keyword) {
    const hasResults = resultCount > 0;
    updateSearchResultsVisibility(hasResults);
    
    if (hasResults) {
        const countAllElement = document.getElementById('count-all');
        const resultKeywordElement = document.getElementById('resultkeyword');
        
        if (countAllElement) countAllElement.textContent = resultCount;
        if (resultKeywordElement) resultKeywordElement.textContent = escapeHtml(keyword);
    }
}

/**
 * ローディング表示を設定 (Vue.js対応版)
 * @returns {void}
 */
function showLoading() {
    const boxClickSearch = document.querySelector('.box-click-search');
    if (boxClickSearch) {
        boxClickSearch.innerHTML = '<div class="loading"><i class="fas fa-spinner fa-spin"></i></div>';
    }
}

/**
 * ヘルプリンクを開く (Vue.js対応版)
 * @param {string} url - 開くURL
 * @param {Event} event - クリックイベント
 * @returns {void}
 */
function openhelplink(url, event) {
    // 検索キーワードを取得・正規化
    const searchInput = document.getElementById('searchkeyword');
    const searchKeywordRaw = searchInput ? searchInput.value : '';
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
            // ネイティブDOM APIを使用
            // const contentElements = newWindow.document.querySelectorAll(".wSearchContent");
            // highlightSearchWord(searchWord, contentElements, "font-weight:bold");
        };
    }
}

/**
 * 表示状態を取得
 * @returns {Object} 表示状態オブジェクト
 */
function getDisplayState() {
    return displayState;
}

/**
 * 表示状態をリセット
 * @returns {void}
 */
function resetDisplayState() {
    displayState.results = [];
    displayState.displayedResults = [];
    displayState.hasResults = false;
    displayState.resultCount = 0;
}