/**
 * 検索機能メイン処理（検索ロジック - Vue.js対応版）
 * jQuery依存を除去し、Vue.jsのリアクティブシステムと互換性を持たせています
 * 
 * 注意: searchCatalog.js, searchUI.js, searchDisplay.js が事前に読み込まれている必要があります
 */

/**
 * 検索状態管理オブジェクト
 * Vue.jsのリアクティブシステムで使用可能
 */
const searchState = {
    keyword: '',
    normalizedKeywords: [],
    results: [],
    totalCount: 0,
    isSearching: false
};

/**
 * メイン検索機能 (Vue.js対応版)
 * 検索フォームの入力値を取得し、検索を実行して結果を表示する
 * 
 * @param {string} keyword - 検索キーワード（省略時はDOM要素から取得）
 * @returns {Object} 検索結果 { keyword, totalCount, results }
 */
function search(keyword) {
    // 検索キーワードを取得
    if (keyword === undefined) {
        // jQuery依存を除去: ネイティブDOM APIを使用
        const searchInput = document.getElementById('searchkeyword');
        keyword = searchInput ? searchInput.value : '';
    }
    
    // 検索状態を更新
    searchState.keyword = keyword;
    searchState.isSearching = true;
    
    // 検索キーワードを正規化（utils.jsの関数を使用）
    const searchWord = normalizeSearchKeyword(keyword);
    searchState.normalizedKeywords = searchWord;
    
    // 検索クエリを構築
    const searchQuery = searchWord.map(word => `:contains(${word})`).join('');
    
    // 検索を実行し、結果を取得
    const searchCatalogueJs = getSearchCatalogueJs();
    const countAllResult = performSearch(searchCatalogueJs, searchQuery);
    
    // 検索状態を更新
    searchState.totalCount = countAllResult;
    searchState.isSearching = false;
    
    // UIを更新
    updateSearchUI(countAllResult, keyword);
    
    // カウント表示と結果表示
    getSearchCatalogue().forEach(displayCount);
    displayResult();
    
    // 検索完了イベントを発火（Vueコンポーネントで監視可能）
    window.dispatchEvent(new CustomEvent('searchCompleted', {
        detail: {
            keyword,
            totalCount: countAllResult,
            normalizedKeywords: searchWord
        }
    }));
    
    // 検索結果を返す
    return {
        keyword,
        totalCount: countAllResult,
        results: searchCatalogueJs.map(cat => ({
            id: cat.id,
            title: cat.title,
            count: cat.findItems ? cat.findItems.length : 0
        }))
    };
}

/**
 * 検索を実行
 * @param {Array} catalogueJs - 検索カタログデータ
 * @param {string} searchQuery - 検索クエリ
 * @returns {number} 検索結果の総数
 */
function performSearch(catalogueJs, searchQuery) {
    let totalCount = 0;
    
    catalogueJs.forEach(catalogue => {
        const findItems = catalogue.searchWords.find(".search_word" + searchQuery);
        catalogue.findItems = findItems;
        totalCount += findItems.length;
    });
    
    return totalCount;
}

/**
 * 検索UIを更新 (Vue.js対応版)
 * jQuery依存を除去し、ネイティブDOM APIを使用
 * 
 * @param {number} count - 検索結果数
 * @param {string} keyword - 検索キーワード
 */
function updateSearchUI(count, keyword) {
    // jQuery依存を除去: ネイティブDOM APIを使用
    const countAllElement = document.getElementById('count-all');
    if (countAllElement) {
        countAllElement.textContent = count;
    }
    
    const resultKeywordElement = document.getElementById('resultkeyword');
    if (resultKeywordElement) {
        resultKeywordElement.textContent = escapeHtml(keyword);
    }
}

/**
 * 検索状態を取得
 * @returns {Object} 検索状態オブジェクト
 */
function getSearchState() {
    return searchState;
}

/**
 * 検索状態をリセット
 */
function resetSearchState() {
    searchState.keyword = '';
    searchState.normalizedKeywords = [];
    searchState.results = [];
    searchState.totalCount = 0;
    searchState.isSearching = false;
}
