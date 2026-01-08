/**
 * 検索機能メイン処理（検索ロジック）
 * 注意: searchCatalog.js, searchUI.js, searchDisplay.js が事前に読み込まれている必要があります
 */

/**
 * メイン検索機能
 * 検索フォームの入力値を取得し、検索を実行して結果を表示する
 */
function search() {
    // 検索キーワードを正規化（utils.jsの関数を使用）
    const searchWord = normalizeSearchKeyword($("#searchkeyword").val());
    
    // 検索クエリを構築（大文字小文字を区別しないカスタムセレクタを使用）
    const searchQuery = searchWord.map(word => `:containsNormalized(${word})`).join('');
    
    // 検索を実行し、結果を取得
    const searchCatalogueJs = getSearchCatalogueJs();
    const countAllResult = performSearch(searchCatalogueJs, searchQuery);
    
    // UIを更新
    updateSearchUI(countAllResult, $("#searchkeyword").val());
    
    // カウント表示と結果表示
    getSearchCatalogue().forEach(displayCount);
    displayResult();
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
 * 検索UIを更新
 * @param {number} count - 検索結果数
 * @param {string} keyword - 検索キーワード
 */
function updateSearchUI(count, keyword) {
    $("#count-all").text(count);
    $("#resultkeyword").text(escapeHtml(keyword));
}
