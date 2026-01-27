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
    
    // 検索を実行し、結果を取得
    const searchCatalogueJs = getSearchCatalogueJs();
    const countAllResult = performSearch(searchCatalogueJs, searchWord);
    
    // UIを更新
    updateSearchUI(countAllResult, $("#searchkeyword").val());
    
    // カウント表示と結果表示
    getSearchCatalogue().forEach(displayCount);
    displayResult();
}

/**
 * 検索を実行
 * @param {Array} catalogueJs - 検索カタログデータ
 * @param {Array} searchWords - 正規化された検索キーワード配列
 * @returns {number} 検索結果の総数
 */
function performSearch(catalogueJs, searchWords) {
    let totalCount = 0;
    
    catalogueJs.forEach(catalogue => {
        const allItems = catalogue.searchWords.find(".search_word");
        
        // 各要素をフィルタリング：すべての検索ワードを含むもののみ残す
        const findItems = allItems.filter(function() {
            let text = $(this).text().toLowerCase();
            
            // すべての検索ワードが含まれているかチェック
            return searchWords.every(word => text.includes(word));
        });
        
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
    $("#resultkeyword").text(keyword);
}
