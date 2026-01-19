/**
 * パンくずリスト関連モジュール
 * - パンくずリストの構築
 * - パンくず情報の取得と作成
 */

/**
 * 検索結果でパンくずリストを構築
 * @param {Array} breadCrumArr - パンくず配列
 * @returns {string} パンくずHTMLストリング
 */
function buildBreadCrum(breadCrumArr) {
    const breadCrumItems = breadCrumArr.map(item => 
        `<li class="breadcrumb-item">${escapeHtml(item)}</li>`
    ).join('');
    
    return `<ol class="breadcrumb">${breadCrumItems}</ol>`;
}

/**
 * パンくずリスト情報を取得
 * @param {Event} event - クリックイベント
 * @returns {Array<string>} パンくずリストのテキスト配列
 */
function getBreadcrumbTexts(event) {
    return $(event.target)
        .closest('.wSearchResultItem')
        .find('.breadcrumb-item')
        .map(function () {
            return $(this).text();
        })
        .get();
}

/**
 * パンくず情報オブジェクトを作成
 * @param {string} url - ヘルプページのURL
 * @param {Array<string>} breadcrumbTexts - パンくずテキスト配列
 * @returns {Object} パンくず情報オブジェクト
 */
function createBreadcrumbData(url, breadcrumbTexts) {
    return {
        path: escapeHtml(url),
        indexType: "search",
        contentid: "",
        categoryTitle: breadcrumbTexts.length > 0 ? escapeHtml(breadcrumbTexts[0]) : "",
        subCategoryTitle: breadcrumbTexts.length > 2 ? escapeHtml(breadcrumbTexts[1]) : "",
        contentsTitle: breadcrumbTexts.length > 2 
            ? escapeHtml(breadcrumbTexts[2]) 
            : breadcrumbTexts.length === 2 
                ? escapeHtml(breadcrumbTexts[1]) 
                : ""
    };
}
