/**
 * パンくずリスト関連モジュール (Vue.js対応版)
 * jQuery依存を減らし、Vue.jsのリアクティブシステムと互換性を持たせています
 * 
 * - パンくずリストの構築
 * - パンくず情報の取得と作成
 */

/**
 * パンくずリスト状態管理オブジェクト
 * Vue.jsのリアクティブシステムで使用可能
 */
const breadcrumbState = {
    currentBreadcrumbs: [],
    lastBreadcrumbData: null
};

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
 * パンくずリスト情報を取得 (Vue.js対応版)
 * jQuery依存を除去し、ネイティブDOM APIを使用
 * @param {Event} event - クリックイベント
 * @returns {Array<string>} パンくずリストのテキスト配列
 */
function getBreadcrumbTexts(event) {
    // クリックされた要素から最も近い.wSearchResultItem要素を取得
    const resultItem = event.target.closest('.wSearchResultItem');
    
    if (!resultItem) {
        return [];
    }
    
    // .breadcrumb-item要素を全て取得してテキストを配列に変換
    const breadcrumbItems = resultItem.querySelectorAll('.breadcrumb-item');
    const breadcrumbTexts = Array.from(breadcrumbItems).map(function(item) {
        return item.textContent;
    });
    
    // 状態を更新
    breadcrumbState.currentBreadcrumbs = breadcrumbTexts;
    
    return breadcrumbTexts;
}

/**
 * パンくず情報オブジェクトを作成 (Vue.js対応版)
 * @param {string} url - ヘルプページのURL
 * @param {Array<string>} breadcrumbTexts - パンくずテキスト配列
 * @returns {Object} パンくず情報オブジェクト
 */
function createBreadcrumbData(url, breadcrumbTexts) {
    const breadcrumbData = {
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
    
    // 状態を更新
    breadcrumbState.lastBreadcrumbData = breadcrumbData;
    
    return breadcrumbData;
}

/**
 * パンくず状態を取得
 * @returns {Object} パンくず状態オブジェクト
 */
function getBreadcrumbState() {
    return breadcrumbState;
}

/**
 * パンくず状態をリセット
 * @returns {void}
 */
function resetBreadcrumbState() {
    breadcrumbState.currentBreadcrumbs = [];
    breadcrumbState.lastBreadcrumbData = null;
}
