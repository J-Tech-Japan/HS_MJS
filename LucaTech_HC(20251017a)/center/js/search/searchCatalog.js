/**
 * 検索カタログ管理モジュール
 * - カタログデータの管理
 * - 検索用JSファイルの動的読み込み
 * - カタログの初期化処理
 */

// メニュー、検索等のためのサーチカタログ
const searchCatalogue = [];

// すべてのsearch.jsを一つの配列にまとめる
const searchCatalogueJs = [];

// 読み込み状態フラグ
let isLoaded = false;

/**
 * 検索機能の初期化
 * searchCatalogueのすべてのjsファイルを読み込む
 */
async function initSearch() {
    for (let i = 0; i < searchCatalogue.length; i++) {
        collectSearchJs(searchCatalogue[i]);
    }
    loadSearchJs(searchCatalogueJs, 0);
}

/**
 * すべてのsearch.jsを一つの配列にまとめる
 * ノードツリーを再帰的に処理してパンくずリストを構築
 * @param {Object} node - カタログノード
 */
function collectSearchJs(node) {
    node.countItem = 0;
    node.breadCrum = node.breadCrum || [];
    node.breadCrum.push(node.title);

    if (node.searchjs) {
        searchCatalogueJs.push(node);
    }
    if (node.childs) {
        for (const child of node.childs) {
            child.breadCrum = node.breadCrum.slice();
            collectSearchJs(child);
        }
    }
}

/**
 * スクリプトを動的に読み込む
 * @param {string} src - スクリプトのパス
 * @returns {Promise} 読み込み完了を示すPromise
 */
function loadScript(src) {
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
    });
}

/**
 * 検索UI表示状態を更新
 */
function updateSearchUI() {
    $('body').addClass('open');
    $('.box-nd-search').removeClass('hidden');
    $('.box-content-s').addClass('hidden');
}

/**
 * jsをブラウザに読み込んでsearchWordsを呼び出す
 * @param {Array} collection - 読み込むJSファイルのコレクション
 * @param {number} startIndex - 開始インデックス（デフォルト: 0）
 */
async function loadSearchJs(collection, startIndex = 0) {
    // すべてのスクリプトを並列で読み込む
    const loadPromises = collection.slice(startIndex).map(async (item) => {
        try {
            await loadScript(item.searchjs);
            item.searchWords = searchWords;
        } catch (error) {
            console.error(`Failed to load script: ${item.searchjs}`, error);
            // エラーが発生してもnullを返して処理を継続
            return null;
        }
    });
    
    // すべての読み込みが完了するまで待機
    await Promise.all(loadPromises);
    
    isLoaded = true;
    updateSearchUI();
    
    // 検索機能が利用可能になったことを通知
    if (typeof search === 'function') {
        search();
    }
}

/**
 * カタログが読み込み完了しているかチェック
 * @returns {boolean} 読み込み完了フラグ
 */
const isCatalogLoaded = () => isLoaded;

/**
 * 検索カタログを取得
 * @returns {Array} 検索カタログ配列
 */
const getSearchCatalogue = () => searchCatalogue;

/**
 * 検索カタログJSを取得
 * @returns {Array} 検索カタログJS配列
 */
const getSearchCatalogueJs = () => searchCatalogueJs;

/**
 * 検索カタログを設定
 * @param {Array} catalogue - 設定するカタログ配列
 */
function setSearchCatalogue(catalogue) {
    searchCatalogue.splice(0, searchCatalogue.length, ...catalogue);
}