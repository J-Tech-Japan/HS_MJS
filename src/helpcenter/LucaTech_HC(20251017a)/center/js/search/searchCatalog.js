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
 * jsをブラウザに読み込んでsearchWordsを呼び出す
 * @param {Array} collection - 読み込むJSファイルのコレクション
 * @param {number} searchCatalogueItemChildPos - 現在処理中のインデックス
 */
async function loadSearchJs(collection, searchCatalogueItemChildPos) {
    for (let i = searchCatalogueItemChildPos; i < collection.length; i++) {
        const item = collection[i];
        try {
            await loadScript(item.searchjs);
            item.searchWords = searchWords;
        } catch (error) {
            console.error(`Failed to load script: ${item.searchjs}`, error);
        }
    }
    
    isLoaded = true;
    $('body').addClass('open');
    $('.box-nd-search').removeClass('hidden');
    $('.box-content-s').addClass('hidden');
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
    searchCatalogue.length = 0;
    searchCatalogue.push(...catalogue);
}