/**
 * 検索カタログ管理モジュール (Vue.js対応版)
 * - カタログデータの管理
 * - 検索用JSファイルの動的読み込み
 * - カタログの初期化処理
 * 
 * Vue.jsのリアクティブシステムと互換性を持ちつつ、
 * 既存のグローバル変数アクセスパターンも維持します
 */

/**
 * カタログストアの状態管理
 * Vue.jsのリアクティブシステムで使用可能な構造
 */
const catalogStore = {
    // メニュー、検索等のためのサーチカタログ
    catalogue: [],
    
    // すべてのsearch.jsを一つの配列にまとめる
    catalogueJs: [],
    
    // 読み込み状態フラグ
    isLoaded: false,
    
    // 読み込み進捗（0-100%）
    loadingProgress: 0
};

// 後方互換性のためのグローバル変数エイリアス
const searchCatalogue = catalogStore.catalogue;
const searchCatalogueJs = catalogStore.catalogueJs;
let isLoaded = catalogStore.isLoaded;

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
    if (node) {
        node.countItem = 0;
        if (!node.breadCrum) {
            node.breadCrum = [];
        }
        node.breadCrum.push(node.title);

        if (node.searchjs) {
            searchCatalogueJs.push(node);
        }
        if (node.childs) {
            for (let i = 0; i < node.childs.length; i++) {
                if (!node.childs[i].breadCrum) {
                    node.childs[i].breadCrum = [];
                }

                node.childs[i].breadCrum = node.breadCrum.slice();

                //node.childs[i].breadCrum[node.childs[i].breadCrum.length] = node.title;
                collectSearchJs(node.childs[i]);
            }
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
 * jsをブラウザに読み込んでsearchWordsを呼び出す (Vue.js対応版)
 * jQuery依存を除去し、ネイティブDOM APIを使用
 * 
 * @param {Array} collection - 読み込むJSファイルのコレクション
 * @param {number} searchCatalogueItemChildPos - 現在処理中のインデックス
 */
async function loadSearchJs(collection, searchCatalogueItemChildPos) {
    const item = collection[searchCatalogueItemChildPos];
    
    try {
        await loadScript(item.searchjs);
        item.searchWords = searchWords;
        
        // 進捗を更新
        catalogStore.loadingProgress = Math.round(((searchCatalogueItemChildPos + 1) / collection.length) * 100);
        
        searchCatalogueItemChildPos++;
        if (searchCatalogueItemChildPos < collection.length) {
            await loadSearchJs(collection, searchCatalogueItemChildPos);
        } else {
            // 読み込み完了
            catalogStore.isLoaded = true;
            isLoaded = true; // 後方互換性
            
            // jQuery依存を除去: ネイティブDOM APIを使用
            document.body.classList.add('open');
            
            const boxNdSearch = document.querySelector('.box-nd-search');
            if (boxNdSearch) boxNdSearch.classList.remove('hidden');
            
            const boxContentS = document.querySelector('.box-content-s');
            if (boxContentS) boxContentS.classList.add('hidden');
            
            // 検索機能が利用可能になったことを通知
            if (typeof search === 'function') {
                search();
            }
            
            // カスタムイベントを発火（Vueコンポーネントで監視可能）
            window.dispatchEvent(new CustomEvent('catalogLoaded', {
                detail: {
                    catalogue: catalogStore.catalogue,
                    catalogueJs: catalogStore.catalogueJs
                }
            }));
        }
    } catch (error) {
        console.error(`Failed to load script: ${item.searchjs}`, error);
        // エラーイベントを発火
        window.dispatchEvent(new CustomEvent('catalogLoadError', {
            detail: { error, script: item.searchjs }
        }));
    }
}

/**
 * カタログが読み込み完了しているかチェック
 * @returns {boolean} 読み込み完了フラグ
 */
function isCatalogLoaded() {
    return catalogStore.isLoaded;
}

/**
 * 検索カタログを取得
 * @returns {Array} 検索カタログ配列
 */
function getSearchCatalogue() {
    return catalogStore.catalogue;
}

/**
 * 検索カタログJSを取得
 * @returns {Array} 検索カタログJS配列
 */
function getSearchCatalogueJs() {
    return catalogStore.catalogueJs;
}

/**
 * 検索カタログを設定
 * @param {Array} catalogue - 設定するカタログ配列
 */
function setSearchCatalogue(catalogue) {
    catalogStore.catalogue.length = 0;
    catalogStore.catalogue.push(...catalogue);
}

/**
 * カタログストア全体を取得（Vue.js用）
 * Vue.jsのリアクティブシステムで使用する場合はこちらを使用
 * @returns {Object} カタログストアオブジェクト
 */
function getCatalogStore() {
    return catalogStore;
}

/**
 * カタログをリセット
 * テストや再初期化時に使用
 */
function resetCatalog() {
    catalogStore.catalogue.length = 0;
    catalogStore.catalogueJs.length = 0;
    catalogStore.isLoaded = false;
    catalogStore.loadingProgress = 0;
    isLoaded = false; // 後方互換性
}

/**
 * 読み込み進捗を取得（0-100%）
 * @returns {number} 読み込み進捗
 */
function getLoadingProgress() {
    return catalogStore.loadingProgress;
}

/**
 * カタログストアを取得
 * @returns {Object} カタログストアオブジェクト
 */
function getCatalogStore() {
    return catalogStore;
}