/**
 * Vue.js検索アプリケーション統合モジュール
 * 全ての検索機能をVue 3リアクティブシステムで統合
 * このファイルは変換した全モジュールを統合し、Vue.jsアプリケーションとして起動します
 */

/**
 * Vue.jsアプリケーションの初期化と起動
 */
function initializeVueSearchApp() {
    if (typeof Vue === 'undefined') {
        console.error('Vue.js is not loaded. Please include Vue.js library.');
        return;
    }

    const { createApp, reactive, computed, watch, onMounted } = Vue;

    // Vue.jsアプリケーションを作成
    const app = createApp({
        setup() {
            // 各モジュールの状態をリアクティブにラップ
            const catalog = reactive(catalogStore);
            const searchData = reactive(searchState);
            const display = reactive(displayState);
            const ui = reactive(uiState);
            const pagination = reactive(paginationState);
            const init = reactive(initState);
            const breadcrumb = reactive(breadcrumbState);

            // 計算プロパティ
            const searchKeyword = computed(() => searchData.keyword);
            const resultCount = computed(() => display.resultCount);
            const hasResults = computed(() => display.hasResults);
            const isSearching = computed(() => searchData.isSearching);
            const isInitialized = computed(() => init.isInitialized);
            const displayedResults = computed(() => display.displayedResults || []);

            // 検索実行メソッド
            const executeSearch = () => {
                const input = document.getElementById('searchkeyword');
                if (input && input.value !== "") {
                    if (isCatalogLoaded()) {
                        search(input.value);
                    }
                }
            };

            // Enterキーでの検索
            const handleSearchKeyup = (event) => {
                if (event.keyCode === 13 || event.key === 'Enter') {
                    executeSearch();
                }
            };

            // カスタムイベントリスナーの設定
            const setupEventListeners = () => {
                // カタログ読み込み完了イベント
                window.addEventListener('catalogLoaded', () => {
                    console.log('Vue App: Catalog loaded', catalog);
                });

                // 検索完了イベント
                window.addEventListener('searchCompleted', (event) => {
                    console.log('Vue App: Search completed', event.detail);
                });

                // ツリービュー構築完了イベント
                window.addEventListener('treeviewBuilt', () => {
                    console.log('Vue App: Treeview built', ui);
                });

                // ページ初期化完了イベント
                window.addEventListener('pageInitialized', () => {
                    console.log('Vue App: Page initialized', init);
                });
            };

            // コンポーネントのマウント時処理
            onMounted(() => {
                console.log('Vue Search App mounted');
                setupEventListeners();
                
                // ページ初期化を実行
                initializePage();
            });

            // テンプレートで使用するデータとメソッドを返す
            return {
                // 状態
                catalog,
                searchData,
                display,
                ui,
                pagination,
                init,
                breadcrumb,

                // 計算プロパティ
                searchKeyword,
                resultCount,
                hasResults,
                isSearching,
                isInitialized,
                displayedResults,

                // メソッド
                executeSearch,
                handleSearchKeyup
            };
        }
    });

    // アプリケーションをマウント
    app.mount('#app');

    console.log('Vue Search App initialized successfully');
}

// アプリケーションの起動
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeVueSearchApp);
} else {
    initializeVueSearchApp();
}
