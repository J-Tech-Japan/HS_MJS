/**
 * search.html ページ初期化モジュール (Vue.js版)
 * ページ初期化処理 - Vue.jsアプリから呼び出されます
 */

/**
 * ページ初期化状態管理オブジェクト
 */
const initState = {
    isInitialized: false,
    hasContents: false,
    initialKeyword: ''
};

/**
 * ページの初期化処理
 * Vue.jsアプリのonMountedで呼び出されます
 */
function initializePage() {
    // localStorageからコンテンツデータが取得できない場合はトップページにリダイレクト
    if (localStorage.getItem('contents') == null) {
        document.location.href = 'index.html';
        return;
    }
    
    // localStorageからコンテンツデータを読み込み
    setSearchCatalogue(JSON.parse(localStorage.getItem('contents')));
    initState.hasContents = true;

    // localStorageから検索キーワードを取得して検索ボックスに設定
    const searchKeyword = localStorage.getItem('searchkeyword');
    if (searchKeyword != null) {
        const searchInput = document.getElementById('searchkeyword');
        if (searchInput) {
            searchInput.value = searchKeyword;
            initState.initialKeyword = searchKeyword;
        }
    }

    // 検索ボタンクリック時のイベント処理 (ネイティブDOM API)
    const searchButtons = document.querySelectorAll('.btn-search');
    searchButtons.forEach(function(button) {
        button.addEventListener('click', function() {
            const searchInput = document.getElementById('searchkeyword');
            // 検索キーワードが入力されており、データ読み込みが完了している場合のみ検索実行
            if (searchInput && searchInput.value !== "") {
                if (isCatalogLoaded()) {
                    search();
                }
            }
        });
    });

    // ローディングアイコンを表示（検索実行中の表示）
    const boxClickSearch = document.querySelector('.box-click-search');
    if (boxClickSearch) {
        boxClickSearch.innerHTML = '<div class="loading"><i class="fas fa-spinner fa-spin"></i></div>';
    }

    // 初期ページの表示処理
    const boxContentS = document.querySelector('.box-content-s');
    if (boxContentS) {
        boxContentS.innerHTML = buildFirstPage();
        addHandleEventInFirstPage();
    }

    // ツリービュー構築（絞り込み項目の階層表示）
    buildTreeview();
    
    // 各検索項目のチェック状態を確認
    const searchInElements = document.querySelectorAll('.search-in');
    searchInElements.forEach(function(element) {
        checkAllInTree(element);
    });

    // 全ての絞り込み項目を閉じる処理 (ネイティブDOM API)
    const checkToggles = document.querySelectorAll('.box-check li .check-toggle');
    checkToggles.forEach(function(toggle) {
        // アクティブクラスを削除
        toggle.classList.remove('active');
        
        // 兄弟要素のul要素を非表示にする
        const nextUl = toggle.nextElementSibling;
        if (nextUl && nextUl.tagName === 'UL') {
            nextUl.style.display = 'none';
        }
        
        // 親のli要素の兄弟要素内のcheck-toggleとulを非表示にする
        const parentLi = toggle.parentElement;
        if (parentLi) {
            const siblings = Array.from(parentLi.parentElement.children).filter(child => child !== parentLi);
            siblings.forEach(function(sibling) {
                const siblingToggle = sibling.querySelector('.check-toggle');
                const siblingUl = sibling.querySelector('ul');
                if (siblingToggle) {
                    siblingToggle.classList.remove('active');
                }
                if (siblingUl) {
                    siblingUl.style.display = 'none';
                }
            });
        }
    });

    // 検索機能の初期化
    initSearch();

    // 検索ボックスでのEnterキー押下時の処理 (ネイティブDOM API)
    const searchInput = document.getElementById('searchkeyword');
    if (searchInput) {
        searchInput.addEventListener('keyup', function(e) {
            // Enterキーが押され、検索キーワードが入力されている場合
            if (e.keyCode === 13 || e.key === 'Enter') {
                if (searchInput.value !== "") {
                    // 検索ボタンのクリックイベントを実行
                    const searchButton = document.querySelector('.btn-search');
                    if (searchButton) {
                        searchButton.click();
                    }
                }
            }
        });
    }
    
    initState.isInitialized = true;
    
    // カスタムイベントを発火（Vue.jsコンポーネントで監視可能）
    window.dispatchEvent(new CustomEvent('pageInitialized', {
        detail: { initState }
    }));
}

/**
 * 初期化状態を取得
 * @returns {Object} 初期化状態オブジェクト
 */
function getInitState() {
    return initState;
}