# Vue.js検索アプリケーション

## 概要
検索機能をVue.js 3で実装した統合アプリケーションです。jQuery 依存を削除し、ネイティブDOM APIとVue.jsリアクティブシステムで構築されています。

## ファイル構成

### コアモジュール（8ファイル）
1. **utils.js** - ユーティリティ関数（escapeHtml、checkAllInTree、検索キーワード処理）
2. **searchCatalog.js** - カタログデータ管理とJS動的読み込み
3. **search.js** - 検索実行ロジック
4. **searchDisplay.js** - 検索結果表示とハイライト
5. **searchUI.js** - UI構築とイベント処理
6. **searchPagination.js** - ページネーション制御
7. **searchPageInit.js** - ページ初期化処理
8. **searchBreadcrumb.js** - パンくずリスト管理
9. **searchApp.js** - Vue.jsアプリケーション統合・起動

## アーキテクチャ

### 状態管理
各モジュールの状態管理オブジェクトをVue.reactive() でラップしてリアクティブ化:

```javascript
// searchCatalog.js
const catalogStore = {
    catalogue: [],
    catalogueJs: [],
    isLoaded: false,
    loadingProgress: 0
};

// search.js
const searchState = {
    keyword: '',
    normalizedKeywords: [],
    results: [],
    totalCount: 0,
    isSearching: false
};

// searchDisplay.js
const displayState = {
    results: [],
    displayedResults: [],
    hasResults: false,
    resultCount: 0
};

// searchUI.js
const uiState = {
    treeviewBuilt: false,
    firstPageBuilt: false,
    eventHandlersAttached: false
};

// searchPagination.js
const paginationState = {
    currentPage: 1,
    totalPages: 0,
    pageSize: 10,
    totalItems: 0,
    isVisible: false
};

// searchPageInit.js
const initState = {
    isInitialized: false,
    hasContents: false,
    initialKeyword: ''
};

// searchBreadcrumb.js
const breadcrumbState = {
    currentBreadcrumbs: [],
    lastBreadcrumbData: null
};
```

### カスタムイベント
モジュール間通信とVue.jsアプリへの通知:

- `catalogLoaded` - カタログ読み込み完了
- `searchCompleted` - 検索完了
- `treeviewBuilt` - ツリービュー構築完了
- `pageInitialized` - ページ初期化完了

### 主要関数
各モジュールの状態取得用関数:

- `getCatalogStore()` - searchCatalog.js
- `getSearchState()` - search.js
- `getDisplayState()` - searchDisplay.js
- `getUIState()` - searchUI.js
- `getPaginationState()` - searchPagination.js
- `getInitState()` - searchPageInit.js
- `getBreadcrumbState()` - searchBreadcrumb.js

## 使用方法

### 前提条件
Vue.js 3が必須です。search.htmlで読み込まれます。

```html
<!-- search.html -->
<script src="js/lib/vue.global.prod.js"></script>
```

### 起動フロー

1. **search.html読み込み**
   - Vue.js 3ライブラリ読み込み
   - 各検索モジュール読み込み
   - searchApp.js読み込み

2. **searchApp.js実行**
   - DOMContentLoaded時にVue.jsアプリケーション作成
   - 全モジュールの状態をreactive()でラップ
   - #appにマウント

3. **initializePage()実行**
   - onMountedフックから呼び出し
   - localStorageからデータ読み込み
   - UI初期化、イベント設定

## Vue.jsアプリケーション構造

### searchApp.jsの主要実装

```javascript
function initializeVueSearchApp() {
    const { createApp, reactive, computed, onMounted } = Vue;

    const app = createApp({
        setup() {
            // 全モジュールの状態をリアクティブ化
            const catalog = reactive(catalogStore);
            const searchData = reactive(searchState);
            const display = reactive(displayState);
            const ui = reactive(uiState);
            const pagination = reactive(paginationState);
            const init = reactive(initState);
            const breadcrumb = reactive(breadcrumbState);

            // 計算プロパティ
            const resultCount = computed(() => display.resultCount);
            const hasResults = computed(() => display.hasResults);

            // 検索実行メソッド
            const executeSearch = () => {
                const input = document.getElementById('searchkeyword');
                if (input && input.value !== "" && isCatalogLoaded()) {
                    search(input.value);
                }
            };

            // マウント時処理
            onMounted(() => {
                setupEventListeners();
                initializePage();
            });

            return {
                catalog, searchData, display, ui, pagination, init, breadcrumb,
                resultCount, hasResults,
                executeSearch
            };
        }
    });

    app.mount('#app');
}
```

## jQuery除去の主な変更

### DOM操作

**Before (jQuery):**
```javascript
$('#searchkeyword').val();
$('.btn-search').click(function() { /* ... */ });
$('.searchresults').find('.wSearchResultItem').each(function() { /* ... */ });
$(this).addClass('active');
```

**After (ネイティブDOM):**
```javascript
document.getElementById('searchkeyword').value;
document.querySelectorAll('.btn-search').forEach(button => {
    button.addEventListener('click', function() { /* ... */ });
});
document.querySelectorAll('.searchresults .wSearchResultItem').forEach(item => { /* ... */ });
element.classList.add('active');
```

### イベント委譲

**Before (jQuery):**
```javascript
$(document).on('click', '.check-toggle', function() { /* ... */ });
```

**After (ネイティブDOM):**
```javascript
document.body.addEventListener('click', function(e) {
    if (e.target.classList.contains('check-toggle')) {
        // イベント処理
    }
});
```

## 状態のリアクティブ化

各モジュールの状態オブジェクトは、searchApp.jsで`reactive()`によりリアクティブ化されます:

```javascript
// モジュール側（例: search.js）
const searchState = {
    keyword: '',
    results: [],
    totalCount: 0,
    isSearching: false
};

// Vue.jsアプリ側（searchApp.js）
const searchData = reactive(searchState);

// searchState.keywordを更新すると、
// searchDataも自動的に更新され、UIに反映される
```


## 技術的制約

### jQuery依存の残存
`searchPagination.js`はjQuery paginationプラグインに依存しています。将来的にはネイティブ実装またはVueコンポーネントへの置き換えを推奨します。

### Vue.js必須
このアプリケーションはVue.js 3が必須です。Vue.jsが読み込まれていない場合、エラーが発生します。

## 今後の拡張案

### 1. Vueコンポーネント化
現在はVue.jsアプリとして統合していますが、各機能をVue.jsコンポーネントに分離することで、再利用性とメンテナンス性が向上します:

```vue
<!-- SearchBox.vue -->
<template>
  <input v-model="keyword" @keyup.enter="executeSearch" />
  <button @click="executeSearch">検索</button>
</template>

<script setup>
import { ref, reactive } from 'vue';
const keyword = ref('');
const searchData = reactive(searchState);

function executeSearch() {
  search(keyword.value);
}
</script>
```

### 2. TypeScript化
型安全性を向上させるため、TypeScriptへの移行を検討:

```typescript
interface SearchState {
  keyword: string;
  results: SearchResult[];
  totalCount: number;
  isSearching: boolean;
}
```

### 3. paginationプラグインの置き換え
jQueryプラグイン依存を解消するため、Vue.jsコンポーネントでの実装:

```vue
<Pagination 
  :currentPage="currentPage"
  :totalPages="totalPages"
  @change="onPageChange"
/>
```

