# Vue.js移行計画書 - 横断検索機能

**作成日**: 2025年12月2日  
**対象**: `center\js\search` 配下の横断検索機能  
**移行方針**: 段階的移行（既存jQuery機能と共存させながら順次Vue.js化）

---

## 目次

1. [現状分析](#1-現状分析)
2. [移行戦略](#2-移行戦略)
3. [フェーズ別実装計画](#3-フェーズ別実装計画)
4. [技術仕様](#4-技術仕様)
5. [リスク管理](#5-リスク管理)
6. [テスト計画](#6-テスト計画)

---

## 1. 現状分析

### 1.1 既存アーキテクチャ

#### ファイル構成
```
center/js/search/
├── utils.js                    # ユーティリティ関数（HTMLエスケープ等）
├── searchCatalog.js            # カタログデータ管理・動的読み込み
├── searchBreadcrumb.js         # パンくずリスト生成
├── searchUI.js                 # UIコンポーネント構築（ツリービュー等）
├── searchDisplay.js            # 検索結果表示・ハイライト
├── searchPagination.js         # ページネーション機能
├── search.js                   # 検索ロジックのコア処理
└── searchPageInit.js           # ページ初期化処理
```

#### 機能依存関係
```
searchPageInit.js (初期化)
    ↓
    ├→ searchCatalog.js (データ管理)
    │   └→ 動的にsearch.jsファイルを読み込み
    ├→ searchUI.js (UI構築)
    │   ├→ buildTreeview() - 左サイドバーツリー
    │   └→ buildFirstPage() - 初期カテゴリ選択画面
    ├→ search.js (検索実行)
    │   └→ searchDisplay.js を呼び出し
    ├→ searchDisplay.js (結果表示)
    │   └→ searchPagination.js を呼び出し
    └→ searchBreadcrumb.js (パンくず)
```

### 1.2 主要機能の実装方式

| 機能 | 現在の実装 | データフロー |
|------|-----------|------------|
| **検索キーワード入力** | jQuery `$("#searchkeyword")` | DOM直接操作 |
| **カタログツリー表示** | jQueryでHTML文字列生成 | `buildTreeviewNode()` 再帰関数 |
| **チェックボックス管理** | jQueryイベントハンドラ | DOM state直接変更 |
| **検索実行** | `search()` 関数 | グローバル関数呼び出し |
| **検索結果表示** | jQueryでDOM生成・挿入 | `$(".searchresults").append()` |
| **ページネーション** | pagination.min.js プラグイン | 外部ライブラリ依存 |
| **ハイライト** | DOM走査とstyle直接変更 | `highlightSearchWord()` |

### 1.3 技術的負債

1. **グローバル変数の多用**
   - `searchCatalogue`, `searchCatalogueJs`, `isLoaded` など
   
2. **DOM直接操作**
   - 状態管理とUI更新が分離されていない
   
3. **イベントハンドラの複雑性**
   - チェックボックス変更時の連動処理が追跡困難
   
4. **再利用性の低さ**
   - 各関数が特定のDOM構造に依存

---

## 2. 移行戦略

### 2.1 基本方針

**段階的移行 (Incremental Migration)**
- 一度にすべてを書き換えない
- フェーズごとに機能を切り出してVue化
- 各フェーズで動作確認とテストを実施
- jQuery機能とVue機能を共存させる

### 2.2 移行原則

1. **後方互換性の維持**
   - 既存のグローバル関数は残す（deprecatedとしてマーク）
   
2. **漸進的強化**
   - まず表示層（UI）からVue化
   - 次にデータ層（状態管理）を整理
   - 最後にロジック層を抽出
   
3. **テスト駆動**
   - 各フェーズ完了後に必ず動作確認
   
4. **ドキュメント更新**
   - 移行済み機能と未移行機能を明確に

---

## 3. フェーズ別実装計画

### フェーズ0: 準備（1-2日）

#### 目標
Vue.js環境のセットアップと既存コードの理解

#### タスク
- [x] 既存コードの完全な依存関係マップ作成
- [ ] Vue.js 3.x のCDN有効化（既にコメントアウトされている）
- [ ] Vueコンポーネント格納ディレクトリ作成 `js/components/search/`
- [ ] 基本的なVueアプリケーションインスタンス作成
- [ ] 開発者ツールのセットアップ

#### 成果物
```javascript
// js/components/search/searchApp.js
const { createApp } = Vue;

const searchApp = createApp({
  data() {
    return {
      // 後続フェーズで追加
    };
  }
});

searchApp.mount('#app');
```

---

### フェーズ1: 検索キーワード入力のVue化（2-3日）

#### 目標
検索ボックスとボタンをVueのリアクティブシステムで管理

#### 対象ファイル
- `search.html` - HTML構造の修正
- `searchPageInit.js` - イベントハンドラの段階的移行

#### 実装内容

**Before (jQuery):**
```javascript
$("#searchkeyword").val()
$(".btn-search").click(function() { ... })
```

**After (Vue):**
```javascript
// データバインディング
data() {
  return {
    searchKeyword: ''
  }
},
methods: {
  onSearch() {
    if (this.searchKeyword.trim() !== '' && isCatalogLoaded()) {
      search(); // 既存関数を呼び出し（共存）
    }
  }
}
```

**HTML修正:**
```html
<!-- Before -->
<input id="searchkeyword" class="form-control-search" type="search">
<button class="btn-search"></button>

<!-- After -->
<input v-model="searchKeyword" 
       class="form-control-search" 
       type="search"
       @keyup.enter="onSearch">
<button class="btn-search" @click="onSearch"></button>
```

#### ブリッジ対応
既存のjQuery依存コードのために、検索実行時に値を同期:
```javascript
methods: {
  onSearch() {
    // Vue → jQuery同期（既存コード対応）
    $("#searchkeyword").val(this.searchKeyword);
    search(); // 既存関数
  }
}
```

#### 検証項目
- [ ] 検索ボックスへの入力が正常
- [ ] Enterキーで検索実行
- [ ] ボタンクリックで検索実行
- [ ] 既存のsearch()関数が正常に動作
- [ ] localStorageからのキーワード復元が動作

---

### フェーズ2: カタログデータの状態管理（3-4日）

#### 目標
`searchCatalogue`と`searchCatalogueJs`をVueのリアクティブデータに移行

#### 対象ファイル
- `searchCatalog.js` - データ管理ロジックの抽出
- `searchApp.js` - Vueアプリへの統合

#### 実装内容

**データ構造の定義:**
```javascript
data() {
  return {
    searchKeyword: '',
    catalogue: {
      items: [],        // searchCatalogue相当
      jsFiles: [],      // searchCatalogueJs相当
      isLoaded: false
    },
    selectedCategories: new Set()
  }
}
```

**既存関数のラップ:**
```javascript
methods: {
  async initializeCatalogue() {
    const data = JSON.parse(localStorage.getItem('contents'));
    if (!data) {
      window.location.href = 'index.html';
      return;
    }
    
    // 既存のsetSearchCatalogue()を呼び出し
    setSearchCatalogue(data);
    
    // Vueの状態と同期
    this.catalogue.items = getSearchCatalogue();
    
    await initSearch(); // 既存の初期化
    this.catalogue.isLoaded = true;
  }
}
```

#### ブリッジ対応
既存のグローバル関数はそのまま維持し、Vue側から呼び出す:
```javascript
// searchCatalog.js に追加
function syncCatalogueToVue(vueApp) {
  // Vueアプリにデータを同期する関数（オプション）
}
```

#### 検証項目
- [ ] カタログデータが正常にロード
- [ ] 既存のツリービュー構築が動作
- [ ] 検索機能が正常に動作
- [ ] ページ遷移（戻る/進む）が正常

---

### フェーズ3: ツリービューコンポーネント化（4-5日）

#### 目標
左サイドバーのツリービューをVueコンポーネントに変換

#### 対象ファイル
- `searchUI.js` - `buildTreeviewNode()` のVue化
- 新規: `js/components/search/TreeView.js`
- 新規: `js/components/search/TreeNode.js`

#### コンポーネント設計

**TreeNode.vue 相当（Single File ComponentなしでCDN版）:**
```javascript
// js/components/search/TreeNode.js
const TreeNode = {
  name: 'TreeNode',
  props: {
    node: {
      type: Object,
      required: true
    },
    level: {
      type: Number,
      default: 0
    }
  },
  data() {
    return {
      isExpanded: false
    };
  },
  computed: {
    hasChildren() {
      return Array.isArray(this.node.childs) && this.node.childs.length > 0;
    },
    checkboxId() {
      return `search-in-${this.node.id}`;
    },
    isChecked: {
      get() {
        return this.$parent.selectedCategories.has(this.node.id);
      },
      set(value) {
        this.$emit('toggle-check', this.node.id, value);
      }
    }
  },
  methods: {
    toggleExpand() {
      this.isExpanded = !this.isExpanded;
    },
    onCheckChange(checked) {
      this.$emit('check-change', {
        nodeId: this.node.id,
        checked: checked,
        propagate: true // 子要素にも伝播
      });
    }
  },
  template: `
    <li>
      <div class="custom-checkbox" :class="{'has-childs': hasChildren}">
        <input 
          type="checkbox" 
          v-model="isChecked"
          :id="checkboxId"
          class="custom-control-input search-in"
          @change="onCheckChange($event.target.checked)">
        <label 
          :for="checkboxId" 
          class="custom-control-label">
          {{ node.title }}
          <span class="count">({{ node.countItem || 0 }})</span>
        </label>
      </div>
      
      <div 
        v-if="hasChildren"
        class="check-toggle"
        :class="{active: isExpanded}"
        @click="toggleExpand">
      </div>
      
      <transition name="slide">
        <ul v-if="hasChildren && isExpanded" class="box-item box-toggle">
          <tree-node
            v-for="child in node.childs"
            :key="child.id"
            :node="child"
            :level="level + 1"
            @check-change="$emit('check-change', $event)">
          </tree-node>
        </ul>
      </transition>
    </li>
  `
};
```

**TreeView コンポーネント:**
```javascript
// js/components/search/TreeView.js
const TreeView = {
  name: 'TreeView',
  components: {
    TreeNode
  },
  props: {
    catalogueItems: {
      type: Array,
      required: true
    },
    selectedIds: {
      type: Set,
      required: true
    }
  },
  methods: {
    handleCheckChange(event) {
      this.$emit('selection-change', event);
    }
  },
  template: `
    <div class="column-left-container">
      <div 
        v-for="catalogue in catalogueItems" 
        :key="catalogue.id"
        class="box-check">
        <ul>
          <tree-node 
            :node="catalogue"
            @check-change="handleCheckChange">
          </tree-node>
        </ul>
      </div>
    </div>
  `
};
```

#### 親アプリでの統合:
```javascript
// searchApp.js
const searchApp = createApp({
  components: {
    TreeView
  },
  data() {
    return {
      searchKeyword: '',
      catalogue: { items: [], jsFiles: [], isLoaded: false },
      selectedCategories: new Set()
    };
  },
  methods: {
    onSelectionChange(event) {
      const { nodeId, checked, propagate } = event;
      
      if (checked) {
        this.selectedCategories.add(nodeId);
        if (propagate) {
          this.addChildrenRecursively(nodeId);
        }
      } else {
        this.selectedCategories.delete(nodeId);
        if (propagate) {
          this.removeChildrenRecursively(nodeId);
        }
      }
      
      // 検索結果を再表示
      if (this.searchKeyword) {
        this.performSearch();
      }
    },
    addChildrenRecursively(nodeId) {
      // 子ノードを再帰的にselectedCategoriesに追加
    },
    removeChildrenRecursively(nodeId) {
      // 子ノードを再帰的にselectedCategoriesから削除
    }
  }
});
```

#### HTML修正:
```html
<!-- Before -->
<div class="column-left-container"></div>

<!-- After -->
<tree-view
  :catalogue-items="catalogue.items"
  :selected-ids="selectedCategories"
  @selection-change="onSelectionChange">
</tree-view>
```

#### 段階的移行戦略

1. **ステップ3.1**: TreeNodeコンポーネント作成・単体テスト
2. **ステップ3.2**: TreeViewコンポーネント作成・統合テスト
3. **ステップ3.3**: 既存のjQuery版と新Vueコンポーネントを切り替えフラグで共存
   ```javascript
   const USE_VUE_TREEVIEW = true; // フラグで切り替え
   
   if (USE_VUE_TREEVIEW) {
     // Vueコンポーネント使用
   } else {
     buildTreeview(); // 既存jQuery版
   }
   ```
4. **ステップ3.4**: 十分なテスト後、jQuery版を削除

#### 検証項目
- [ ] ツリー表示が既存と同等
- [ ] チェックボックス連動が正常
- [ ] 折りたたみ/展開動作が正常
- [ ] 検索結果のカウント表示が正常
- [ ] 親チェック時に子も全選択
- [ ] 「すべて選択」機能が動作
- [ ] パフォーマンス（大量ノード時）

---

### フェーズ4: 検索結果表示のVue化（3-4日）

#### 目標
検索結果リストをVueで管理、ハイライト機能の整理

#### 対象ファイル
- `searchDisplay.js` - 結果表示ロジックの抽出
- 新規: `js/components/search/SearchResults.js`

#### コンポーネント設計

```javascript
// js/components/search/SearchResultItem.js
const SearchResultItem = {
  name: 'SearchResultItem',
  props: {
    item: {
      type: Object,
      required: true
    },
    searchWords: {
      type: Array,
      default: () => []
    }
  },
  computed: {
    highlightedContent() {
      let text = this.item.displayText;
      this.searchWords.forEach(word => {
        const regex = new RegExp(`(${this.escapeRegExp(word)})`, 'gi');
        text = text.replace(regex, '<mark>$1</mark>');
      });
      return text;
    }
  },
  methods: {
    escapeRegExp(str) {
      return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    },
    openLink(event) {
      event.preventDefault();
      openhelplink(this.item.url, event); // 既存関数
    }
  },
  template: `
    <div class="wSearchResultItem">
      <div class="wSearchResultTitle title-s">
        <a class="nolink" href="#" @click="openLink">
          {{ item.title }}
        </a>
      </div>
      <div class="wSearchResultBreadCrum">
        <span v-for="(crumb, index) in item.breadcrumbs" :key="index">
          {{ crumb }}
          <span v-if="index < item.breadcrumbs.length - 1"> &gt; </span>
        </span>
      </div>
      <div class="wSearchContent">
        <span class="wSearchContext nd-p" v-html="highlightedContent"></span>
      </div>
    </div>
  `
};
```

```javascript
// js/components/search/SearchResults.js
const SearchResults = {
  name: 'SearchResults',
  components: {
    SearchResultItem
  },
  props: {
    results: {
      type: Array,
      default: () => []
    },
    searchWords: {
      type: Array,
      default: () => []
    },
    currentPage: {
      type: Number,
      default: 1
    },
    pageSize: {
      type: Number,
      default: 10
    }
  },
  computed: {
    totalResults() {
      return this.results.length;
    },
    paginatedResults() {
      const start = (this.currentPage - 1) * this.pageSize;
      const end = start + this.pageSize;
      return this.results.slice(start, end);
    },
    hasResults() {
      return this.results.length > 0;
    }
  },
  template: `
    <div>
      <div class="column-right-header">
        <div v-if="hasResults" class="top-content hasresult">
          「<span>{{ searchWords.join(' ') }}</span>」の検索結果：
          <span>{{ totalResults }}</span>件見つかりました。
        </div>
        <div v-else class="top-content noresult">
          <p class="icon">
            <img src="images/ic_g_search_result_non_bl_96px.svg">
          </p>
          <p>申し訳ありません。<br>検索条件に一致するトピックはありませんでした。</p>
        </div>
      </div>
      
      <div class="searchresults items">
        <search-result-item
          v-for="(item, index) in paginatedResults"
          :key="index"
          :item="item"
          :search-words="searchWords">
        </search-result-item>
      </div>
    </div>
  `
};
```

#### 親アプリでの統合:
```javascript
data() {
  return {
    searchKeyword: '',
    catalogue: { items: [], jsFiles: [], isLoaded: false },
    selectedCategories: new Set(),
    searchResults: [],
    currentPage: 1
  };
},
methods: {
  performSearch() {
    const searchWords = normalizeSearchKeyword(this.searchKeyword);
    const searchQuery = searchWords.map(w => `:contains(${w})`).join('');
    
    this.searchResults = [];
    
    this.catalogue.jsFiles.forEach(catalogueItem => {
      if (this.selectedCategories.has(catalogueItem.id)) {
        const findItems = catalogueItem.searchWords.find(
          ".search_word" + searchQuery
        );
        
        findItems.each((index, element) => {
          this.searchResults.push({
            id: `${catalogueItem.id}-${index}`,
            title: $(element).parent().find('.search_title').text(),
            displayText: searchKeywordsInString(
              $(element).parent().text(), 
              searchWords
            ),
            breadcrumbs: catalogueItem.breadCrum,
            url: this.buildResultUrl(catalogueItem, element)
          });
        });
      }
    });
    
    this.currentPage = 1; // 検索時は1ページ目に
  }
}
```

#### 検証項目
- [ ] 検索結果が正しく表示
- [ ] ハイライトが正常に機能
- [ ] パンくずリストが表示
- [ ] リンククリックで正しいページが開く
- [ ] 結果数カウントが正確
- [ ] 結果なし時のメッセージ表示

---

### フェーズ5: ページネーション統合（2-3日）

#### 目標
Vue.jsでページネーションを管理、既存プラグインから脱却

#### 実装方法

**Option A: Vueネイティブ実装（推奨）**
```javascript
// js/components/search/Pagination.js
const Pagination = {
  name: 'Pagination',
  props: {
    currentPage: Number,
    totalItems: Number,
    pageSize: Number
  },
  computed: {
    totalPages() {
      return Math.ceil(this.totalItems / this.pageSize);
    },
    pages() {
      const pages = [];
      const maxVisible = 5;
      let start = Math.max(1, this.currentPage - 2);
      let end = Math.min(this.totalPages, start + maxVisible - 1);
      
      if (end - start < maxVisible - 1) {
        start = Math.max(1, end - maxVisible + 1);
      }
      
      for (let i = start; i <= end; i++) {
        pages.push(i);
      }
      return pages;
    }
  },
  methods: {
    goToPage(page) {
      if (page >= 1 && page <= this.totalPages) {
        this.$emit('page-change', page);
      }
    }
  },
  template: `
    <div class="pagination-container">
      <ul class="pagination">
        <li :class="{disabled: currentPage === 1}">
          <a href="#" @click.prevent="goToPage(currentPage - 1)">
            前へ
          </a>
        </li>
        
        <li v-if="pages[0] > 1">
          <a href="#" @click.prevent="goToPage(1)">1</a>
        </li>
        <li v-if="pages[0] > 2"><span>...</span></li>
        
        <li 
          v-for="page in pages" 
          :key="page"
          :class="{active: page === currentPage}">
          <a href="#" @click.prevent="goToPage(page)">{{ page }}</a>
        </li>
        
        <li v-if="pages[pages.length - 1] < totalPages - 1">
          <span>...</span>
        </li>
        <li v-if="pages[pages.length - 1] < totalPages">
          <a href="#" @click.prevent="goToPage(totalPages)">
            {{ totalPages }}
          </a>
        </li>
        
        <li :class="{disabled: currentPage === totalPages}">
          <a href="#" @click.prevent="goToPage(currentPage + 1)">
            次へ
          </a>
        </li>
      </ul>
    </div>
  `
};
```

**Option B: 既存プラグインをVueでラップ**
```javascript
// 過渡期の対応として
mounted() {
  this.initPaginationPlugin();
},
watch: {
  results() {
    this.updatePaginationPlugin();
  }
}
```

#### 検証項目
- [ ] ページ切り替えが正常
- [ ] 上下のページネーションが連動
- [ ] ページ数表示が正確
- [ ] 前へ/次へボタンが適切に無効化
- [ ] URL hash との連動（オプション）

---

### フェーズ6: 初期カテゴリ選択画面のVue化（2-3日）

#### 目標
`buildFirstPage()` で生成していた初期画面をVueコンポーネント化

#### コンポーネント設計

```javascript
// js/components/search/CategorySelection.js
const CategorySelection = {
  name: 'CategorySelection',
  props: {
    catalogueItems: Array
  },
  data() {
    return {
      selections: new Map() // カテゴリID -> {parent: boolean, children: Set}
    };
  },
  methods: {
    initializeSelections() {
      this.catalogueItems.forEach(catalogue => {
        this.selections.set(catalogue.id, {
          parent: true,
          children: new Set(catalogue.childs.map(c => c.id))
        });
      });
    },
    toggleParent(catalogueId) {
      const selection = this.selections.get(catalogueId);
      selection.parent = !selection.parent;
      
      const catalogue = this.catalogueItems.find(c => c.id === catalogueId);
      if (selection.parent) {
        catalogue.childs.forEach(child => {
          selection.children.add(child.id);
        });
      } else {
        selection.children.clear();
      }
    },
    toggleChild(catalogueId, childId) {
      const selection = this.selections.get(catalogueId);
      
      if (selection.children.has(childId)) {
        selection.children.delete(childId);
      } else {
        selection.children.add(childId);
      }
      
      const catalogue = this.catalogueItems.find(c => c.id === catalogueId);
      selection.parent = selection.children.size === catalogue.childs.length;
    },
    proceedToSearch() {
      // 選択状態を親アプリに伝達
      const selectedIds = new Set();
      this.selections.forEach((selection, catalogueId) => {
        selection.children.forEach(childId => {
          selectedIds.add(childId);
        });
      });
      this.$emit('selections-confirmed', selectedIds);
    }
  },
  mounted() {
    this.initializeSelections();
  },
  template: `
    <div class="box-content-s">
      <div 
        v-for="catalogue in catalogueItems" 
        :key="catalogue.id"
        class="box-s-1">
        <div class="h1 custom-checkbox">
          <input 
            type="checkbox"
            :checked="selections.get(catalogue.id)?.parent"
            @change="toggleParent(catalogue.id)"
            class="custom-control-input parent"
            :id="'search-in-top-' + catalogue.id">
          <label 
            :for="'search-in-top-' + catalogue.id"
            class="custom-control-label">
            {{ catalogue.title }}
          </label>
        </div>
        
        <div class="box-item">
          <ul class="box-ul box-ul-1">
            <li v-for="child in catalogue.childs" :key="child.id">
              <div class="custom-checkbox">
                <input 
                  type="checkbox"
                  :checked="selections.get(catalogue.id)?.children.has(child.id)"
                  @change="toggleChild(catalogue.id, child.id)"
                  class="custom-control-input child"
                  :id="'search-in-top-' + child.id">
                <label 
                  :for="'search-in-top-' + child.id"
                  class="custom-control-label">
                  {{ child.title }}
                </label>
              </div>
            </li>
          </ul>
        </div>
      </div>
      
      <button @click="proceedToSearch" class="btn-proceed">
        検索画面へ進む
      </button>
    </div>
  `
};
```

#### 検証項目
- [ ] 初期表示で全カテゴリチェック済み
- [ ] 親チェックで子も連動
- [ ] 子の部分選択で親が半選択状態（視覚）
- [ ] 選択状態が検索画面に引き継がれる

---

### フェーズ7: グローバル状態管理の最適化（3-4日）

#### 目標
Pinia または VueX を導入して状態管理を集約（オプション）

#### Option A: Pinia導入（推奨 - Vue 3標準）

```javascript
// js/stores/searchStore.js
const { defineStore } = Pinia;

const useSearchStore = defineStore('search', {
  state: () => ({
    keyword: '',
    catalogue: {
      items: [],
      jsFiles: [],
      isLoaded: false
    },
    selectedCategories: new Set(),
    searchResults: [],
    currentPage: 1,
    pageSize: 10
  }),
  
  getters: {
    totalResults: (state) => state.searchResults.length,
    paginatedResults: (state) => {
      const start = (state.currentPage - 1) * state.pageSize;
      const end = start + state.pageSize;
      return state.searchResults.slice(start, end);
    },
    hasResults: (state) => state.searchResults.length > 0
  },
  
  actions: {
    async initializeCatalogue() {
      const data = JSON.parse(localStorage.getItem('contents'));
      if (!data) {
        window.location.href = 'index.html';
        return;
      }
      
      setSearchCatalogue(data);
      this.catalogue.items = getSearchCatalogue();
      await initSearch();
      this.catalogue.isLoaded = true;
    },
    
    performSearch() {
      const searchWords = normalizeSearchKeyword(this.keyword);
      const searchQuery = searchWords.map(w => `:contains(${w})`).join('');
      
      this.searchResults = [];
      
      this.catalogue.jsFiles.forEach(catalogueItem => {
        if (this.selectedCategories.has(catalogueItem.id)) {
          const findItems = catalogueItem.searchWords.find(
            ".search_word" + searchQuery
          );
          
          findItems.each((index, element) => {
            this.searchResults.push({
              // ... 結果データ構築
            });
          });
        }
      });
      
      this.currentPage = 1;
    },
    
    toggleCategory(nodeId, checked) {
      if (checked) {
        this.selectedCategories.add(nodeId);
      } else {
        this.selectedCategories.delete(nodeId);
      }
    },
    
    setPage(page) {
      this.currentPage = page;
    }
  }
});
```

#### Option B: シンプルなComposable関数（小規模向け）

```javascript
// js/composables/useSearch.js
function useSearch() {
  const state = Vue.reactive({
    keyword: '',
    results: [],
    isLoading: false
  });
  
  const performSearch = async () => {
    state.isLoading = true;
    // 検索ロジック
    state.isLoading = false;
  };
  
  return {
    state: Vue.readonly(state),
    performSearch
  };
}
```

#### 検証項目
- [ ] 状態の一元管理が機能
- [ ] コンポーネント間のデータ共有が正常
- [ ] パフォーマンスが改善（不要な再レンダリング削減）

---

### フェーズ8: jQuery完全削除と最適化（2-3日）

#### 目標
jQuery依存を完全に排除、ネイティブJSまたはVue機能に置き換え

#### 置き換え対象

| jQuery | Vue/ネイティブJS |
|--------|-----------------|
| `$().find()` | `document.querySelector()` または `ref` |
| `$().on('click')` | `@click` |
| `$().val()` | `v-model` |
| `$().text()` | `{{ }}` または `textContent` |
| `$().append()` | `v-for` + `push()` |
| `$().slideToggle()` | `<Transition>` |

#### クリーンアップ作業
- [ ] jQuery依存コードを全検索して置き換え
- [ ] 未使用の関数を削除
- [ ] グローバル変数を整理
- [ ] ESModules化を検討

#### 検証項目
- [ ] 全機能が動作
- [ ] バンドルサイズが削減
- [ ] パフォーマンスが向上

---

## 4. 技術仕様

### 4.1 開発環境

| 項目 | 選択 | 理由 |
|------|------|------|
| **Vue.js バージョン** | Vue 3.x (CDN) | 既存プロジェクト構成に適合 |
| **ビルドツール** | なし（CDN直接利用） | シンプルな統合、学習コスト低 |
| **状態管理** | Pinia (オプション) | Vue 3公式推奨 |
| **コンポーネントスタイル** | Options API | CDN版での実装容易性 |

### 4.2 ディレクトリ構造（移行後）

```
center/js/
├── lib/
│   ├── vue.global.prod.js         # Vue 3 CDN
│   └── pinia.iife.js              # Pinia (オプション)
├── components/
│   └── search/
│       ├── searchApp.js           # メインアプリ
│       ├── TreeView.js            # ツリービューコンポーネント
│       ├── TreeNode.js            # ツリーノードコンポーネント
│       ├── SearchResults.js       # 検索結果コンポーネント
│       ├── SearchResultItem.js    # 検索結果アイテム
│       ├── Pagination.js          # ページネーション
│       └── CategorySelection.js   # カテゴリ選択
├── stores/
│   └── searchStore.js             # Pinia Store (オプション)
├── composables/
│   └── useSearch.js               # 検索ロジック（オプション）
└── search/
    ├── utils.js                   # ユーティリティ（そのまま）
    ├── searchCatalog.js           # データ管理（部分的に残る可能性）
    ├── searchBreadcrumb.js        # パンくず（そのまま or Vue化）
    └── [非推奨]
        ├── searchUI.js            # → TreeView.js へ移行
        ├── searchDisplay.js       # → SearchResults.js へ移行
        ├── searchPagination.js    # → Pagination.js へ移行
        ├── search.js              # → searchStore.js へ移行
        └── searchPageInit.js      # → searchApp.js へ移行
```

### 4.3 命名規則

| 対象 | 規則 | 例 |
|------|------|-----|
| コンポーネント | PascalCase | `TreeView`, `SearchResults` |
| メソッド | camelCase | `performSearch`, `toggleCategory` |
| データプロパティ | camelCase | `searchKeyword`, `selectedCategories` |
| イベント名 | kebab-case | `@selection-change`, `@page-change` |
| CSS クラス | kebab-case | `.search-result-item` |

---

## 5. リスク管理

### 5.1 リスク一覧

| リスク | 影響度 | 対策 |
|--------|--------|------|
| **既存機能の破壊** | 高 | フェーズごとに十分なテスト、フラグ切り替えで共存 |
| **パフォーマンス劣化** | 中 | ベンチマーク取得、仮想スクロール導入検討 |
| **学習コスト** | 中 | Vue.js基礎研修、ペアプログラミング |
| **スコープクリープ** | 中 | 各フェーズの範囲を厳密に定義 |
| **依存ライブラリの競合** | 低 | Vue 3の独立性、jQuery noConflict() |

### 5.2 ロールバック計画

各フェーズでGit tagを作成:
```
git tag -a vue-migration-phase-1 -m "フェーズ1完了: 検索キーワード入力"
git tag -a vue-migration-phase-2 -m "フェーズ2完了: カタログ状態管理"
...
```

問題発生時は該当タグに戻す:
```
git checkout vue-migration-phase-N
```

---

## 6. テスト計画

### 6.1 テストレベル

#### ユニットテスト（推奨）
```javascript
// Vitest or Jest
describe('TreeNode Component', () => {
  it('should toggle expand state', () => {
    const wrapper = mount(TreeNode, {
      props: { node: mockNode }
    });
    
    wrapper.find('.check-toggle').trigger('click');
    expect(wrapper.vm.isExpanded).toBe(true);
  });
});
```

#### 統合テスト
- 検索フロー全体の動作確認
- カテゴリ選択 → 検索実行 → 結果表示 → ページング

#### E2Eテスト（オプション）
- Playwright or Cypress
- 主要ユーザーシナリオの自動化

### 6.2 手動テストチェックリスト

各フェーズで以下を確認:

**機能テスト**
- [ ] 検索キーワード入力
- [ ] 検索実行（ボタン/Enter）
- [ ] カテゴリツリー表示
- [ ] チェックボックス操作
- [ ] 検索結果表示
- [ ] ハイライト機能
- [ ] ページネーション
- [ ] リンク遷移

**非機能テスト**
- [ ] レスポンシブ対応
- [ ] ブラウザ互換性（Chrome, Firefox, Edge, Safari）
- [ ] パフォーマンス（1000件以上の結果）
- [ ] アクセシビリティ（キーボード操作）

**回帰テスト**
- [ ] 既存機能が破壊されていないか
- [ ] localStorageの読み書き
- [ ] URL履歴管理

---

## 7. 実装スケジュール

### 7.1 推奨スケジュール（3-4週間）

| 週 | フェーズ | 作業内容 | 成果物 |
|----|---------|---------|--------|
| **Week 1** | Phase 0-1 | 環境準備、検索ボックスVue化 | 動作する検索ボックス |
| **Week 2** | Phase 2-3 | カタログ管理、ツリーVue化 | 動作するツリービュー |
| **Week 3** | Phase 4-5 | 検索結果表示、ページネーション | 完全動作する検索機能 |
| **Week 4** | Phase 6-8 | 初期画面、最適化、jQuery削除 | 本番リリース可能 |

### 7.2 マイルストーン

- **M1 (Week 1終了)**: 検索ボックスがVueで動作
- **M2 (Week 2終了)**: ツリービューがVueで動作
- **M3 (Week 3終了)**: 検索機能全体がVueで動作
- **M4 (Week 4終了)**: jQuery完全削除、本番リリース

---

## 8. 成功の指標

### 8.1 KPI

- [ ] jQuery依存度: 0% （完全削除）
- [ ] コード行数: 20-30%削減
- [ ] バンドルサイズ: 現状維持または削減
- [ ] 検索速度: 現状維持または向上
- [ ] メンテナンス性スコア: +50%（主観評価）

### 8.2 品質基準

- [ ] すべてのテストケースがパス
- [ ] ブラウザ互換性確認完了
- [ ] コードレビュー完了
- [ ] ドキュメント更新完了

---

## 9. 次のステップ

移行計画が承認されたら、以下の順で進めます:

1. **Phase 0の実行**: Vue.js環境セットアップ
2. **Phase 1の実装開始**: 検索ボックスVue化
3. **定期的なレビュー**: 各フェーズ終了時に進捗確認
4. **問題の早期発見**: デイリーで動作確認

---

## 付録

### A. 参考資料

- [Vue.js 3 公式ドキュメント](https://v3.vuejs.org/)
- [Pinia ドキュメント](https://pinia.vuejs.org/)
- [Vue.js Migration Guide (Vue 2 → 3)](https://v3-migration.vuejs.org/)
- [jQuery to Vue.js Migration Patterns](https://www.vuemastery.com/courses/real-world-vue3/orientation)

### B. 用語集

| 用語 | 説明 |
|------|------|
| **リアクティブ** | データ変更が自動的にUIに反映される仕組み |
| **コンポーネント** | 再利用可能なUI部品 |
| **プロパティ (Props)** | 親から子への一方向データフロー |
| **イベントエミット** | 子から親への通知 |
| **Computed** | 依存データに基づく算出プロパティ |
| **Watcher** | データ変更を監視するハンドラ |

### C. チェックリスト（印刷用）

```
□ Phase 0: 環境準備
  □ Vue.js CDN有効化
  □ 基本アプリ作成
  □ 動作確認

□ Phase 1: 検索ボックス
  □ v-model設定
  □ @click/@keyup.enter設定
  □ 既存関数との連携確認

□ Phase 2: カタログ管理
  □ データ構造定義
  □ 初期化処理
  □ 既存関数との同期

□ Phase 3: ツリービュー
  □ TreeNodeコンポーネント
  □ TreeViewコンポーネント
  □ チェックボックスロジック

□ Phase 4: 検索結果
  □ SearchResultItemコンポーネント
  □ SearchResultsコンポーネント
  □ ハイライト機能

□ Phase 5: ページネーション
  □ Paginationコンポーネント
  □ ページ切り替えロジック

□ Phase 6: 初期画面
  □ CategorySelectionコンポーネント
  □ 選択状態の引き継ぎ

□ Phase 7: 状態管理最適化
  □ Pinia導入（オプション）
  □ Store作成

□ Phase 8: jQuery削除
  □ jQuery依存コード置き換え
  □ 最終テスト
  □ 本番リリース
```

---

**計画書バージョン**: 1.0  
**最終更新**: 2025年12月2日  
**作成者**: GitHub Copilot  
**レビュー**: 未実施
