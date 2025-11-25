# ツリービュー構築関数ドキュメント

## 概要

検索画面の左メニューに表示されるカタログツリービューを構築する関数群です。
親子関係を明確にするため、以下のような命名規則を採用しています：

- **`buildTreeview()`**: ツリービュー全体を構築するメイン関数（エントリーポイント）
- **`buildTreeviewNode(node)`**: 個々のノードをHTML化する補助関数（再帰処理）

---

## 関数詳細

### `buildTreeview()`

**役割**: ツリービュー全体を構築・表示するエントリーポイント関数

**処理フロー**:
1. 左メニューコンテナ（`.column-left-container`）をクリア
2. `getSearchCatalogue()` から検索カタログデータを取得
3. 各カタログに対して `buildTreeviewNode()` を呼び出し、HTML文字列を生成
4. 生成したHTMLをDOMに一括挿入
5. `setupTreeviewEventHandlers()` でイベントハンドラーを設定

**コード**:
```javascript
function buildTreeview() {
    const $container = $(".column-left > .column-left-container");
    $container.empty();
    
    const searchCatalogue = getSearchCatalogue();
    
    // DOM操作を一度にまとめる
    const htmlFragments = searchCatalogue.map(catalogue => 
        `<div class='box-check'><ul>${buildTreeviewNode(catalogue)}</ul></div>`
    );
    
    $container.html(htmlFragments.join(''));
    
    // イベントハンドラーを設定
    setupTreeviewEventHandlers();
}
```

**呼び出し箇所**:
- `searchPageInit.js`: ページ初期化時
- `searchUI.js` の `addHandleEventInFirstPage()`: チェックボックス変更時（2箇所）

---

### `buildTreeviewNode(node)`

**役割**: 個々のノードを再帰的にHTML文字列に変換する補助関数

**パラメータ**:
- `node` (Object): カタログノードオブジェクト
  - `id` (string): ノードID
  - `title` (string): ノードタイトル
  - `childs` (Array): 子ノードの配列（オプション）
  - `checked` (boolean): チェック状態（オプション）

**戻り値**:
- (string): `<li>...</li>` 形式のHTML文字列

**処理フロー**:
1. **入力値検証**: ノードオブジェクトの妥当性をチェック
2. **子要素の有無を判定**: `node.childs` が存在するかチェック
3. **親ノードの場合**:
   - チェックボックス要素を生成（`has-childs root` クラス付き）
   - トグルボタン（`.check-toggle`）を追加
   - 「すべて選択」オプションを追加
   - **再帰呼び出し**: 各子ノードに対して `buildTreeviewNode()` を実行
4. **リーフノードの場合**:
   - チェックボックス要素のみを生成

**コード**:
```javascript
function buildTreeviewNode(node) {
    // 入力値検証
    if (!node || typeof node !== 'object') {
        console.warn('buildTreeviewNode: Invalid node object');
        return '<li></li>';
    }
    
    if (!node.id || !node.title) {
        console.warn('buildTreeviewNode: Missing required node properties (id, title)');
        return '<li></li>';
    }
    
    const hasChildren = node.childs && Array.isArray(node.childs) && node.childs.length > 0;
    
    let treeview = '<li>';
    
    if (hasChildren) {
        // 親ノード（子要素あり）の場合
        treeview += createCheckboxElement(node, 'has-childs root', false);
        treeview += '<div class="check-toggle active"></div>';
        treeview += '<ul class="box-item box-toggle">';
        
        // 「すべて選択」オプション
        treeview += '<li>';
        treeview += createCheckboxElement(node, '', true);
        treeview += '</li>';
        
        // 子ノードを再帰的に構築
        for (const childNode of node.childs) {
            treeview += buildTreeviewNode(childNode);
        }
        
        treeview += '</ul>';
    } else {
        // リーフノード（子要素なし）の場合
        treeview += createCheckboxElement(node);
    }
    
    treeview += '</li>';
    return treeview;
}
```

**呼び出し箇所**:
- `buildTreeview()`: ルートカタログノードの処理
- `buildTreeviewNode()` 自身: 子ノードの再帰処理

---

## 親子関係の図解

```
buildTreeview()  ← メイン関数（エントリーポイント）
    ├─ getSearchCatalogue() を取得
    ├─ 各カタログに対して処理
    │   └─ buildTreeviewNode(catalogue)  ← 再帰関数
    │       ├─ チェックボックス要素生成
    │       ├─ 子ノードがある場合
    │       │   ├─ トグルボタン追加
    │       │   ├─ 「すべて選択」追加
    │       │   └─ for (childNode of childs)
    │       │       └─ buildTreeviewNode(childNode)  ← 再帰呼び出し
    │       └─ リーフノードの場合
    │           └─ チェックボックスのみ生成
    ├─ 生成したHTMLをDOM挿入
    └─ setupTreeviewEventHandlers() 実行
```

---

## データ構造例

### 入力データ（`searchCatalogue`）:
```javascript
[
  {
    id: "CMN_JNT",
    title: "共通業務",
    checked: true,
    childs: [
      {
        id: "CMN00000",
        title: "ログイン・ログアウト",
        checked: true,
        childs: []  // リーフノード
      },
      {
        id: "CMN11000",
        title: "基本操作",
        checked: true,
        childs: [
          {
            id: "CMN11001",
            title: "画面の構成",
            checked: true,
            childs: []
          }
        ]
      }
    ]
  },
  {
    id: "MAS_JNT",
    title: "マスタ管理",
    checked: true,
    childs: [...]
  }
]
```

### 出力HTML（簡略化）:
```html
<div class="column-left-container">
  <div class='box-check'>
    <ul>
      <li>
        <div class="custom-checkbox leaf has-childs root">
          <input type="checkbox" checked class="custom-control-input search-in has-childs root" id="search-in-CMN_JNT">
          <label for="search-in-CMN_JNT" class="custom-control-label has-childs root">共通業務 <span class="count" id="count-CMN_JNT">(0)</span></label>
        </div>
        <div class="check-toggle active"></div>
        <ul class="box-item box-toggle">
          <li>
            <div class="custom-checkbox leaf">
              <input type="checkbox" checked class="custom-control-input search-in-all" id="search-in-all-CMN_JNT">
              <label for="search-in-all-CMN_JNT" class="custom-control-label">(すべて選択)</label>
            </div>
          </li>
          <li>
            <div class="custom-checkbox leaf">
              <input type="checkbox" checked class="custom-control-input search-in" id="search-in-CMN00000">
              <label for="search-in-CMN00000" class="custom-control-label">ログイン・ログアウト <span class="count" id="count-CMN00000">(0)</span></label>
            </div>
          </li>
          <!-- 他の子ノード... -->
        </ul>
      </li>
    </ul>
  </div>
  <!-- 他のカタログ... -->
</div>
```

---

## 依存関数

### `createCheckboxElement(node, additionalClass, isSelectAll)`
チェックボックスのHTML要素を生成する補助関数。

**パラメータ**:
- `node`: ノードオブジェクト
- `additionalClass`: 追加するCSSクラス（`'has-childs root'` など）
- `isSelectAll`: 「すべて選択」フラグ

### `getSearchCatalogue()`
検索カタログデータを取得する関数。カタログのツリー構造データを返す。

### `setupTreeviewEventHandlers()`
ツリービューのイベントハンドラー（トグル、チェックボックス変更など）を設定する関数。

---

## 設計上の注意点

### 1. 命名規則
- **メイン関数**: `buildXxx()` - 全体を構築
- **補助関数**: `buildXxxNode()` - 個々の要素を処理

### 2. パフォーマンス最適化
- DOM操作を最小限に: `htmlFragments.join('')` で一括挿入
- イベントデリゲーションの活用（`setupTreeviewEventHandlers()` 内）

### 3. 再帰処理の安全性
- 入力値検証により無限ループを防止
- 不正なノードは空の `<li></li>` を返してスキップ

### 4. 拡張性
- チェックボックス生成ロジックを `createCheckboxElement()` に分離
- カタログデータの構造変更に柔軟に対応可能

---

## 変更履歴

| 日付 | 変更内容 |
|------|---------|
| 2025-01-13 | 関数名を変更: `buildTreeview()` → `buildTreeviewNode()`, `buildTreeview2()` → `buildTreeview()` |
| 2025-01-13 | 親子関係を明確にするドキュメントを作成 |

---

## 関連ファイル

- **`searchUI.js`**: 本関数群の実装ファイル
- **`searchPageInit.js`**: `buildTreeview()` の呼び出し元
- **`utils.js`**: `escapeHtml()` などの共通ユーティリティ
- **`searchData.js`**: `getSearchCatalogue()` の実装（推定）
