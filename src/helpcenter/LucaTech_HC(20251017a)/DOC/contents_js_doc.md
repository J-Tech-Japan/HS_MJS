# contents.js 役割説明書

## 概要

`center/js/core/contents.js` は、LucaTech ヘルプシステム全体のコンテンツ構造とメタデータを管理するマスターファイルです。このファイルはヘルプシステムの基盤となる重要な役割を担っています。

## 主な機能

### 1. コンテンツカタログの定義

#### CONTENTS_ALL配列
システム全体で利用可能なすべてのヘルプコンテンツを階層構造で定義しています。

```javascript
CONTENTS_ALL = [
    {
        ID: "CMN",               // カテゴリID
        TYPE: "category",        // タイプ（category/contents）
        TITLE: "共通",           // 表示タイトル
        STYLE: "0",             // スタイル識別子
        CONTENTS: [             // 子コンテンツ配列
            {
                ID: "CMN_OPE",
                TITLE: "『LucaTech GX』について",
                TYPE: "contents",
                PATH: "../contents/CMN_OPE/"
            },
            // ...その他のコンテンツ
        ]
    },
    // ...その他のカテゴリ
];
```

#### データ構造の特徴
- **階層構造**: カテゴリとコンテンツの2階層で構成
- **TYPE識別**: `category`（カテゴリ）と`contents`（実コンテンツ）を区別
- **PATH指定**: 各コンテンツの実際のファイルパスを定義
- **一意ID**: 各要素に固有のIDを割り当て

### 2. ライセンス管理機能

#### Cookieベースのライセンス制御
```javascript
var cookie = document.cookie.split('; ').reduce((prev, current) => {
    const [name, ...value] = current.split('=');
    prev[name] = value.join('=');
    return prev;
}, {});

var gi_license = "gi_license" in cookie ? cookie.gi_license.split(",") : [];
```

#### 動的コンテンツフィルタリング
```javascript
CONTENTS = CONTENTS_ALL.map(category => {
    return {
        ...category,
        CONTENTS: category.CONTENTS.filter(contents => 
            contents.TYPE === "category" 
                ? contents.CONTENTS.some(sub => gi_license.includes(sub.ID))
                : gi_license.includes(contents.ID)
        ).map(contents => {
            return contents.TYPE === "category" ? {
                ...contents,
                CONTENTS: contents.CONTENTS.filter(sub => gi_license.includes(sub.ID))
            } : contents;
        })
    };
}).filter(category => category.CONTENTS.length);
```

この処理により：
- ユーザーが保有するライセンスに応じてコンテンツを動的に表示/非表示
- 未ライセンス機能へのアクセスを制御
- カスタマイズされたヘルプ環境を提供

### 3. 他システムとの連携

#### パス管理
```javascript
var baseFolderWebHelp = "../contents";
```
- コンテンツファイルの基準パスを一元管理
- 将来的なディレクトリ構造変更に対する柔軟性を確保

#### データ変換の起点
`contents.js`で定義されたデータは以下の流れで活用されます：

1. **CONTENTS** → 他のJSファイルで検索用データに変換
2. **変換後データ** → `localStorage['contents']`に保存
3. **localStorage** → 検索ページで読み込み・活用

## システム内での位置づけ

### データフロー図
```
contents.js (CONTENTS_ALL)
    ↓ ライセンスフィルタリング
contents.js (CONTENTS)
    ↓ 検索用データ変換
localStorage['contents']
    ↓ 読み込み
search.html (検索機能)
```

### 関連ファイルとの関係

| ファイル | 関係 | 役割 |
|---------|------|------|
| `utils.js` | データ変換 | CONTENTSを検索用データに変換 |
| `searchPageInit.js` | データ読み込み | localStorageからデータ取得 |
| `searchCatalog.js` | 検索実行 | 変換されたデータで検索処理 |
| `menuRender.js` | メニュー表示 | CONTENTSを使用したナビゲーション構築 |

## カスタマイズ方法

### 新しいコンテンツの追加
1. `CONTENTS_ALL`配列に新しいオブジェクトを追加
2. 対応する`contents/[ID]/`フォルダを作成
3. `search.js`ファイルを含むコンテンツファイルを配置

### ライセンス制御の変更
1. `gi_license`変数の処理ロジックを変更
2. フィルタリング条件をカスタマイズ

## 注意点

- **データ整合性**: IDの重複や循環参照に注意
- **パス管理**: `baseFolderWebHelp`の変更時は全体への影響を確認
- **ライセンス依存**: Cookie情報の変更がコンテンツ表示に直接影響
- **階層制限**: 現在は2階層まで対応（category → contents）

## メンテナンス

### 定期確認項目
1. 新機能追加時のCONTENTS_ALL更新
2. ライセンス体系変更時のフィルタリングロジック見直し
3. ディレクトリ構造変更時のパス設定確認
4. パフォーマンス影響の監視（大規模コンテンツ増加時）

このファイルはヘルプシステムの根幹を担っているため、変更時は十分なテストと影響範囲の確認が必要です。