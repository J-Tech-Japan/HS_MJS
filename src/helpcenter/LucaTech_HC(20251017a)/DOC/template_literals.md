# テンプレートリテラル（Template Literals）

## 概要

テンプレートリテラルは、ES6（ES2015）で導入されたJavaScriptの新しい文字列記法です。バッククォート（`）を使用して文字列を定義し、従来の文字列連結よりも読みやすく、保守しやすいコードを書くことができます。

## 基本構文

### 通常の文字列との比較

```javascript
// 従来の文字列
const str1 = "Hello World";
const str2 = 'Hello World';

// テンプレートリテラル
const str3 = `Hello World`;
```

### 変数の埋め込み（補間）

```javascript
const name = "山田太郎";
const age = 30;

// 従来の文字列連結
const message1 = "私の名前は" + name + "で、年齢は" + age + "歳です。";

// テンプレートリテラル
const message2 = `私の名前は${name}で、年齢は${age}歳です。`;
```

### 式の埋め込み

```javascript
const a = 10;
const b = 20;

// 計算結果の埋め込み
const result = `${a} + ${b} = ${a + b}`;
console.log(result); // "10 + 20 = 30"

// 関数呼び出しの埋め込み
const greeting = `こんにちは、${getName()}さん！`;

// 三項演算子の使用
const status = `ユーザーは${isLoggedIn ? 'ログイン中' : '未ログイン'}です`;
```

## 主な機能

### 1. 複数行文字列

```javascript
// 従来の方法
const html1 = '<div class="container">\n' +
              '  <h1>タイトル</h1>\n' +
              '  <p>内容</p>\n' +
              '</div>';

// テンプレートリテラル
const html2 = `
<div class="container">
  <h1>タイトル</h1>
  <p>内容</p>
</div>
`;
```

### 2. HTMLテンプレートの構築

```javascript
function createUserCard(user) {
  return `
    <div class="user-card">
      <img src="${user.avatar}" alt="${user.name}のアバター">
      <h3>${user.name}</h3>
      <p>年齢: ${user.age}歳</p>
      <p>職業: ${user.occupation}</p>
      ${user.isActive ? '<span class="status active">オンライン</span>' : '<span class="status inactive">オフライン</span>'}
    </div>
  `;
}
```

### 3. 配列操作との組み合わせ

```javascript
const items = ['りんご', 'バナナ', 'オレンジ'];

// リストHTMLの生成
const listHtml = `
<ul>
  ${items.map(item => `<li>${item}</li>`).join('')}
</ul>
`;
```

## プロジェクトでの活用例

### buildFirstPage関数での使用

```javascript
function buildFirstPage() {
    let html = "";
    const searchCatalogue = getSearchCatalogue();
    
    for (let i = 0; i < searchCatalogue.length; i++) {
        const catalogue = searchCatalogue[i];
        
        html += `
        <div class="box-s-1">
            <div class="h1 custom-checkbox">
                <input type="checkbox" 
                       checked 
                       class="custom-control-input parent" 
                       search="${catalogue.id}" 
                       id="search-in-top-${catalogue.id}">
                <label for="search-in-top-${catalogue.id}" 
                       class="custom-control-label">
                    ${escapeHtml(catalogue.title)}
                </label>
            </div>
            <div class="box-item">
                <ul class="box-ul box-ul-1">
                    ${catalogue.childs.map(child => `
                        <li>
                            <div class="custom-checkbox">
                                <input type="checkbox" 
                                       checked 
                                       class="custom-control-input child" 
                                       search="${child.id}" 
                                       id="search-in-top-${child.id}">
                                <label for="search-in-top-${child.id}" 
                                       class="custom-control-label">
                                    ${escapeHtml(child.title)}
                                </label>
                            </div>
                        </li>
                    `).join('')}
                </ul>
            </div>
        </div>`;
    }
    return html;
}
```

### createCheckboxElement関数での使用

```javascript
function createCheckboxElement(node, additionalClass = '', isSelectAll = false) {
    const checked = node.checked ? 'checked' : '';
    const id = isSelectAll ? `search-in-all-${node.id}` : `search-in-${node.id}`;
    const className = isSelectAll ? 'search-in-all' : 'search-in';
    const labelText = isSelectAll ? '(すべて選択)' : escapeHtml(node.title);
    const countSpan = isSelectAll ? '' : ` <span class="count" id="count-${node.id}">(0)</span>`;
    
    return `
        <div class="custom-checkbox leaf${additionalClass ? ' ' + additionalClass : ''}">
            <input type="checkbox" ${checked} class="custom-control-input ${className}${additionalClass ? ' ' + additionalClass : ''}" id="${id}">
            <label for="${id}" class="custom-control-label${additionalClass ? ' ' + additionalClass : ''}">${labelText}${countSpan}</label>
        </div>
    `;
}
```

## 利点

### 1. 可読性の向上
- 文字列連結（+）よりも自然で読みやすい
- HTMLテンプレートの構造が視覚的に分かりやすい

### 2. 保守性の向上
- 変数名の変更が容易
- IDEでの構文ハイライトが効く
- エラーの特定が簡単

### 3. パフォーマンス
- 文字列連結よりも高速（エンジンレベルでの最適化）

### 4. 機能性
- 複数行文字列が簡単
- 式の埋め込みが直感的
- 関数呼び出しやメソッドチェーンも可能

## 注意点

### 1. ブラウザサポート
- IE11以下では使用不可
- Babelなどのトランスパイルが必要な場合がある

### 2. セキュリティ
- XSS攻撃のリスクがあるため、ユーザー入力は適切にエスケープする
- `escapeHtml`関数の使用を忘れずに

```javascript
// 危険な例
const userInput = '<script>alert("XSS")</script>';
const html = `<div>${userInput}</div>`; // XSS脆弱性

// 安全な例
const html = `<div>${escapeHtml(userInput)}</div>`;
```

### 3. パフォーマンス考慮
- 大量のテンプレートを処理する場合は、キャッシュやバッチ処理を検討

## まとめ

テンプレートリテラルは、現代のJavaScript開発において必須の機能です。特にHTMLテンプレートの構築や動的文字列生成において、コードの可読性と保守性を大幅に改善できます。

プロジェクトでは既に効果的に活用されており、従来の文字列連結から移行することで、より良いコード品質を実現できています。