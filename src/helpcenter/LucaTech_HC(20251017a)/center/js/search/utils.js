/**
 * 共通ユーティリティ関数モジュール (Vue.js対応版)
 * - HTMLエスケープ
 * - 文字列処理
 * - その他の汎用機能
 * 
 * このモジュールはVue.jsとの互換性を保ちつつ、
 * jQuery依存を最小限にするよう設計されています
 */

/**
 * HTMLエスケープ関数
 * 特殊文字をHTMLエンティティに変換してXSS攻撃を防ぐ
 * @param {string} unsafe - エスケープする文字列
 * @returns {string} エスケープされた文字列
 */
function escapeHtml(unsafe) {
    if (typeof unsafe !== 'string') return '';
    
    const escapeMap = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    
    return unsafe.replace(/[&<>"']/g, char => escapeMap[char]);
}

/**
 * チェックボックスの階層的な状態管理関数 (Vue.js対応版)
 * 子チェックボックスの状態に応じて親チェックボックスの状態を更新し、
 * 必要に応じて再帰的に上位階層も更新する
 * 
 * jQueryからネイティブDOMメソッドに変換済み
 * 
 * @param {Element} node - チェック状態が変更されたチェックボックス要素
 */
function checkAllInTree(node) {
    // 最も近い<ul>要素を取得し、その親の<li>要素内の.search-inを探す
    const closestUl = node.closest("ul");
    if (!closestUl) return;
    
    const parentLi = closestUl.closest("li");
    if (!parentLi) return;
    
    const parentCheckbox = parentLi.querySelector(".search-in");
    if (!parentCheckbox) return;
    
    // 同階層の全ての兄弟要素の状態を確認
    const siblings = Array.from(closestUl.querySelectorAll(".search-in"));
    const checkedStates = siblings.map(checkbox => checkbox.checked);
    
    const allChecked = checkedStates.every(state => state);
    const someChecked = checkedStates.some(state => state);
    const allUnchecked = checkedStates.every(state => !state);
    
    // 親要素の参照を取得
    const parentDiv = parentCheckbox.closest("div");
    const parentLiForAll = parentDiv ? parentDiv.closest("li") : null;
    
    let allCheckbox = null;
    let allDiv = null;
    
    if (parentLiForAll) {
        const firstLi = parentLiForAll.querySelector("ul li:first-child");
        if (firstLi) {
            allCheckbox = firstLi.querySelector(".search-in-all");
            allDiv = firstLi.querySelector("div.custom-checkbox");
        }
    }
    
    // 状態に応じて更新
    if (allChecked) {
        parentCheckbox.checked = true;
        if (allCheckbox) allCheckbox.checked = true;
        if (parentDiv) parentDiv.classList.remove("check-new");
        if (allDiv) allDiv.classList.remove("check-new");
    } else if (someChecked) {
        parentCheckbox.checked = false;
        if (allCheckbox) allCheckbox.checked = false;
        if (parentDiv) parentDiv.classList.add("check-new");
        if (allDiv) allDiv.classList.add("check-new");
    } else if (allUnchecked) {
        parentCheckbox.checked = false;
        if (allCheckbox) allCheckbox.checked = false;
        if (parentDiv) parentDiv.classList.remove("check-new");
        if (allDiv) allDiv.classList.remove("check-new");
    }
    
    // 再帰的に上位階層も更新
    checkAllInTree(parentCheckbox);
}

/**
 * 文字列内でキーワードを検索し、該当箇所の前後のコンテキストを返す
 * @param {string} str - 検索対象の文字列
 * @param {Array<string>} keywords - 検索キーワードの配列
 * @returns {string} マッチした箇所の前後コンテキスト（"..."区切り）
 */
function searchKeywordsInString(str, keywords) {
    if (!keywords[0]?.length) return "";
    
    let result = "";
    let lastEnd = -1;
    
    for (let keyword of keywords) {
        const regex = new RegExp(keyword, "g");
        const match = regex.exec(str);
        
        if (match) {
            const index = match.index;
            const length = match[0].length;
            const start = Math.max(0, index - 10);
            const end = Math.min(str.length, index + length + 100);
            
            result += start <= lastEnd 
                ? "..." + str.slice(lastEnd + 1, end)
                : "..." + str.slice(start, end);
            
            lastEnd = end;
        }
    }

    return result;
}

/**
 * 検索機能で使用するツリー構造のデータモデルを構築する関数
 * ツリー構造のノードデータを検索システム用のJSONオブジェクトに変換
 * @param {Object} node - 変換元のノードデータ（ID, TITLE, CONTENTS, PATH等を含む）
 * @returns {Object} 検索用に最適化されたツリー構造のJSONオブジェクト
 */
function buildSearchTreeModel(node) {
    const json = {
        id: node.ID,
        title: node.TITLE,
        checked: true
    };
    
    if (node.CONTENTS !== undefined) {
        json.childs = node.CONTENTS
            .map(child => buildSearchTreeModel(child))
            .filter(model => model !== undefined);
    } else {
        json.baseUrl = node.PATH;
        json.searchjs = `${node.PATH}/search.js`;
    }
    
    return json;
}

/**
 * 検索キーワードを正規化
 * @param {string} keyword - 元のキーワード
 * @returns {Array<string>} 正規化されたキーワード配列
 */
function normalizeSearchKeyword(keyword) {
    if (!keyword) return [];
    
    let normalized = escapeHtml(keyword)
        .replace(/(.*?)(?:　| )+(.*?)/g, "$1 $2")
        .trim()
        .toLowerCase();
    
    wide.forEach((w, i) => {
        normalized = normalized.split(w).join(narrow[i]);
    });
    
    return normalized.split(" ");
}

/**
 * キー入力イベントの遅延実行用関数（デバウンス）
 * 連続したキー入力を制御し、一定時間後に処理を実行する
 * 
 * Vue.jsのイベントハンドラでも使用可能
 * 
 * @param {Function} func - 実行する関数
 * @param {number} wait - 待機時間（ミリ秒）
 * @returns {Function} デバウンスされた関数
 */
function throttle(func, wait) {
    let timeoutId = null;
    
    return function (...args) {
        const context = this;
        
        // 既存のタイマーをクリア
        if (timeoutId !== null) {
            clearTimeout(timeoutId);
        }
        
        // 指定した待機時間後に関数を実行
        timeoutId = setTimeout(() => {
            func.apply(context, args);
            timeoutId = null;
        }, wait);
    };
}

/**
 * デバウンス関数のエイリアス
 * より適切な名前として提供
 * @param {Function} func - 実行する関数
 * @param {number} wait - 待機時間（ミリ秒）
 * @returns {Function} デバウンスされた関数
 */
function debounce(func, wait) {
    return throttle(func, wait);
}

// 他のユーティリティ関数を今後追加する場合はここに記述
// 例: 日付フォーマット、数値フォーマット、バリデーション関数など

/**
 * 将来のモジュール化のためのエクスポート宣言
 * ES6モジュールとして使用する場合は、以下のコメントを解除してください
 */
/*
export {
    escapeHtml,
    checkAllInTree,
    searchKeywordsInString,
    buildSearchTreeModel,
    normalizeSearchKeyword,
    throttle,
    debounce
};
*/