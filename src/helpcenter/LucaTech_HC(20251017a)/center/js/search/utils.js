/**
 * 共通ユーティリティ関数モジュール
 * - HTMLエスケープ
 * - 文字列処理
 * - その他の汎用機能
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
 * チェックボックスの階層的な状態管理関数
 * 子チェックボックスの状態に応じて親チェックボックスの状態を更新し、
 * 必要に応じて再帰的に上位階層も更新する
 * @param {Element} node - チェック状態が変更されたチェックボックス要素
 */
function checkAllInTree(node) {
    const $node = $(node);
    const $parent = $($node.closest("ul").closest("li").find(".search-in")[0]);
    
    if ($parent.length === 0) return;
    
    // 同階層の全ての兄弟要素の状態を確認
    const $siblings = $node.closest("ul").find(".search-in");
    const checkedStates = $siblings.map(function() { 
        return $(this).is(":checked"); 
    }).get();
    
    const allChecked = checkedStates.every(state => state);
    const someChecked = checkedStates.some(state => state);
    const allUnchecked = checkedStates.every(state => !state);
    
    // 親要素とすべて選択要素の共通参照
    const $parentDiv = $parent.closest("div");
    const $allCheckbox = $parentDiv.closest("li").find("ul li:first .search-in-all");
    const $allDiv = $parentDiv.closest("li").find("ul li:first div.custom-checkbox");
    
    // 状態に応じて更新
    if (allChecked) {
        $parent.prop("checked", true);
        $allCheckbox.prop("checked", true);
        $parentDiv.add($allDiv).removeClass("check-new");
    } else if (someChecked) {
        $parent.prop("checked", false);
        $allCheckbox.prop("checked", false);
        $parentDiv.add($allDiv).addClass("check-new");
    } else if (allUnchecked) {
        $parent.prop("checked", false);
        $allCheckbox.prop("checked", false);
        $parentDiv.add($allDiv).removeClass("check-new");
    }
    
    checkAllInTree($parent[0]); // 再帰的に上位階層も更新
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
 * キー入力イベントの遅延実行用関数（スロットリング）
 * 連続したキー入力を制御し、一定時間後に処理を実行する
 * @param {Function} func - 実行する関数
 * @param {number} wait - 待機時間（ミリ秒）
 * @returns {Function} スロットリングされた関数
 */
function throttle(func, wait) {
    return function () {
        var that = this,
            args = [].slice.call(arguments);

        // 既存のタイマーをクリア
        clearTimeout(func._throttleTimeout);

        // 指定した待機時間後に関数を実行
        func._throttleTimeout = setTimeout(function () {
            func.apply(that, args);
        }, wait);
    };
}

// 他のユーティリティ関数を今後追加する場合はここに記述
// 例: 日付フォーマット、数値フォーマット、バリデーション関数など