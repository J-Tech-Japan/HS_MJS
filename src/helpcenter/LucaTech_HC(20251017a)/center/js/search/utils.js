/**
 * 共通ユーティリティ関数モジュール
 * - HTMLエスケープ
 * - 文字列処理
 * - その他の汎用機能
 */

// ========================================
// 文字変換配列（新しい形式のコンテンツ対応）
// ========================================
const wide = ["０", "１", "２", "３", "４", "５", "６", "７", "８", "９", "ａ", "ｂ", "ｃ", "ｄ", "ｅ", "ｆ", "ｇ", "ｈ", "ｉ", "ｊ", "ｋ", "ｌ", "ｍ", "ｎ", "ｏ", "ｐ", "ｑ", "ｒ", "ｓ", "ｔ", "ｕ", "ｖ", "ｗ", "ｘ", "ｙ", "ｚ", "ガ", "ギ", "グ", "ゲ", "ゴ", "ザ", "ジ", "ズ", "ゼ", "ゾ", "ダ", "ヂ", "ヅ", "デ", "ド", "バ", "ビ", "ブ", "ベ", "ボ", "パ", "ピ", "プ", "ペ", "ポ", "。", "「", "」", "、", "ヲ", "ァ", "ィ", "ゥ", "ェ", "ォ", "ャ", "ュ", "ョ", "ッ", "ー", "ア", "イ", "ウ", "エ", "オ", "カ", "キ", "ク", "ケ", "コ", "サ", "シ", "ス", "セ", "ソ", "タ", "チ", "ツ", "テ", "ト", "ナ", "ニ", "ヌ", "ネ", "ノ", "ハ", "ヒ", "フ", "ヘ", "ホ", "マ", "ミ", "ム", "メ", "モ", "ヤ", "ユ", "ヨ", "ラ", "リ", "ル", "レ", "ロ", "ワ", "ン"];
const narrow = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "ｶﾞ", "ｷﾞ", "ｸﾞ", "ｹﾞ", "ｺﾞ", "ｻﾞ", "ｼﾞ", "ｽﾞ", "ｾﾞ", "ｿﾞ", "ﾀﾞ", "ﾁﾞ", "ﾂﾞ", "ﾃﾞ", "ﾄﾞ", "ﾊﾞ", "ﾋﾞ", "ﾌﾞ", "ﾍﾞ", "ﾎﾞ", "ﾊﾟ", "ﾋﾟ", "ﾌﾟ", "ﾍﾟ", "ﾎﾟ", "｡", "｢", "｣", "､", "ｦ", "ｧ", "ｨ", "ｩ", "ｪ", "ｫ", "ｬ", "ｭ", "ｮ", "ｯ", "ｰ", "ｱ", "ｲ", "ｳ", "ｴ", "ｵ", "ｶ", "ｷ", "ｸ", "ｹ", "ｺ", "ｻ", "ｼ", "ｽ", "ｾ", "ｿ", "ﾀ", "ﾁ", "ﾂ", "ﾃ", "ﾄ", "ﾅ", "ﾆ", "ﾇ", "ﾈ", "ﾉ", "ﾊ", "ﾋ", "ﾌ", "ﾍ", "ﾎ", "ﾏ", "ﾐ", "ﾑ", "ﾒ", "ﾓ", "ﾔ", "ﾕ", "ﾖ", "ﾗ", "ﾘ", "ﾙ", "ﾚ", "ﾛ", "ﾜ", "ﾝ"];
const highlight = ["(?:０|0)", "(?:１|1)", "(?:２|2)", "(?:３|3)", "(?:４|4)", "(?:５|5)", "(?:６|6)", "(?:７|7)", "(?:８|8)", "(?:９|9)", "(?:ａ|a)", "(?:ｂ|b)", "(?:ｃ|c)", "(?:ｄ|d)", "(?:ｅ|e)", "(?:ｆ|f)", "(?:ｇ|g)", "(?:ｈ|h)", "(?:ｉ|i)", "(?:ｊ|j)", "(?:ｋ|k)", "(?:ｌ|l)", "(?:ｍ|m)", "(?:ｎ|n)", "(?:ｏ|o)", "(?:ｐ|p)", "(?:ｑ|q)", "(?:ｒ|r)", "(?:ｓ|s)", "(?:ｔ|t)", "(?:ｕ|u)", "(?:ｖ|v)", "(?:ｗ|w)", "(?:ｘ|x)", "(?:ｙ|y)", "(?:ｚ|z)", "(?:ガ|ｶﾞ)", "(?:ギ|ｷﾞ)", "(?:グ|ｸﾞ)", "(?:ゲ|ｹﾞ)", "(?:ゴ|ｺﾞ)", "(?:ザ|ｻﾞ)", "(?:ジ|ｼﾞ)", "(?:ズ|ｽﾞ)", "(?:ゼ|ｾﾞ)", "(?:ゾ|ｿﾞ)", "(?:ダ|ﾀﾞ)", "(?:ヂ|ﾁﾞ)", "(?:ヅ|ﾂﾞ)", "(?:デ|ﾃﾞ)", "(?:ド|ﾄﾞ)", "(?:バ|ﾊﾞ)", "(?:ビ|ﾋﾞ)", "(?:ブ|ﾌﾞ)", "(?:ベ|ﾍﾞ)", "(?:ボ|ﾎﾞ)", "(?:パ|ﾊﾟ)", "(?:ピ|ﾋﾟ)", "(?:プ|ﾌﾟ)", "(?:ペ|ﾍﾟ)", "(?:ポ|ﾎﾟ)", "(?:。|｡)", "(?:「|｢)", "(?:」|｣)", "(?:、|､)", "(?:ヲ|ｦ)", "(?:ァ|ｧ)", "(?:ィ|ｨ)", "(?:ゥ|ｩ)", "(?:ェ|ｪ)", "(?:ォ|ｫ)", "(?:ャ|ｬ)", "(?:ュ|ｭ)", "(?:ョ|ｮ)", "(?:ッ|ｯ)", "(?:ー|ｰ)", "(?:ア|ｱ)", "(?:イ|ｲ)", "(?:ウ|ｳ)", "(?:エ|ｴ)", "(?:オ|ｵ)", "(?:カ|ｶ)", "(?:キ|ｷ)", "(?:ク|ｸ)", "(?:ケ|ｹ)", "(?:コ|ｺ)", "(?:サ|ｻ)", "(?:シ|ｼ)", "(?:ス|ｽ)", "(?:セ|ｾ)", "(?:ソ|ｿ)", "(?:タ|ﾀ)", "(?:チ|ﾁ)", "(?:ツ|ﾂ)", "(?:テ|ﾃ)", "(?:ト|ﾄ)", "(?:ナ|ﾅ)", "(?:ニ|ﾆ)", "(?:ヌ|ﾇ)", "(?:ネ|ﾈ)", "(?:ノ|ﾉ)", "(?:ハ|ﾊ)", "(?:ヒ|ﾋ)", "(?:フ|ﾌ)", "(?:ヘ|ﾍ)", "(?:ホ|ﾎ)", "(?:マ|ﾏ)", "(?:ミ|ﾐ)", "(?:ム|ﾑ)", "(?:メ|ﾒ)", "(?:モ|ﾓ)", "(?:ヤ|ﾔ)", "(?:ユ|ﾕ)", "(?:ヨ|ﾖ)", "(?:ラ|ﾗ)", "(?:リ|ﾘ)", "(?:ル|ﾙ)", "(?:レ|ﾚ)", "(?:ロ|ﾛ)", "(?:ワ|ﾜ)", "(?:ン|ﾝ)"];

// 効率化のための変換マップ（初期化時に一度だけ生成）
const wideToNarrowMap = (() => {
    const map = {};
    // 全角アルファベット→半角アルファベット（インデックス10-35）
    for (let i = 10; i < 36; i++) {
        map[wide[i]] = narrow[i];
    }
    return map;
})();

const narrowToWideMap = (() => {
    const map = {};
    // 半角カタカナ→全角カタカナ（インデックス36以降）
    for (let i = 36; i < narrow.length; i++) {
        map[narrow[i]] = wide[i];
    }
    return map;
})();

// 効率的な変換用の正規表現（初期化時に一度だけ生成）
const wideAlphabetRegex = new RegExp(`[${Object.keys(wideToNarrowMap).join('')}]`, 'g');
const narrowKatakanaRegex = new RegExp(
    Object.keys(narrowToWideMap)
        .map(k => k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
        .join('|'), 
    'g'
);

// ========================================
// jQueryカスタムセレクタ（大文字小文字を区別しない検索）
// ========================================
$.expr.pseudos.containsNormalized = $.expr.createPseudo(function(arg) {
    return function(elem) {
        const normalizedElemText = $(elem).text().toLowerCase();
        const normalizedSearchText = arg.toLowerCase();
        return normalizedElemText.includes(normalizedSearchText);
    };
});

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
 * セレクタ用のエスケープ処理
 * 正規表現の特殊文字をエスケープして安全な文字列にする
 * @param {string} val - エスケープする文字列
 * @returns {string} エスケープされた文字列
 */
function selectorEscape(val) {
    return val.replace(/[-\/\\^$*+?.()|[\]{}\!]/g, '\\$&');
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
 * 検索キーワードを正規化（最適化版）
 * 全角アルファベット→半角、半角カタカナ→全角に変換し、スペース区切りで配列化
 * @param {string} keyword - 元のキーワード
 * @returns {Array<string>} 正規化されたキーワード配列
 */
function normalizeSearchKeyword(keyword) {
    if (!keyword) return [];
    
    let normalized = escapeHtml(keyword)
        .replace(/(.*?)(?:　| )+(.*?)/g, "$1 $2")
        .trim()
        .toLowerCase();
    
    // 全角アルファベット→半角アルファベットに変換（事前生成された正規表現とマップを使用）
    normalized = normalized.replace(wideAlphabetRegex, match => wideToNarrowMap[match]);
    
    // 半角カタカナ→全角カタカナに変換（事前生成された正規表現とマップを使用）
    normalized = normalized.replace(narrowKatakanaRegex, match => narrowToWideMap[match]);
    
    // 正規化後の値をスペース区切りで配列化
    return normalized.split(/\s+/).filter(word => word.length > 0);
}