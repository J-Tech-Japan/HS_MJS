// コンテンツタイプの定義
const CONTENT_TYPE = {
    CATEGORY: "category",    // カテゴリタイプ
    CONTENTS: "contents"     // コンテンツタイプ
};

// ベースフォルダのパス設定
var baseFolderWebHelp = "../contents";

// コンテンツパスを生成する関数
const createPath = (id) => `${baseFolderWebHelp}/${id}/`;

// 全コンテンツの定義（ライセンスフィルタ適用前）
const CONTENTS_ALL = [
    // 共通カテゴリ
    {
        ID: "CMN",
        TYPE: CONTENT_TYPE.CATEGORY,
        TITLE: "共通",
        STYLE: "0",
        CONTENTS: [
            {
                ID: "CMN_OPE",
                TITLE: "『LucaTech GX』について",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("CMN_OPE")
            },
            {
                ID: "CMN_JNT",
                TITLE: "導入処理",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("CMN_JNT")
            },
            {
                ID: "CMN_MAS",
                TITLE: "マスター登録処理",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("CMN_MAS")
            },
            {
                ID: "CMN_CMP",
                TITLE: "会社データ管理",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("CMN_CMP")
            },
            {
                ID: "CMN_UTL",
                TITLE: "システム管理",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("CMN_UTL")
            },
            {
                ID: "CMN_DAT",
                TITLE: "データ交換",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("CMN_DAT")
            },
        ]
    },
    // 財務会計システムカテゴリ
    {
        ID: "MAS",
        TYPE: CONTENT_TYPE.CATEGORY,
        TITLE: "財務会計システム",
        STYLE: "1",
        CONTENTS: [
            {
                ID: "MAS_OPE",
                TITLE: "財務会計システムについて",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("MAS_OPE")
            },
            {
                ID: "MAS_sub01",
                TITLE: "導入処理",
                TYPE: CONTENT_TYPE.CATEGORY,
                CONTENTS: [
                    {
                        ID: "MAS_JNT",
                        TITLE: "基本設定",
                        TYPE: CONTENT_TYPE.CONTENTS,
                        PATH: createPath("MAS_JNT")
                    },
                    {
                        ID: "MAS_DEP",
                        TITLE: "部門設定",
                        TYPE: CONTENT_TYPE.CONTENTS,
                        PATH: createPath("MAS_DEP")
                    },
                    {
                        ID: "MAS_BAB",
                        TITLE: "残高・予算設定",
                        TYPE: CONTENT_TYPE.CONTENTS,
                        PATH: createPath("MAS_BAB")
                    },
                ]
            },
            {
                ID: "MAS_DAY",
                TITLE: "日常処理",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("MAS_DAY")
            },
            {
                ID: "MAS_SOA",
                TITLE: "決算処理",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("MAS_SOA")
            },
            {
                ID: "MAS_EAL",
                TITLE: "電子帳簿",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("MAS_EAL")
            },
        ]
    },
    // 固定資産システムカテゴリ
    {
        ID: "DEP",
        TYPE: CONTENT_TYPE.CATEGORY,
        TITLE: "固定資産システム",
        STYLE: "2",
        CONTENTS: [
            {
                ID: "DEP_JNT",
                TITLE: "導入処理",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("DEP_JNT")
            },
            {
                ID: "DEP_SIS",
                TITLE: "資産管理",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("DEP_SIS")
            },
        ]
    },
    // ワークフローシステムカテゴリ
    {
        ID: "FRT",
        TYPE: CONTENT_TYPE.CATEGORY,
        TITLE: "ワークフローシステム",
        STYLE: "5",
        CONTENTS: [
            {
                ID: "WFL_JNT",
                TITLE: "導入処理",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("WFL_JNT")
            },
            {
                ID: "WFL_FOM",
                TITLE: "フォーム設定",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("WFL_FOM")
            },
            {
                ID: "WFL_DAY",
                TITLE: "日常処理",
                TYPE: CONTENT_TYPE.CONTENTS,
                PATH: createPath("WFL_DAY")
            },
        ]
    }
];

/**
 * Cookieを解析してオブジェクトに変換する関数
 * @returns {Object} Cookie名と値のペアを持つオブジェクト
 */
function parseCookies() {
    try {
        // Cookieが存在しない場合は空オブジェクトを返す
        if (!document.cookie) {
            return {};
        }
        // Cookieを分割してオブジェクトに変換
        return document.cookie.split('; ').reduce((prev, current) => {
            const [name, ...value] = current.split('=');
            if (name) {
                prev[name] = value.join('=');
            }
            return prev;
        }, {});
    } catch (error) {
        console.error('Cookie parsing failed:', error);
        return {};
    }
}

/**
 * ライセンスCookieを解析してIDの配列に変換する関数
 * @param {Object} cookies - Cookieオブジェクト
 * @returns {Array<string>} ライセンスIDの配列
 */
function parseLicenseCookie(cookies) {
    const licenseValue = cookies['gi_license'];
    // ライセンス値が存在しない、または空の場合は空配列を返す
    if (!licenseValue || licenseValue.trim() === '') {
        return [];
    }
    // カンマ区切りのライセンスIDを配列に変換
    return licenseValue.split(',').map(id => id.trim()).filter(Boolean);
}

// Cookieを解析
var cookie = parseCookies();

// ライセンス情報を取得
var gi_license = parseLicenseCookie(cookie);

/**
 * コンテンツまたはカテゴリがライセンスでアクセス可能かチェック
 * @param {Object} contents - チェック対象のコンテンツオブジェクト
 * @param {Array<string>} licenseIds - 有効なライセンスIDの配列
 * @returns {boolean} アクセス可能な場合true
 */
function isContentAccessible(contents, licenseIds) {
    if (contents.TYPE === CONTENT_TYPE.CATEGORY) {
        // カテゴリの場合、サブコンテンツのいずれかがアクセス可能ならtrue
        return contents.CONTENTS.some(sub => licenseIds.includes(sub.ID));
    }
    // 通常のコンテンツの場合、そのIDがライセンスに含まれているかチェック
    return licenseIds.includes(contents.ID);
}

/**
 * カテゴリ内のサブコンテンツをライセンスでフィルタリング
 * @param {Object} contents - フィルタリング対象のコンテンツオブジェクト
 * @param {Array<string>} licenseIds - 有効なライセンスIDの配列
 * @returns {Object} フィルタリング済みのコンテンツオブジェクト
 */
function filterCategoryContents(contents, licenseIds) {
    if (contents.TYPE === CONTENT_TYPE.CATEGORY) {
        return {
            ...contents,
            CONTENTS: contents.CONTENTS.filter(sub => licenseIds.includes(sub.ID))
        };
    }
    return contents;
}

/**
 * カテゴリのコンテンツ配列をライセンスでフィルタリング
 * @param {Array<Object>} contentsList - カテゴリ内のコンテンツ配列
 * @param {Array<string>} licenseIds - 有効なライセンスIDの配列
 * @returns {Array<Object>} フィルタリング済みのコンテンツ配列
 */
function filterContentsByLicense(contentsList, licenseIds) {
    return contentsList
        .filter(contents => isContentAccessible(contents, licenseIds))
        .map(contents => filterCategoryContents(contents, licenseIds));
}

/**
 * ライセンスに基づいてフィルタリングされたコンテンツ配列
 * 各カテゴリとそのコンテンツをライセンスIDでフィルタリングし、
 * アクセス可能なコンテンツのみを含む配列を生成
 */
const CONTENTS = CONTENTS_ALL
    .map(category => ({
        ...category,
        CONTENTS: filterContentsByLicense(category.CONTENTS, gi_license)
    }))
    .filter(category => category.CONTENTS.length > 0);