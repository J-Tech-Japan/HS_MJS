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
    // システム共通カテゴリ
    {
        ID: "CMN",
        TYPE: CONTENT_TYPE.CATEGORY,
        TITLE: "システム共通",
        STYLE: "0",
        CONTENTS: [
            {
                ID: "CMN_OPE",
                TITLE: "LucaTech GXについて",
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
    // 財務会計（財務大将）カテゴリ
    {
        ID: "MAS",
        TYPE: CONTENT_TYPE.CATEGORY,
        TITLE: "財務大将",
        STYLE: "1",
        CONTENTS: [
            {
                ID: "MAS_OPE",
                TITLE: "財務大将について",
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
    // 固定資産管理カテゴリ
    {
        ID: "DEP",
        TYPE: CONTENT_TYPE.CATEGORY,
        TITLE: "固定資産管理",
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
    // ワークフローカテゴリ
    {
        ID: "FRT",
        TYPE: CONTENT_TYPE.CATEGORY,
        TITLE: "ワークフロー",
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
LicenseCodeModel = {
    "CMN_CMP": ["0046710001"],
    "CMN_DAT": ["0046710130", "0046735030"],
    "CMN_JNT": ["0046710001"],
    "CMN_MAS": ["0046710001"],
    "CMN_OPE": ["0046710001"],
    "CMN_UTL": ["0046710001"],
    "DEP_JNT": ["0046735010"],
    "DEP_SIS": ["0046735010"],
    "MAS_BAB": ["0046710010"],
    "MAS_DAY": ["0046710010"],
    "MAS_DEP": ["0046710010"],
    "MAS_EAL": ["0046710010"],
    "MAS_JNT": ["0046710010"],
    "MAS_OPE": ["0046710010"],
    "MAS_SOA": ["0046710010"],
    "WFL_DAY": ["0046735130"],
    "WFL_FOM": ["0046735130"],
    "WFL_JNT": ["0046735130"]
};


var cookie = document.cookie.split('; ').reduce((prev, current) => {
    const [name, ...value] = current.split('=');
    prev[name] = value.join('=');
    return prev;
}, {});

var gilicense = "gilicense" in cookie ? cookie.gilicense.split("%2C") : [];
let codeList = [];

for (const key in LicenseCodeModel) {
    const values = LicenseCodeModel[key];
    for (const target of gilicense) {
        if (values.includes(target)) {
            const prefix = key.split("/")[0];
            codeList.push(prefix);
            break; // 同じキーに対して複数回追加しないように
        }
    }
}

const uniqueCodeList = [...new Set(codeList)];
CONTENTS = CONTENTS_ALL.map(category => {
    return {
        ...category,
        CONTENTS: category.CONTENTS.filter(contents => contents.TYPE === CONTENT_TYPE.CATEGORY ? contents.CONTENTS.some(sub => uniqueCodeList.includes(sub.ID)) : uniqueCodeList.includes(contents.ID)).map(contents => {
            return contents.TYPE === CONTENT_TYPE.CATEGORY ? {
                ...contents,
                CONTENTS: contents.CONTENTS.filter(sub => uniqueCodeList.includes(sub.ID))
            } : contents;
        })
    };
}).filter(category => category.CONTENTS.length);