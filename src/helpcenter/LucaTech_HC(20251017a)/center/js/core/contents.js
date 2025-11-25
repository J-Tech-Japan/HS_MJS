	var baseFolderWebHelp = "../contents";

CONTENTS_ALL = [
    {
        ID: "CMN",
        TYPE: "category",
        TITLE: "共通",
        STYLE: "0",
        CONTENTS: [
            {
                ID: "CMN_OPE",
                TITLE: "『LucaTech GX』について",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/CMN_OPE/"
            },
            {
                ID: "CMN_JNT",
                TITLE: "導入処理",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/CMN_JNT/",
            },
            {
                ID: "CMN_MAS",
                TITLE: "マスター登録処理",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/CMN_MAS/",
            },
            {
                ID: "CMN_CMP",
                TITLE: "会社データ管理",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/CMN_CMP/"
            },
            {
                ID: "CMN_UTL",
                TITLE: "システム管理",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/CMN_UTL/"
            },
            {
                ID: "CMN_DAT",
                TITLE: "データ交換",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/CMN_DAT/"
            },
        ]
    },
    {
        ID: "MAS",
        TYPE: "category",
        TITLE: "財務会計システム",
        STYLE: "1",
        CONTENTS: [
            {
                ID: "MAS_OPE",
                TITLE: "財務会計システムについて",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/MAS_OPE/",
            },
            {
                ID: "MAS_sub01",
                TITLE: "導入処理",
                TYPE: "category",
                CONTENTS: [
                    {
                        ID: "MAS_JNT",
                        TITLE: "基本設定",
                        TYPE: "contents",
                        PATH: baseFolderWebHelp + "/MAS_JNT/"
                    },
                    {
                        ID: "MAS_DEP",
                        TITLE: "部門設定",
                        TYPE: "contents",
                        PATH: baseFolderWebHelp + "/MAS_DEP/"
                    },
                    {
                        ID: "MAS_BAB",
                        TITLE: "残高・予算設定",
                        TYPE: "contents",
                        PATH: baseFolderWebHelp + "/MAS_BAB/"
                    },
                ]
            },
            {
                ID: "MAS_DAY",
                TITLE: "日常処理",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/MAS_DAY/"
            },
            {
                ID: "MAS_SOA",
                TITLE: "決算処理",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/MAS_SOA/"
            },
            {
                ID: "MAS_EAL",
                TITLE: "電子帳簿",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/MAS_EAL/"
            },
        ]
    },
    {
        ID: "DEP",
        TYPE: "category",
        TITLE: "固定資産システム",
        STYLE: "2",
        CONTENTS: [
            {
                ID: "DEP_JNT",
                TITLE: "導入処理",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/DEP_JNT/"
            },
            {
                ID: "DEP_SIS",
                TITLE: "資産管理",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/DEP_SIS/"
            },
        ]
    },
    {
        ID: "FRT",
        TYPE: "category",
        TITLE: "ワークフローシステム",
        STYLE: "5",
        CONTENTS: [
            {
                ID: "WFL_JNT",
                TITLE: "導入処理",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/WFL_JNT/"
            },
            {
                ID: "WFL_FOM",
                TITLE: "フォーム設定",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/WFL_FOM/"
            },
            {
                ID: "WFL_DAY",
                TITLE: "日常処理",
                TYPE: "contents",
                PATH: baseFolderWebHelp + "/WFL_DAY/"
            },
        ]
    }
];


var cookie = document.cookie.split('; ').reduce((prev, current) => {
    const [name, ...value] = current.split('=');
    prev[name] = value.join('=');
    return prev;
}, {});

var gi_license = "gi_license" in cookie ? cookie.gi_license.split(",") : [];

CONTENTS = CONTENTS_ALL.map(category => {
    return {
        ...category,
        CONTENTS: category.CONTENTS.filter(contents => contents.TYPE === "category" ? contents.CONTENTS.some(sub => gi_license.includes(sub.ID)) : gi_license.includes(contents.ID)).map(contents => {
            return contents.TYPE === "category" ? {
                ...contents,
                CONTENTS: contents.CONTENTS.filter(sub => gi_license.includes(sub.ID))
            } : contents;
        })
    };
}).filter(category => category.CONTENTS.length);