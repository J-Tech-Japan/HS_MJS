// 共通のインデックスページコンポーネント (index.html、menu.html、purpose.html、sys_top.html で使用)
function createIndexPageApp(indexType) {
  const { createApp } = Vue;

  return createApp({
    data() {
      return {
        contentsData: null,
        menuData: null,
        categoryTitle: "",
        subCategoryTitle: "",
        searchKeyword: "",
        isChecked: true,
        indexType: indexType, // "index"、"menu"、"purpose"、または "sys_top"
        introLinks: [
          {
            TITLE: "スタートガイド",
            HREF: "../contents/startguide.pdf",
            TARGET: "_blank"
          },
          {
            TITLE: "導入の流れ",
            HREF: "../contents/CMN_JNT/index.html#t=CMN11003.html"
          }
        ]
      };
    },
    created() {
      // index.html の場合は contentid チェックをスキップ
      if (this.indexType === "index") {
        this.contentsData = CONTENTS;
        // search.htmlからの戻りの場合、検索キーワードを復元
        if (document.referrer.indexOf("search.html") > -1) {
          const searchkeyword = localStorage.getItem("searchkeyword");
          if (searchkeyword) {
            this.searchKeyword = searchkeyword;
          }
        }
        return;
      }

      if (!this.contentid) {
        location.replace("index.html");
        return;
      }
      this.findContentsData();
      if (this.indexType !== "sys_top") {
        this.loadIndexData();
      } else {
        document.title = `${this.contentsTitle} | LucaTech GX ヘルプセンター`;
      }
    },
    computed: {
      contentid() {
        const searchParams = new URLSearchParams(window.location.search);
        return searchParams.get('contentid');
      },
      contentidPrefix() {
        return this.contentid?.toLowerCase().slice(0, 3) ?? "";
      },
      contentsTitle() {
        return this.contentsData?.TITLE ?? "";
      },
      title() {
        return this.menuData?.TITLE ?? "";
      },
      category() {
        return this.menuData?.CATEGORY ?? [];
      },
      contents() {
        return this.menuData?.CONTENTS ?? {};
      },
      path() {
        return this.contentsData?.PATH ?? "";
      },
      breadcrumbData() {
        return {
          indexType: this.indexType,
          contentid: this.contentid,
          categoryTitle: this.categoryTitle,
          subCategoryTitle: this.subCategoryTitle,
          contentsTitle: this.contentsTitle
        };
      },
      breadcrumbLabel() {
        if (this.indexType === "menu") return "メニューから探す";
        if (this.indexType === "purpose") return "目的から探す";
        return ""; // sys_top の場合は空文字
      }
    },
    methods: {
      findContentsData() {
        for (const category of CONTENTS) {
          for (const item of category.CONTENTS) {
            if (item.TYPE === "category") {
              const found = item.CONTENTS.find(item2 => item2.ID === this.contentid);
              if (found) {
                this.contentsData = found;
                this.categoryTitle = category.TITLE;
                this.subCategoryTitle = item.TITLE;
                return;
              }
            } else if (item.TYPE === "contents" && item.ID === this.contentid) {
              this.contentsData = item;
              this.categoryTitle = category.TITLE;
              return;
            }
          }
        }
      },
      loadIndexData() {
        fetch(`./data/${this.indexType}/${this.contentid}_${this.indexType}.json`)
          .then(response => response.json())
          .then(data => {
            this.menuData = data;
            document.title = `${this.title} | ${this.contentsTitle} | LucaTech GX ヘルプセンター`;
          })
          .catch(error => {
            console.error(`${this.breadcrumbLabel}データの読み込みエラー:`, error);
          });
      },
      getUrl(id) {
        const parts = id.split("#");
        parts[0] = parts[0] + '.html';
        return this.path + 'index.html#t=' + parts.join("#");
      },
      setBreadCrumbLS(event, subContentsTitle) {
        event.preventDefault();
        const path = event.currentTarget.href;
        localStorage.setItem("breadcrumb", JSON.stringify({
          ...this.breadcrumbData,
          path,
          subContentsTitle
        }));
        window.location.href = path;
      },
      onSearch() {
        if (!this.searchKeyword) return;
        
        const jsonModel = CONTENTS
          .map(content => buildSearchTreeModel(content))
          .filter(model => model != null);

        // index.html 以外の場合のみチェックボックスの状態を保存
        if (this.indexType !== "index") {
          const checkId = Array.from(document.querySelectorAll(".search-in:checked"))
            .map(checkbox => checkbox.id);
          localStorage.setItem("checkId", JSON.stringify(checkId));
        }

        localStorage.setItem("contents", JSON.stringify(jsonModel));
        localStorage.setItem("searchkeyword", this.searchKeyword);
        document.location.href = "search.html";
      }
    }
  });
}
