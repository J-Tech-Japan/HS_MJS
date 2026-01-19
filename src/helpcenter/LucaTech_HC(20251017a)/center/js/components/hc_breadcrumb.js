/**
 * Breadcrumb Navigation Module
 * パンくずリスト機能を提供するモジュール
 */

var BreadcrumbManager;
if (typeof BreadcrumbManager === 'undefined') {
	BreadcrumbManager = function () {
		this.init();
	};
}

BreadcrumbManager.prototype = {
	init: function () {
		this.breadcrumbData = {};
		this.isEnabled = true;
	},

	/**
	 * breadcrumbデータを設定
	 * @param {Object} data - breadcrumbデータ
	 */
	setBreadcrumbData: function (data) {
		this.breadcrumbData = data || {};
	},

	/**
	 * breadcrumbデータを取得
	 * @returns {Object} breadcrumbデータ
	 */
	getBreadcrumbData: function () {
		return this.breadcrumbData;
	},

	/**
	 * localStorageからbreadcrumbデータを読み込み、表示処理を実行
	 * @param {Object} respInstance - respClassのインスタンス
	 */
	initializeBreadcrumb: function (respInstance) {
		if (!this.isEnabled) return;

		try {
			const lsBreadcrumb = localStorage.getItem("breadcrumb");
			this.breadcrumbData = lsBreadcrumb ? JSON.parse(lsBreadcrumb) : {};
			const path = location.pathname.slice(location.pathname.indexOf('/contents/'));
			
			if (this.breadcrumbData.path && this.breadcrumbData.path.match(path)) {
				this.showBreadcrumb(false);
				
				// 検索キーワードが存在する場合、検索ボックスに設定
				if (this.breadcrumbData.searchKeyword) {
					this.setSearchKeyword(this.breadcrumbData.searchKeyword);
				}
			} else {
				this.showBreadcrumb(true);
			}
			
			console.log("set breadcrumb", this.breadcrumbData);
			localStorage.removeItem("breadcrumb");
		} catch (e) {
			console.log("Breadcrumb initialization error:", e);
		}
		
		$('body').addClass('has-breadcrumb');
	},

	/**
	 * breadcrumbを表示
	 * @param {boolean} init - 初期化モードかどうか
	 */
	showBreadcrumb: function (init) {
		if (!this.isEnabled) return;

		init = init || false;
		
		// header要素が存在しない場合は作成
		if ($('body').find('header.menu-breadcrumb').length === 0) {
			$('body').prepend('<header class="menu-breadcrumb"></header>');
		}

		if (!init && Object.keys(this.breadcrumbData).length > 0) {
			if (this.breadcrumbData.indexType === "search") {
				this._renderSearchBreadcrumb();
			} else {
				this._renderNormalBreadcrumb();
			}
		} else {
			this._renderDefaultBreadcrumb();
		}
	},

	/**
	 * 検索結果用のbreadcrumbを描画
	 * @private
	 */
	_renderSearchBreadcrumb: function () {
		const html = '<ul class="search-result">' +
			'<li><a href="../../center/index.html" target="_self">LucaTech GX ヘルプTOP</a></li>' +
			'<li>検索結果</li>' +
			'<li>' + this._escapeHtml(this.breadcrumbData.categoryTitle) + '</li>' +
			(this.breadcrumbData.subCategoryTitle ? 
				('<li>' + this._escapeHtml(this.breadcrumbData.subCategoryTitle) + '</li>') : '') +
			'<li>' + this._escapeHtml(this.breadcrumbData.contentsTitle) + '</li>' +
			'</ul>';
		
		$('header.menu-breadcrumb').html(html);
	},

	/**
	 * 通常のbreadcrumbを描画
	 * @private
	 */
	_renderNormalBreadcrumb: function () {
		const purposeLink = this.breadcrumbData.indexType === "purpose" ?
			'<a href="../../center/purpose.html?contentid=' + this._escapeHtml(this.breadcrumbData.contentid) + '" target="_self">目的から探す</a>' :
			'<a href="../../center/menu.html?contentid=' + this._escapeHtml(this.breadcrumbData.contentid) + '" target="_self">メニューから探す</a>';

		const html = '<ul>' +
			'<li><a href="../../center/index.html" target="_self">LucaTech GX ヘルプTOP</a></li>' +
			'<li>' + this._escapeHtml(this.breadcrumbData.categoryTitle) + '</li>' +
			(this.breadcrumbData.subCategoryTitle ? 
				('<li>' + this._escapeHtml(this.breadcrumbData.subCategoryTitle) + '</li>') : '') +
			'<li><a href="../../center/sys_top.html?contentid=' + this._escapeHtml(this.breadcrumbData.contentid) + '" target="_self">' + 
				this._escapeHtml(this.breadcrumbData.contentsTitle) + 'TOP</a></li>' +
			'<li>' + purposeLink + '</li>' +
			(this.breadcrumbData.subContentsTitle ? 
				('<li>' + this._escapeHtml(this.breadcrumbData.subContentsTitle) + '</li>') : '') +
			'</ul>';
		
		$('header.menu-breadcrumb').html(html);
	},

	/**
	 * デフォルトのbreadcrumbを描画
	 * @private
	 */
	_renderDefaultBreadcrumb: function () {
		const html = '<ul><li><a href="../../center/index.html" target="_self">LucaTech GX ヘルプTOP</a></li></ul>';
		$('header.menu-breadcrumb').html(html);
	},

	/**
	 * HTMLエスケープ処理
	 * @param {string} str - エスケープする文字列
	 * @returns {string} エスケープされた文字列
	 * @private
	 */
	_escapeHtml: function (str) {
		if (!str) return '';
		return str.toString()
			.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;')
			.replace(/'/g, '&#39;');
	},

	/**
	 * breadcrumb機能を無効化
	 */
	disable: function () {
		this.isEnabled = false;
	},

	/**
	 * breadcrumb機能を有効化
	 */
	enable: function () {
		this.isEnabled = true;
	},

	/**
	 * breadcrumb機能が有効かどうかを確認
	 * @returns {boolean} 有効かどうか
	 */
	isActive: function () {
		return this.isEnabled;
	},

	/**
	 * 検索キーワードを検索ボックスに設定
	 * @param {string} keyword - 設定するキーワード
	 */
	setSearchKeyword: function (keyword) {
		if (!keyword) return;
		
		// 検索ボックスに値を設定（複数のセレクタを試行）
		const $searchField = $('.wSearchField');
		if ($searchField.length > 0) {
			$searchField.val(keyword);
			// keyupイベントをトリガーして検索を実行
			$searchField.trigger('keyup');
		}
	}
};

// グローバルインスタンスを作成
if (typeof window !== 'undefined') {
	window.breadcrumbManager = new BreadcrumbManager();
	
	// resp.jsとの連携用のヘルパー関数
	window.initBreadcrumbIfAvailable = function(respInstance) {
		if (window.breadcrumbManager && typeof window.breadcrumbManager.initializeBreadcrumb === 'function') {
			window.breadcrumbManager.initializeBreadcrumb(respInstance);
		}
	};

	window.showBreadcrumbIfAvailable = function(init) {
		if (window.breadcrumbManager && typeof window.breadcrumbManager.showBreadcrumb === 'function') {
			window.breadcrumbManager.showBreadcrumb(init);
		}
	};
}
