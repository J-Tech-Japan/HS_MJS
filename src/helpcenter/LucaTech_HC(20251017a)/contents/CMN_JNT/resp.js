var resp;
if (typeof respClass === 'undefined') {
	var respClass = function () {
		this.init();
	};
}


respClass.prototype = {
	init: function () {
		this.ww = window.innerWidth;
		this._frame = '';
		this.reloaded = false;
		this._type = '';
		this._setToc = false;
		this.userAgent = window.navigator.userAgent.toLowerCase();
		this.is_iPhone = false;
		this.is_iPad = false;
		this.is_android = false;
		this.is_androidTab = false;

		if (this.userAgent.match(/iphone/i)) { this.is_iPhone = true; }
		if (this.userAgent.match(/ipad/i)) { this.is_iPad = true; }
		if (this.userAgent.match(/android/i) && this.userAgent.match(/mobile/i)) { this.is_android = true; }
		if (this.userAgent.match(/android/i) && !this.userAgent.match(/mobile/i)) { this.is_androidTab = true; }

	},
	ready: function () {
		//console.log('resp ready');
		resp.resizeHandler();
		resp.initialize();
		resp.setNavi();
		resp.set_iframeLoad();

		// breadcrumb機能が利用可能な場合のみ実行
		if (typeof window.initBreadcrumbIfAvailable === 'function') {
			window.initBreadcrumbIfAvailable(this);
		}

		// 初回読み込み時のアンカー処理（少し遅延させてiframe準備を待つ）
		setTimeout(function () {
			resp.handleAnchorNavigation();
		}, 1000);
	},
	initialize: function () {
		//console.log('initialize');
		if (resp._type !== 'desktop') {
			$('.nav').find('.toc').removeClass('active');
		}
	},
	loaded: function () {
		setTimeout(function () {
			resp.initialize();
		}, 100);
	},
	setNavi: function () {
		$('.nav').find('.toc').on('click', function () {
			if ($('body').hasClass('media-landscape') || $('body').hasClass('media-mobile')) {
				resp.closeSearch();
				resp.toggleToc();
			}
		});

		$('.mobilespecialfunctions').find('.menubutton').on('click', function () {
			if ($('body').hasClass('media-landscape') || $('body').hasClass('media-mobile')) {
				resp.closeSearch();
				resp.toggleToc();
			}
		});

		$('.nav').find('.fts').on('click', function () {
			if ($('body').hasClass('media-landscape') || $('body').hasClass('media-mobile')) {
				resp.closeToc();
				resp.toggleSearch();
			}
		});

		$('.mobilespecialfunctions').find('.fts').on('click', function () {
			if ($('body').hasClass('media-landscape') || $('body').hasClass('media-mobile')) {
				resp.closeToc();
				resp.openSearch();
			}
		});

		$('.mobile_back').on('click', function () {
			if ($('body').hasClass('media-landscape') || $('body').hasClass('media-mobile')) {
				resp.closeSearch();
			}
		});

		$('.wSearchResultItemsBlock').on('click', '.nolink', function () {
			if ($('body').hasClass('media-landscape') || $('body').hasClass('media-mobile')) {
				resp.closeSearch();
				resp.closeToc();
			}
		});

		// ページ内遷移時のアンカー処理
		window.addEventListener('hashchange', function () {
			setTimeout(function () {
				resp.handleAnchorNavigation();
			}, 300);
		});

		// $('body').append('<div class="toc-close-btn">×</div>');
		// $('.toc-close-btn').on('click', function(){
		// 	resp.closeToc();
		// });
	},
	setTocNavi: function () {
		if (!resp._setToc) {
			resp._setToc = true;

			$('.toc-holder .toc').on('click', 'a', function () {
				if ($('body').hasClass('media-landscape') || $('body').hasClass('media-mobile')) {
					if (!$(this).parent().hasClass('book')) {
						resp.closeSearch();
						resp.closeToc();
					}
				}
			});
		}
	},
	set_iframeLoad: function () {
		$('.topic').on('load', function () {
			if (this.inited) {
				// breadcrumb機能が利用可能な場合のみ実行
				if (typeof window.showBreadcrumbIfAvailable === 'function') {
					window.showBreadcrumbIfAvailable(true);
				}
			}
			resp.frameLoadedResp();

			// アンカーリンク対応：iframe読み込み完了後にアンカー位置に移動
			resp.handleAnchorNavigation();

			this.inited = true;
		});
	},

	// アンカーナビゲーション処理
	handleAnchorNavigation: function () {
		var hash = window.location.hash;
		if (hash && hash.length > 1) {
			// URLハッシュからアンカー部分を抽出（例：#t=KAC22150.html#anchor -> #anchor）
			var anchorMatch = hash.match(/#t=[^#]+#(.+)$/);
			if (anchorMatch && anchorMatch[1]) {
				var anchorId = anchorMatch[1];
				resp.scrollToAnchorInIframe(anchorId);
			}
		}
	},

	// iframe内の指定アンカーにスクロール
	scrollToAnchorInIframe: function (anchorId, retryCount) {
		retryCount = retryCount || 0;
		var maxRetries = 10; // 最大リトライ回数
		var retryDelay = 200; // リトライ間隔（ミリ秒）

		try {
			var iframe = document.querySelector('.topic');
			if (iframe && iframe.contentWindow && iframe.contentWindow.document) {
				var iframeDoc = iframe.contentWindow.document;
				var targetElement = iframeDoc.getElementById(anchorId) ||
					iframeDoc.querySelector('[name="' + anchorId + '"]');

				if (targetElement) {
					// 要素が見つかった場合、スクロール実行
					var elementTop = targetElement.offsetTop;
					iframe.contentWindow.scrollTo(0, elementTop);
					console.log('アンカーに移動しました:', anchorId);
					return true;
				} else if (retryCount < maxRetries) {
					// 要素が見つからない場合、少し待ってリトライ
					setTimeout(function () {
						resp.scrollToAnchorInIframe(anchorId, retryCount + 1);
					}, retryDelay);
					return false;
				} else {
					console.warn('アンカー要素が見つかりませんでした:', anchorId);
					return false;
				}
			} else if (retryCount < maxRetries) {
				// iframe が準備できていない場合、リトライ
				setTimeout(function () {
					resp.scrollToAnchorInIframe(anchorId, retryCount + 1);
				}, retryDelay);
				return false;
			}
		} catch (error) {
			console.error('アンカー移動エラー:', error);
			if (retryCount < maxRetries) {
				setTimeout(function () {
					resp.scrollToAnchorInIframe(anchorId, retryCount + 1);
				}, retryDelay);
			}
		}
		return false;
	},
	openToc: function () {
		$('.nav').find('.toc').addClass('on');
		$('.toc-holder').addClass('on');
		// $('.toc-close-btn').addClass('on');
		$('.mobilespecialfunctions').find('.menubutton').addClass('on');
		resp.setTocNavi();//初期化時だと付与されないため
	},
	closeToc: function () {
		$('.nav').find('.toc').removeClass('on');
		$('.toc-holder').removeClass('on');
		// $('.toc-close-btn').removeClass('on');
		$('.mobilespecialfunctions').find('.menubutton').removeClass('on');
	},
	toggleToc: function () {
		if ($('.nav').find('.toc').hasClass('on')) {
			resp.closeToc();
			$('.mobilespecialfunctions').find('.menubutton').removeClass('on');
		} else {
			resp.openToc();
			$('.mobilespecialfunctions').find('.menubutton').addClass('on');
		}
	},
	openSearch: function () {
		$('.nav').find('.fts').addClass('on');
		$('div.searchbar').addClass('on');
	},
	closeSearch: function () {
		setTimeout(function () {
			$('.nav').find('.fts').removeClass('on');
			$('div.searchbar').removeClass('on').removeClass('sidebar-opened').removeClass('searchpage-mode').removeClass('layout-visible');
			$('div.searchresults').removeClass('sidebar-opened').removeClass('layout-visible');
			$('div.topic.main').removeClass('sidebar-opened');
			$('div.mobilespecialfunctions').removeClass('sidebar-opened').removeClass('searchpage-mode');
			$('div.functionbar').removeClass('sidebar-opened');
			$('div.filter-holder').removeClass('sidebar-opened');
			$('div.search-input').addClass('rh-hide');
		}, 500);

		//searchresults left-pane search-sidebar layout-visible

		//filter-holder left-pane
		//filter-holder left-pane sidebar-opened

		//
		//functionbar

		//mobilespecialfunctions sidebar-opened searchpage-mode
		//mobilespecialfunctions

		//topic main sidebar-opened
		//topic main

		//searchresults left-pane search-sidebar sidebar-opened layout-visible
		//searchresults left-pane search-sidebar
	},
	toggleSearch: function () {
		if ($('.nav').find('.fts').hasClass('on')) {
			resp.closeSearch();
		} else {
			resp.openSearch();
		}
	},
	frameLoadedResp: function () {
		resp.resizeHandler();
	},
	checkDevice: function () {
		var _device = '';



		return _device;
	},
	resizeHandler: function () {
		resp.ww = window.innerWidth;
		var _body = $('body');
		_body.removeClass('media-desktop');
		_body.removeClass('media-landscape');
		_body.removeClass('media-mobile');
		_body.removeClass('f_small');
		_body.removeClass('f_medium');
		_body.removeClass('f_large');
		var _frame = document.getElementsByClassName("topic");
		var _frameBody = _frame[1].contentWindow.document.body;
		_frameBody.classList.remove('media-desktop');
		_frameBody.classList.remove('media-landscape');
		_frameBody.classList.remove('media-mobile');

		if (resp.ww > 1024 && !resp.is_androidTab && !resp.is_iPad && !resp.is_iPhone && !resp.is_android) {
			resp._type = 'desktop';
			_body.addClass('media-desktop');
			_frameBody.classList.add('media-desktop');
			fc.fontChangeIndex();
			// fc.fontChangeIframe();
		} else if ((resp.ww > 700 || resp.is_androidTab || resp.is_iPad) && !resp.is_iPhone && !resp.is_android) {
			resp._type = 'landscape';
			_body.addClass('media-landscape');
			_frameBody.classList.add('media-landscape');
			_body.addClass('f_medium');
		} else if (resp.ww < 700 || resp.is_iPhone || resp.is_android) {
			resp._type = 'mobile';
			_body.addClass('media-mobile');
			_frameBody.classList.add('media-mobile');
			_body.addClass('f_medium');
		}

		resp.initialize();

		//console.log(resp._type);
		//console.log(resp.ww);
	},
	openAllBookBefore: function () {
		$(".book").each(function () {
			if (!$(this).hasClass("expanded")) {
				var aTag = $($(this).find("a")[0]);
				aTag.attr("hreftemp", aTag.attr('href'));
				aTag.attr('href', "#");
				$(this).trigger("click");
			}
		});
	},
	openAllBookEnd: function () {
		$(".book").each(function () {
			var aTag = $($(this).find("a")[0]);
			aTag.attr("href", aTag.attr('hreftemp'));
			aTag.removeAttr('hreftemp');
		});
	}
};

resp = new respClass();

$(function () {
	resp.ready();

	//setTimeout(function(){
	//	resp.openAllBookBefore();// level1
	//	setTimeout(function(){
	//		resp.openAllBookBefore();// level2
	//		setTimeout(function(){
	//			resp.openAllBookBefore();// level3
	//			setTimeout(function(){
	//				resp.openAllBookBefore();// level4
	//				resp.openAllBookEnd();
	//			},500);
	//		},500);
	//	},500);
	//},500);

	$('iframe.topic').on('load', function () {
		resp.resizeHandler();

		var timer = setTimeout(function () {
			resp.resizeHandler();

			timer = setTimeout(function () {
				resp.resizeHandler();

				timer = setTimeout(function () {
					resp.resizeHandler();
				}, 500);
			}, 500);
		}, 500);
	});

	$(window).on('load resize', function () {
		resp.resizeHandler();

		var timer = setTimeout(function () {
			resp.resizeHandler();

			timer = setTimeout(function () {
				resp.resizeHandler();

				timer = setTimeout(function () {
					resp.resizeHandler();
				}, 500);
			}, 500);
		}, 500);
	});
});