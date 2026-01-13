// ページ読み込み完了時の処理
$(function () {

   // localStorageからコンテンツデータが取得できない場合はトップページにリダイレクト
   if (localStorage.getItem('contents') == null) {
      document.location.href = 'index.html';
   } else {
      // localStorageからコンテンツデータを読み込み
      setSearchCatalogue(JSON.parse(localStorage.getItem('contents')));

      // localStorageから検索キーワードを取得して検索ボックスに設定
      var searchKeyword = localStorage.getItem('searchkeyword');
      if (searchKeyword != null) {
         $("#searchkeyword").val(searchKeyword);
      }
   }

   // 検索ボタンクリック時のイベント処理
   $(".btn-search").click(function () {
      // 検索キーワードが入力されており、データ読み込みが完了している場合のみ検索実行
      if ($("#searchkeyword").val() != "") {
         if (isCatalogLoaded()) {
            search();
         }
      }
   });

   // ローディングアイコンを表示（検索実行中の表示）
   $('.box-click-search').html('<div class="loading"><i class="fas fa-spinner fa-spin"></i></div>');

   // 検索UIのイベントハンドラーを初期化（一度だけ実行）
   initializeSearchUI();

   // 初期ページの表示処理
   $('.box-content-s').html(buildFirstPage());

   // ツリービュー構築（絞り込み項目の階層表示）
   buildTreeView();
   
   // 各検索項目のチェック状態を確認
   $(".search-in").each(function () {
      checkAllInTree(this);
   });

   // 全ての絞り込み項目を閉じる処理
   $('.box-check li .check-toggle').each(function () {
      // アクティブクラスを削除し、サブメニューをスライドアップで非表示にする
      $(this).removeClass('active').siblings('ul').slideUp(0);
      $(this).parent().siblings().children('.check-toggle').removeClass('active');
      $(this).parent().siblings().children('ul').slideUp(0);
   });

   // 検索機能の初期化
   initSearch();

   // 検索ボックスでのEnterキー押下時の処理
   $('#searchkeyword').keyup(function (e) {
      // Enterキーが押され、検索キーワードが入力されている場合
      if (e.keyCode == 13) {
         if ($("#searchkeyword").val() != "") {
            // 検索ボタンのクリックイベントを実行
            $(".btn-search").click();
         }
      }
   });
});