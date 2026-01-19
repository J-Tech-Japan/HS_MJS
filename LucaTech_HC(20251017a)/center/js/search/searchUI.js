/**
 * 検索UI構築とイベント処理モジュール
 * - ツリービューの構築
 * - 検索結果表示の構築
 * - イベントハンドラーの設定
 */

/**
 * チェックボックス要素を生成
 * @param {Object} node - ノードオブジェクト
 * @param {string} additionalClass - 追加のクラス名
 * @param {boolean} isSelectAll - 「すべて選択」フラグ
 * @returns {string} チェックボックスのHTML
 */
function createCheckboxElement(node, additionalClass = '', isSelectAll = false) {
    // チェック状態の初期値を設定
    const checked = node.checked ? 'checked' : '';
    
    // 「すべて選択」か通常チェックボックスかで属性を分岐
    const id = isSelectAll ? `search-in-all-${node.id}` : `search-in-${node.id}`;
    const className = isSelectAll ? 'search-in-all' : 'search-in';
    const labelText = isSelectAll ? '(すべて選択)' : escapeHtml(node.title);
    
    // 通常チェックボックスの場合のみカウント表示を追加
    const countSpan = isSelectAll ? '' : ` <span class="count" id="count-${node.id}">(0)</span>`;
    
    return `
        <div class="custom-checkbox leaf${additionalClass ? ' ' + additionalClass : ''}">
            <input type="checkbox" ${checked} class="custom-control-input ${className}${additionalClass ? ' ' + additionalClass : ''}" id="${id}">
            <label for="${id}" class="custom-control-label${additionalClass ? ' ' + additionalClass : ''}">${labelText}${countSpan}</label>
        </div>
    `;
}

/**
 * 左メニューのツリービューのノードを構築（再帰関数）
 * @param {Object} node - カタログノード
 * @returns {string} HTMLストリング
 */
function buildTreeViewNode(node) {
    // 入力値検証（早期リターンで簡潔化）
    if (!node || typeof node !== 'object' || !node.id || !node.title) {
        console.warn('buildTreeViewNode: Invalid or missing node properties');
        return '<li></li>';
    }

    // 子ノードの有無を確認
    const hasChildren = Array.isArray(node.childs) && node.childs.length > 0;
    const parts = ['<li>'];
    
    if (hasChildren) {
        // 親ノードの場合：チェックボックス、トグルボタン、子リストを構築
        parts.push(
            createCheckboxElement(node, 'has-childs root', false),
            '<div class="check-toggle active"></div>',
            '<ul class="box-item box-toggle">',
            '<li>', createCheckboxElement(node, '', true), '</li>'
        );
        
        // 再帰的に子ノードを構築
        for (const childNode of node.childs) {
            parts.push(buildTreeViewNode(childNode));
        }
        
        parts.push('</ul>');
    } else {
        // 葉ノードの場合：シンプルなチェックボックスのみ
        parts.push(createCheckboxElement(node));
    }
    
    parts.push('</li>');
    return parts.join('');
}

/**
 * 左メニュー全体を構築（ツリービューのエントリーポイント）
 */
function buildTreeView() {
    const $container = $(".column-left > .column-left-container");
    const searchCatalogue = getSearchCatalogue();
    
    // DOM操作を一括で行うため、fragmentを利用（パフォーマンス最適化）
    const fragment = document.createDocumentFragment();
    
    // 各カタログのツリー構造を構築
    searchCatalogue.forEach(catalogue => {
        const div = document.createElement('div');
        div.className = 'box-check';
        div.innerHTML = `<ul>${buildTreeViewNode(catalogue)}</ul>`;
        fragment.appendChild(div);
    });
    
    // 既存の内容をクリアし、新しいツリーを挿入
    $container.empty().append(fragment);
}

/**
 * ツリービューのイベントハンドラーを設定
 * 注意: この関数は初期化時に一度だけ呼び出されることを想定
 * イベント委譲を使用しているため、動的に生成された要素にも自動的に適用される
 */
function setupTreeViewEventHandlers() {
    // イベント委譲で統一（一度だけ設定）
    $(document).on('click', '.box-check li .check-toggle', onToggleClick);
    $(document).on('click', '.box-check li label.custom-control-label.root', onRootLabelClick);
    $(document).on('change', '.box-check .search-in', onSearchInChange);
    $(document).on('change', '.box-check .search-in-all', onSearchInAllChange);
    $(document).on('click', 'input[type=checkbox].search-in.root', preventRootCheckboxClick);
}

/**
 * ツリーノードの展開/折りたたみトグル処理
 * @param {Event} e - クリックイベント
 */
function onToggleClick(e) {
    const $this = $(this);
    const $siblings = $this.parent().siblings();
    
    // 現在のノードを展開/折りたたみ（アニメーション500ms）
    $this.toggleClass('active').siblings('ul').slideToggle(500);
    
    // 兄弟ノードを全て折りたたむ
    $siblings.children('.check-toggle').removeClass('active');
    $siblings.children('ul').slideUp(500);
}

/**
 * ルートラベルクリック時にトグルボタンをクリック
 * @param {Event} e - クリックイベント
 */
function onRootLabelClick(e) {
    // ラベルクリックで対応するトグルボタンをトリガー
    $(this).closest("li").find(".check-toggle:first").click();
}

/**
 * 検索範囲チェックボックスの変更イベントハンドラー
 * チェック状態を子要素に伝播し、検索結果を更新
 */
function onSearchInChange() {
    const $this = $(this);
    const isChecked = $this.is(":checked");
    const $parentLi = $this.parent().parent();
    
    // 子要素のすべてのチェックボックスを同期
    const $childCheckboxes = $parentLi.find("ul .search-in, ul .search-in-all");
    $childCheckboxes.prop("checked", isChecked);
    
    // 部分選択状態のスタイルをクリア
    const $currentDiv = $this.closest("div");
    const $checkNewElements = $parentLi.find(".check-new");
    $currentDiv.add($checkNewElements).removeClass("check-new");
    
    // 検索結果の表示を更新
    if (typeof displayResult === 'function') {
        displayResult();
    }
    
    // ツリー全体のチェック状態を更新
    if (typeof checkAllInTree === 'function') {
        checkAllInTree(this);
    }
}

/**
 * 「すべて選択」チェックボックスの変更イベントハンドラー
 * 親ノードのチェックボックスと連動
 */
function onSearchInAllChange() {
    // 対応する親チェックボックスのIDを取得
    const id = $(this).attr("id").replace("search-in-all-", "");
    // 親チェックボックスの状態を同期し、changeイベントをトリガー
    $("#search-in-" + id).prop("checked", $(this).is(":checked")).trigger("change");
}

/**
 * ルートチェックボックスの直接クリックを防止
 * （トグルボタン経由でのみ操作可能にする）
 * @param {Event} e - クリックイベント
 */
function preventRootCheckboxClick(e) {
    e.preventDefault();
}

/**
 * 最初のページ（検索カテゴリ選択画面）を構築
 * @returns {string} HTMLストリング
 */
function buildFirstPage() {
    const searchCatalogue = getSearchCatalogue();
    
    return searchCatalogue.map(catalogue => `
        <div class="box-s-1">
            <div class="h1 custom-checkbox">
                <input type="checkbox" 
                       checked 
                       class="custom-control-input parent" 
                       search="${catalogue.id}" 
                       id="search-in-top-${catalogue.id}">
                <label for="search-in-top-${catalogue.id}" 
                       class="custom-control-label">
                    ${escapeHtml(catalogue.title)}
                </label>
            </div>
            <div class="box-item">
                <ul class="box-ul box-ul-1">
                    ${catalogue.childs.map(child => `
                        <li>
                            <div class="custom-checkbox">
                                <input type="checkbox" 
                                       checked 
                                       class="custom-control-input child" 
                                       search="${child.id}" 
                                       id="search-in-top-${child.id}">
                                <label for="search-in-top-${child.id}" 
                                       class="custom-control-label">
                                    ${escapeHtml(child.title)}
                                </label>
                            </div>
                        </li>
                    `).join('')}
                </ul>
            </div>
        </div>
    `).join('');
}

/**
 * 子チェックボックスにクリックイベントを追加
 * 注意: この関数は初期化時に一度だけ呼び出されることを想定
 * イベント委譲を使用しているため、動的に生成された要素にも自動的に適用される
 */
function addHandleEventInFirstPage() {
    // 子チェックボックスのクリックイベント（イベント委譲で一度だけ設定）
    $(document).on("click", "input[type=checkbox].child", function(){
        const $this = $(this);
        const $boxS1 = $this.closest(".box-s-1");
        const $parent = $boxS1.find(".parent");
        const $siblings = $this.closest("ul").find(".child");
        
        // 全ての子チェックボックスの状態を取得
        const checkedStates = $siblings.map(function() { 
            return $(this).is(":checked"); 
        }).get();
        
        // すべてチェック済みか、一部チェック済みかを判定
        const allChecked = checkedStates.every(state => state);
        const someChecked = checkedStates.some(state => state);
        
        // 親チェックボックスの状態を更新
        const $parentDiv = $parent.closest("div");
        $parent.prop("checked", allChecked);
        // 部分選択の場合は視覚的に表示
        $parentDiv.toggleClass("check-new", someChecked && !allChecked);
        
        // ツリービューを再構築
        buildTreeView();
    });
    
    // 親チェックボックスのクリックイベント（イベント委譲で一度だけ設定）
    $(document).on("click", "input[type=checkbox].parent", function(){
        const $this = $(this);
        const isCheck = $this.is(":checked");
        const $boxS1 = $this.closest(".box-s-1");
        
        // 全ての子チェックボックスを親と同じ状態に設定
        $boxS1.find(".child").prop("checked", isCheck);
        
        // 部分選択スタイルをクリア
        $this.closest("div").removeClass("check-new");
        
        // ツリービューを再構築
        buildTreeView();
    });
}

/**
 * 検索UIモジュールを初期化
 * すべてのイベントハンドラーを一度だけ設定する
 * ページ読み込み時に一度だけ呼び出すこと
 */
function initializeSearchUI() {
    // ツリービューのイベントハンドラーを設定
    setupTreeViewEventHandlers();
    
    // 最初のページのイベントハンドラーを設定
    addHandleEventInFirstPage();
}
