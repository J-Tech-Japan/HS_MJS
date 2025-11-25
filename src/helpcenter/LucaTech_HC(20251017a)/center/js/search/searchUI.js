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
    const checked = node.checked ? 'checked' : '';
    const id = isSelectAll ? `search-in-all-${node.id}` : `search-in-${node.id}`;
    const className = isSelectAll ? 'search-in-all' : 'search-in';
    const labelText = isSelectAll ? '(すべて選択)' : escapeHtml(node.title);
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
function buildTreeviewNode(node) {
    // 入力値検証
    if (!node || typeof node !== 'object') {
        console.warn('buildTreeviewNode: Invalid node object');
        return '<li></li>';
    }
    
    if (!node.id || !node.title) {
        console.warn('buildTreeviewNode: Missing required node properties (id, title)');
        return '<li></li>';
    }
    
    const hasChildren = node.childs && Array.isArray(node.childs) && node.childs.length > 0;
    
    let treeview = '<li>';
    
    if (hasChildren) {
        // 親ノード（子要素あり）の場合
        treeview += createCheckboxElement(node, 'has-childs root', false);
        treeview += '<div class="check-toggle active"></div>';
        treeview += '<ul class="box-item box-toggle">';
        
        // 「すべて選択」オプション
        treeview += '<li>';
        treeview += createCheckboxElement(node, '', true);
        treeview += '</li>';
        
        // 子ノードを再帰的に構築
        for (const childNode of node.childs) {
            treeview += buildTreeviewNode(childNode);
        }
        
        treeview += '</ul>';
    } else {
        // リーフノード（子要素なし）の場合
        treeview += createCheckboxElement(node);
    }
    
    treeview += '</li>';
    return treeview;
}

/**
 * 左メニュー全体を構築（ツリービューのエントリーポイント）
 */
function buildTreeview() {
    const $container = $(".column-left > .column-left-container");
    const searchCatalogue = getSearchCatalogue();
    
    const html = searchCatalogue
        .map(catalogue => `<div class='box-check'><ul>${buildTreeviewNode(catalogue)}</ul></div>`)
        .join('');
    
    $container.html(html);
    setupTreeviewEventHandlers();
}

/**
 * ツリービューのイベントハンドラーを設定
 */
function setupTreeviewEventHandlers() {
    const $boxCheck = $('.box-check');
    
    // トグルクリック
    $boxCheck.on('click', 'li .check-toggle', function (e) {
        const $this = $(this);
        const $siblings = $this.parent().siblings();
        
        $this.toggleClass('active').siblings('ul').slideToggle(500);
        $siblings.children('.check-toggle').removeClass('active');
        $siblings.children('ul').slideUp(500);
    });

    $boxCheck.on('click', 'li label.custom-control-label.root', function (e) {
        $(this).closest("li").find(".check-toggle:first").click();
    });

    // search-in変更イベント
    $boxCheck.on('change', '.search-in', function() {
        const $this = $(this);
        const check = $this.is(":checked");
        const $parent = $this.parent().parent();
        
        $parent.find("ul .search-in, ul .search-in-all").prop("checked", check);
        $this.closest("div").add($parent.find(".check-new")).removeClass("check-new");
        
        if (typeof displayResult === 'function') {
            displayResult();
        }
        
        checkAllInTree(this);
    });

    // search-in-all変更イベント
    $boxCheck.on('change', '.search-in-all', function() {
        const id = $(this).attr("id").replace("search-in-all-", "");
        $("#search-in-" + id).prop("checked", $(this).is(":checked")).trigger("change");
    });

    // rootチェックボックスのクリック防止
    $(document).on("click", "input[type=checkbox].search-in.root", function(e){
        e.preventDefault();
    });
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
 */
function addHandleEventInFirstPage() {
    const $container = $("body");
    
    $container.on("click", "input[type=checkbox].child", function(){
        const $this = $(this);
        const $parent = $this.closest(".box-s-1").find(".parent");
        const $siblings = $this.closest("ul").find(".child");
        
        const checkedStates = $siblings.map(function() { 
            return $(this).is(":checked"); 
        }).get();
        
        const allChecked = checkedStates.every(state => state);
        const someChecked = checkedStates.some(state => state);
        
        const $parentDiv = $parent.closest("div");
        
        $parent.prop("checked", allChecked);
        $parentDiv.toggleClass("check-new", someChecked && !allChecked);

        buildTreeview();
    });

    $container.on("click", "input[type=checkbox].parent", function(){
        const $this = $(this);
        const isCheck = $this.is(":checked");
        
        $this.closest(".box-s-1").find(".child").prop("checked", isCheck);
        $this.closest("div").removeClass("check-new");
        buildTreeview();
    });
}


