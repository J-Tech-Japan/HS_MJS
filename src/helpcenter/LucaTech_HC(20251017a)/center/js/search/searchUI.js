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
function buildTreeViewNode(node) {
    // 入力値検証（早期リターンで簡潔化）
    if (!node || typeof node !== 'object' || !node.id || !node.title) {
        console.warn('buildTreeViewNode: Invalid or missing node properties');
        return '<li></li>';
    }

    const hasChildren = Array.isArray(node.childs) && node.childs.length > 0;
    const parts = ['<li>'];
    
    if (hasChildren) {
        parts.push(
            createCheckboxElement(node, 'has-childs root', false),
            '<div class="check-toggle active"></div>',
            '<ul class="box-item box-toggle">',
            '<li>', createCheckboxElement(node, '', true), '</li>'
        );
        
        for (const childNode of node.childs) {
            parts.push(buildTreeViewNode(childNode));
        }
        
        parts.push('</ul>');
    } else {
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
    // DOM操作を一括で行うため、fragmentを利用
    const fragment = document.createDocumentFragment();
    searchCatalogue.forEach(catalogue => {
        const div = document.createElement('div');
        div.className = 'box-check';
        div.innerHTML = `<ul>${buildTreeViewNode(catalogue)}</ul>`;
        fragment.appendChild(div);
    });
    $container.empty().append(fragment);
    setupTreeViewEventHandlers();
}

/**
 * ツリービューのイベントハンドラーを設定
 */
function setupTreeViewEventHandlers() {
    const $doc = $(document);
    // 重複登録を防ぐため既存のイベントを解除
    $doc.off('click', '.box-check li .check-toggle');
    $doc.off('click', '.box-check li label.custom-control-label.root');
    $doc.off('change', '.box-check .search-in');
    $doc.off('change', '.box-check .search-in-all');
    $doc.off('click', 'input[type=checkbox].search-in.root');
    
    // イベント委譲を統一
    $doc.on('click', '.box-check li .check-toggle', onToggleClick);
    $doc.on('click', '.box-check li label.custom-control-label.root', onRootLabelClick);
    $doc.on('change', '.box-check .search-in', onSearchInChange);
    $doc.on('change', '.box-check .search-in-all', onSearchInAllChange);
    $doc.on('click', 'input[type=checkbox].search-in.root', preventRootCheckboxClick);
}

function onToggleClick(e) {
    const $this = $(this);
    const $siblings = $this.parent().siblings();
    $this.toggleClass('active').siblings('ul').slideToggle(500);
    $siblings.children('.check-toggle').removeClass('active');
    $siblings.children('ul').slideUp(500);
}

function onRootLabelClick(e) {
    $(this).closest("li").find(".check-toggle:first").click();
}

function onSearchInChange() {
    const $this = $(this);
    const check = $this.is(":checked");
    const $parent = $this.parent().parent();
    const $ul = $parent.find("ul");
    const $checkboxes = $ul.find(".search-in, .search-in-all");
    $checkboxes.prop("checked", check);
    const $div = $this.closest("div");
    $div.add($parent.find(".check-new")).removeClass("check-new");
    if (typeof displayResult === 'function') {
        displayResult();
    }
    checkAllInTree(this);
}

function onSearchInAllChange() {
    const id = $(this).attr("id").replace("search-in-all-", "");
    $("#search-in-" + id).prop("checked", $(this).is(":checked")).trigger("change");
}

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
 */
function addHandleEventInFirstPage() {
    const $doc = $(document);
    
    // 重複登録を防ぐため既存のイベントを解除
    $doc.off("click", "input[type=checkbox].child");
    $doc.off("click", "input[type=checkbox].parent");
    
    $doc.on("click", "input[type=checkbox].child", function(){
        const $this = $(this);
        const $boxS1 = $this.closest(".box-s-1");
        const $parent = $boxS1.find(".parent");
        const $siblings = $this.closest("ul").find(".child");
        const checkedStates = $siblings.map(function() { 
            return $(this).is(":checked"); 
        }).get();
        const allChecked = checkedStates.every(state => state);
        const someChecked = checkedStates.some(state => state);
        const $parentDiv = $parent.closest("div");
        $parent.prop("checked", allChecked);
        $parentDiv.toggleClass("check-new", someChecked && !allChecked);
        buildTreeView();
    });
    $doc.on("click", "input[type=checkbox].parent", function(){
        const $this = $(this);
        const isCheck = $this.is(":checked");
        const $boxS1 = $this.closest(".box-s-1");
        $boxS1.find(".child").prop("checked", isCheck);
        $this.closest("div").removeClass("check-new");
        buildTreeView();
    });
}


