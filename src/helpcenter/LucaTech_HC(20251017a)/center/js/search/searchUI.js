/**
 * 検索UI構築とイベント処理モジュール (Vue.js対応版)
 * jQuery依存を除去し、Vue.jsのリアクティブシステムと互換性を持たせています
 * 
 * - ツリービューの構築
 * - 検索結果表示の構築
 * - イベントハンドラーの設定
 */

/**
 * UI状態管理オブジェクト
 * Vue.jsのリアクティブシステムで使用可能
 */
const uiState = {
    treeviewBuilt: false,
    firstPageBuilt: false,
    eventHandlersAttached: false
};

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
    // 入力値検証（早期リターンで簡潔化）
    if (!node || typeof node !== 'object' || !node.id || !node.title) {
        console.warn('buildTreeviewNode: Invalid or missing node properties');
        return '<li></li>';
    }

    const hasChildren = Array.isArray(node.childs) && node.childs.length > 0;
    let treeview = '<li>';
    if (hasChildren) {
        treeview += createCheckboxElement(node, 'has-childs root', false);
        treeview += '<div class="check-toggle active"></div>';
        treeview += '<ul class="box-item box-toggle">';
        treeview += '<li>' + createCheckboxElement(node, '', true) + '</li>';
        for (const childNode of node.childs) {
            treeview += buildTreeviewNode(childNode);
        }
        treeview += '</ul>';
    } else {
        treeview += createCheckboxElement(node);
    }
    treeview += '</li>';
    return treeview;
}

/**
 * 左メニュー全体を構築 (Vue.js対応版)
 * jQuery依存を除去し、ネイティブDOM APIを使用
 */
function buildTreeview() {
    // ネイティブDOM APIでコンテナを取得
    const container = document.querySelector(".column-left > .column-left-container");
    if (!container) {
        console.warn('buildTreeview: .column-left-container not found');
        return;
    }
    
    const searchCatalogue = getSearchCatalogue();
    
    // DOM操作を一括で行うため、fragmentを利用
    const fragment = document.createDocumentFragment();
    searchCatalogue.forEach(catalogue => {
        const div = document.createElement('div');
        div.className = 'box-check';
        div.innerHTML = `<ul>${buildTreeviewNode(catalogue)}</ul>`;
        fragment.appendChild(div);
    });
    
    // コンテナをクリアしてfragmentを追加
    container.innerHTML = '';
    container.appendChild(fragment);
    
    // 状態を更新
    uiState.treeviewBuilt = true;
    
    // イベントハンドラーを設定
    setupTreeviewEventHandlers();
    
    // ツリービュー構築完了イベントを発火
    window.dispatchEvent(new CustomEvent('treeviewBuilt', {
        detail: { catalogue: searchCatalogue }
    }));
}

/**
 * ツリービューのイベントハンドラーを設定 (Vue.js対応版)
 * jQueryのイベントデリゲーションをネイティブイベントリスナーに変換
 */
function setupTreeviewEventHandlers() {
    // .box-check要素にイベントリスナーを設定（イベントデリゲーション）
    const boxCheckElements = document.querySelectorAll('.box-check');
    
    boxCheckElements.forEach(boxCheck => {
        // check-toggleのクリックイベント
        boxCheck.addEventListener('click', function(e) {
            if (e.target.classList.contains('check-toggle')) {
                onToggleClick(e);
            }
            // root labelのクリックイベント
            if (e.target.tagName === 'LABEL' && 
                e.target.classList.contains('custom-control-label') && 
                e.target.classList.contains('root')) {
                onRootLabelClick(e);
            }
        });
        
        // search-inのchangeイベント
        boxCheck.addEventListener('change', function(e) {
            if (e.target.classList.contains('search-in')) {
                onSearchInChange.call(e.target, e);
            }
            if (e.target.classList.contains('search-in-all')) {
                onSearchInAllChange.call(e.target, e);
            }
        });
    });
    
    // root checkboxのクリック防止（documentレベル）
    document.addEventListener('click', function(e) {
        if (e.target.type === 'checkbox' && 
            e.target.classList.contains('search-in') && 
            e.target.classList.contains('root')) {
            preventRootCheckboxClick(e);
        }
    });
    
    // 状態を更新
    uiState.eventHandlersAttached = true;
}

function onToggleClick(e) {
    const target = e.target;
    const parentLi = target.parentElement;
    
    // activeクラスをトグル
    target.classList.toggle('active');
    
    // 隣接する<ul>要素を取得してスライドトグル
    const nextUl = target.nextElementSibling;
    if (nextUl && nextUl.tagName === 'UL') {
        // 簡易的なスライドアニメーション（CSS transitionを使用することを推奨）
        if (nextUl.style.display === 'none' || !nextUl.style.display) {
            nextUl.style.display = 'block';
        } else {
            nextUl.style.display = 'none';
        }
    }
    
    // 兄弟要素の処理
    const siblings = Array.from(parentLi.parentElement.children).filter(child => child !== parentLi);
    siblings.forEach(sibling => {
        const siblingToggle = sibling.querySelector('.check-toggle');
        const siblingUl = sibling.querySelector('ul');
        
        if (siblingToggle) {
            siblingToggle.classList.remove('active');
        }
        if (siblingUl) {
            siblingUl.style.display = 'none';
        }
    });
}

function onRootLabelClick(e) {
    const label = e.target;
    const li = label.closest('li');
    if (li) {
        const checkToggle = li.querySelector('.check-toggle');
        if (checkToggle) {
            checkToggle.click();
        }
    }
}

function onSearchInChange(e) {
    const checkbox = this; // thisはイベントが発火されたチェックボックス
    const isChecked = checkbox.checked;
    
    // 親要素を取得
    const parentLi = checkbox.closest('li');
    if (!parentLi) return;
    
    const ul = parentLi.querySelector('ul');
    if (ul) {
        // 子要素のチェックボックスをすべて更新
        const checkboxes = ul.querySelectorAll('.search-in, .search-in-all');
        checkboxes.forEach(cb => {
            cb.checked = isChecked;
        });
    }
    
    // check-newクラスを削除
    const div = checkbox.closest('div');
    if (div) {
        div.classList.remove('check-new');
    }
    
    const checkNewElements = parentLi.querySelectorAll('.check-new');
    checkNewElements.forEach(el => el.classList.remove('check-new'));
    
    // 検索結果を再表示
    if (typeof displayResult === 'function') {
        displayResult();
    }
    
    // ツリー全体のチェック状態を更新
    checkAllInTree(checkbox);
}

function onSearchInAllChange(e) {
    const checkbox = this;
    const id = checkbox.id.replace('search-in-all-', '');
    const targetCheckbox = document.getElementById('search-in-' + id);
    
    if (targetCheckbox) {
        targetCheckbox.checked = checkbox.checked;
        // changeイベントを発火
        targetCheckbox.dispatchEvent(new Event('change', { bubbles: true }));
    }
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
 * 子チェックボックスにクリックイベントを追加 (Vue.js対応版)
 * jQueryのイベントデリゲーションをネイティブイベントリスナーに変換
 */
function addHandleEventInFirstPage() {
    // body要素にイベントリスナーを設定（イベントデリゲーション）
    document.body.addEventListener('click', function(e) {
        const target = e.target;
        
        // childチェックボックスの処理
        if (target.type === 'checkbox' && target.classList.contains('child')) {
            const boxS1 = target.closest('.box-s-1');
            if (!boxS1) return;
            
            const parentCheckbox = boxS1.querySelector('.parent');
            const ul = target.closest('ul');
            if (!ul) return;
            
            const siblings = ul.querySelectorAll('.child');
            const checkedStates = Array.from(siblings).map(cb => cb.checked);
            
            const allChecked = checkedStates.every(state => state);
            const someChecked = checkedStates.some(state => state);
            
            if (parentCheckbox) {
                parentCheckbox.checked = allChecked;
                const parentDiv = parentCheckbox.closest('div');
                if (parentDiv) {
                    if (someChecked && !allChecked) {
                        parentDiv.classList.add('check-new');
                    } else {
                        parentDiv.classList.remove('check-new');
                    }
                }
            }
            
            buildTreeview();
        }
        
        // parentチェックボックスの処理
        if (target.type === 'checkbox' && target.classList.contains('parent')) {
            const isChecked = target.checked;
            const boxS1 = target.closest('.box-s-1');
            if (!boxS1) return;
            
            const childCheckboxes = boxS1.querySelectorAll('.child');
            childCheckboxes.forEach(cb => {
                cb.checked = isChecked;
            });
            
            const div = target.closest('div');
            if (div) {
                div.classList.remove('check-new');
            }
            
            buildTreeview();
        }
    });
    
    // 状態を更新
    uiState.firstPageBuilt = true;
}

/**
 * UI状態を取得
 * @returns {Object} UI状態オブジェクト
 */
function getUIState() {
    return uiState;
}

