/**
 * 左列の階層メニューHTMLを生成する関数
 * ツリー構造のメニューを再帰的に構築し、各項目にチェックボックスと展開/収納機能を付与
 * @param {Object} node - ノードデータ（ID, TITLE, CONTENTS等を含む）
 * @param {number} level - 現在の階層レベル（0から開始）
 * @returns {string} 生成されたHTML文字列
 */
function renderLeftColumn(node, level) {
    const checkboxClass = level === 0 ? 'custom-checkbox' : 'custom-checkbox leaf';
    const hasChildren = node.CONTENTS !== undefined;
    
    let html = `<li class="">
        <div class="${checkboxClass}">
            <input type="checkbox" class="search-in custom-control-input" id="checkbox-${node.ID}" checked>
            <label for="checkbox-${node.ID}" class="custom-control-label">${node.TITLE}</label>
        </div>`;
    
    if (hasChildren) {
        const childrenHtml = node.CONTENTS
            .map(child => renderLeftColumn(child, level + 1))
            .join('');
        
        html += `
        <span class='check-toggle active level${level}'>
            <i class='fas fa-caret-down fa-fw'></i>
        </span>
        <ul class='box-item box-toggle'>${childrenHtml}</ul>`;
    }
    
    html += `</li>`;
    return html;
}

/**
 * リーフノード（末端ページ）のHTMLを生成
 * @param {Object} content - コンテンツデータ
 * @returns {string} 生成されたHTML文字列
 */
function renderLeafNode(content) {
    return `<p><a target="_blank" href="${content.PATH}/index.html">${content.TITLE}</a></p>`;
}

/**
 * 階層ノード（サブメニュー）のHTMLを生成
 * @param {Object} content - コンテンツデータ
 * @returns {string} 生成されたHTML文字列
 */
function renderHierarchyNode(content) {
    const subContentHtml = content.CONTENTS
        .map(subContent => `<p><a target="_blank" href="${subContent.PATH}/index.html">${subContent.TITLE}</a></p>`)
        .join('');
    
    return `<ul id='more_services'>
        <li>
            <p>${content.TITLE}<span class='click-sp'><i class='fas fa-caret-down fa-fw'></i></span></p>
            <div class='secondary-services-description'>${subContentHtml}</div>
        </li>
    </ul>`;
}

/**
 * トップページのメインコンテンツHTMLを生成する関数
 * ノードデータに基づいて、カテゴリ別のコンテンツ表示エリアを構築
 * @param {Object} node - ノードデータ（TITLE, CONTENTS, STYLE等を含む）
 * @returns {string} 生成されたHTML文字列
 */
function renderTopPage(node) {
    let html = `<div class='row-box-1'>
        <h2>${node.TITLE}</h2>`;
    
    if (node.CONTENTS !== undefined) {
        const contentHtml = node.CONTENTS
            .map(content => content.CONTENTS === undefined 
                ? renderLeafNode(content) 
                : renderHierarchyNode(content))
            .join('');
        
        html += `<div class='nd-box nd-box-color-${node.STYLE}'>${contentHtml}</div>`;
    }
    
    html += `</div>`;
    return html;
}
