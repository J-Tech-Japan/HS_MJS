const fs = require('fs');
const path = require('path');

/**
 * JSファイル（MENU = {...}形式）をJSONファイルに変換するスクリプト
 */

const MENU_DIR = path.join(__dirname, '..', 'center', 'js', 'menu');

// すべての*_menu.jsファイルを検索
const files = fs.readdirSync(MENU_DIR).filter(file => 
  file.endsWith('_menu.js') && file !== 'convert_to_json.js'
);

console.log(`見つかったファイル数: ${files.length}`);

files.forEach(file => {
  const jsFilePath = path.join(MENU_DIR, file);
  const jsonFileName = file.replace('.js', '.json');
  const jsonFilePath = path.join(MENU_DIR, jsonFileName);

  console.log(`\n処理中: ${file}`);

  try {
    // JSファイルを読み込み
    const jsContent = fs.readFileSync(jsFilePath, 'utf-8');

    // MENU = {...}; の形式から{...}部分を抽出
    const match = jsContent.match(/MENU\s*=\s*(\{[\s\S]*\});?\s*$/);
    
    if (!match) {
      console.error(`  ❌ エラー: MENU変数が見つかりません - ${file}`);
      return;
    }

    const menuObjectString = match[1];

    // JavaScriptオブジェクトをevalで評価
    let menuObject;
    try {
      // eslint-disable-next-line no-eval
      menuObject = eval('(' + menuObjectString + ')');
    } catch (evalError) {
      console.error(`  ❌ エラー: JavaScript評価失敗 - ${file}`);
      console.error(`     ${evalError.message}`);
      return;
    }

    // JSONとして整形して保存
    const jsonContent = JSON.stringify(menuObject, null, 2);
    fs.writeFileSync(jsonFilePath, jsonContent, 'utf-8');

    console.log(`  ✅ 成功: ${jsonFileName} を作成しました`);

  } catch (error) {
    console.error(`  ❌ エラー: ${file} - ${error.message}`);
  }
});

console.log('\n変換処理が完了しました！');
console.log('\n次のステップ:');
console.log('1. 生成されたJSONファイルを確認してください');
console.log('2. menu.htmlとpurpose.htmlを更新してJSONを読み込むように変更してください');
