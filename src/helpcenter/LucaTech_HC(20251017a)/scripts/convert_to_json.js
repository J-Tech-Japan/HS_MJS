const fs = require('fs');
const path = require('path');

/**
 * JSファイル（MENU = {...}またはPURPOSE = {...}形式）をJSONファイルに変換するスクリプト
 */

const MENU_DIR = path.join(__dirname, '..', 'center', 'js', 'menu');
const PURPOSE_DIR = path.join(__dirname, '..', 'center', 'js', 'purpose');

/**
 * 指定ディレクトリ内のJSファイルをJSON化する
 * @param {string} dir - 対象ディレクトリ
 * @param {string} suffix - ファイルサフィックス (_menu.js または _purpose.js)
 * @param {string} varName - 抽出する変数名 (MENU または PURPOSE)
 */
function convertToJson(dir, suffix, varName) {
  if (!fs.existsSync(dir)) {
    console.log(`ディレクトリが存在しません: ${dir}`);
    return;
  }

  // すべての対象ファイルを検索
  const files = fs.readdirSync(dir).filter(file => 
    file.endsWith(suffix) && file !== 'convert_to_json.js'
  );

  console.log(`\n[${path.basename(dir)}] 見つかったファイル数: ${files.length}`);

  files.forEach(file => {
    const jsFilePath = path.join(dir, file);
    const jsonFileName = file.replace('.js', '.json');
    const jsonFilePath = path.join(dir, jsonFileName);

    console.log(`\n処理中: ${file}`);

    try {
      // JSファイルを読み込み
      const jsContent = fs.readFileSync(jsFilePath, 'utf-8');

      // MENU = {...}; または PURPOSE = {...}; の形式から{...}部分を抽出
      const regex = new RegExp(`${varName}\\s*=\\s*(\\{[\\s\\S]*\\});?\\s*$`);
      const match = jsContent.match(regex);
      
      if (!match) {
        console.error(`  ❌ エラー: ${varName}変数が見つかりません - ${file}`);
        return;
      }

      const objectString = match[1];

      // JavaScriptオブジェクトをevalで評価
      let dataObject;
      try {
        // eslint-disable-next-line no-eval
        dataObject = eval('(' + objectString + ')');
      } catch (evalError) {
        console.error(`  ❌ エラー: JavaScript評価失敗 - ${file}`);
        console.error(`     ${evalError.message}`);
        return;
      }

      // JSONとして整形して保存
      const jsonContent = JSON.stringify(dataObject, null, 2);
      fs.writeFileSync(jsonFilePath, jsonContent, 'utf-8');

      console.log(`  ✅ 成功: ${jsonFileName} を作成しました`);

    } catch (error) {
      console.error(`  ❌ エラー: ${file} - ${error.message}`);
    }
  });
}

// menuファイルを変換
console.log('=== メニューファイルの変換 ===');
convertToJson(MENU_DIR, '_menu.js', 'MENU');

// purposeファイルを変換（purposeファイルもMENU変数を使用）
console.log('\n=== 目的ファイルの変換 ===');
convertToJson(PURPOSE_DIR, '_purpose.js', 'MENU');

console.log('\n変換処理が完了しました！');
console.log('\n次のステップ:');
console.log('1. 生成されたJSONファイルを確認してください');
console.log('2. menu.htmlとpurpose.htmlを更新してJSONを読み込むように変更してください');
console.log('3. js/menu と js/purpose のJSONファイルを data/menu と data/purpose に移動してください');
