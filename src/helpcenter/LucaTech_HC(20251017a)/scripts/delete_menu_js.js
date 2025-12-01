const fs = require('fs');
const path = require('path');

/**
 * *_menu.jsファイルを一括削除するスクリプト
 */

const MENU_DIR = path.join(__dirname, '..', 'center', 'js', 'menu');

// すべての*_menu.jsファイルを検索
const files = fs.readdirSync(MENU_DIR).filter(file => 
  file.endsWith('_menu.js') && file !== 'convert_to_json.js' && file !== 'delete_menu_js.js'
);

console.log(`削除対象ファイル数: ${files.length}\n`);

if (files.length === 0) {
  console.log('削除対象のファイルが見つかりませんでした。');
  process.exit(0);
}

console.log('以下のファイルを削除します:');
files.forEach(file => {
  console.log(`  - ${file}`);
});

console.log('\n削除を実行しています...\n');

let successCount = 0;
let errorCount = 0;

files.forEach(file => {
  const filePath = path.join(MENU_DIR, file);
  
  try {
    fs.unlinkSync(filePath);
    console.log(`  ✅ 削除成功: ${file}`);
    successCount++;
  } catch (error) {
    console.error(`  ❌ 削除失敗: ${file} - ${error.message}`);
    errorCount++;
  }
});

console.log(`\n削除処理が完了しました！`);
console.log(`成功: ${successCount}件`);
if (errorCount > 0) {
  console.log(`失敗: ${errorCount}件`);
}
