const fs = require('fs');
const path = require('path');

/**
 * *_menu.jsと*_purpose.jsファイルを一括削除するスクリプト
 */

const MENU_DIR = path.join(__dirname, '..', 'center', 'js', 'menu');
const PURPOSE_DIR = path.join(__dirname, '..', 'center', 'js', 'purpose');

/**
 * 指定ディレクトリ内のJSファイルを削除
 * @param {string} dir - 対象ディレクトリ
 * @param {string} suffix - ファイルサフィックス (_menu.js または _purpose.js)
 */
function deleteJsFiles(dir, suffix) {
  if (!fs.existsSync(dir)) {
    console.log(`ディレクトリが存在しません: ${dir}`);
    return { success: 0, error: 0 };
  }

  // すべての対象ファイルを検索
  const files = fs.readdirSync(dir).filter(file => 
    file.endsWith(suffix) && file !== 'convert_to_json.js' && file !== 'delete_menu_js.js'
  );

  console.log(`\n[${path.basename(dir)}] 削除対象ファイル数: ${files.length}`);

  if (files.length === 0) {
    console.log('  削除対象のファイルが見つかりませんでした。');
    return { success: 0, error: 0 };
  }

  console.log('以下のファイルを削除します:');
  files.forEach(file => {
    console.log(`  - ${file}`);
  });

  console.log('\n削除を実行しています...');

  let successCount = 0;
  let errorCount = 0;

  files.forEach(file => {
    const filePath = path.join(dir, file);
    
    try {
      fs.unlinkSync(filePath);
      console.log(`  ✅ 削除成功: ${file}`);
      successCount++;
    } catch (error) {
      console.error(`  ❌ 削除失敗: ${file} - ${error.message}`);
      errorCount++;
    }
  });

  return { success: successCount, error: errorCount };
}

console.log('=== JSファイルの削除 ===');

// menuファイルを削除
console.log('\n--- メニューファイルの削除 ---');
const menuResult = deleteJsFiles(MENU_DIR, '_menu.js');

// purposeファイルを削除
console.log('\n--- 目的ファイルの削除 ---');
const purposeResult = deleteJsFiles(PURPOSE_DIR, '_purpose.js');

console.log('\n=== 削除処理が完了しました！ ===');
console.log(`合計成功: ${menuResult.success + purposeResult.success}件`);
if (menuResult.error + purposeResult.error > 0) {
  console.log(`合計失敗: ${menuResult.error + purposeResult.error}件`);
}
