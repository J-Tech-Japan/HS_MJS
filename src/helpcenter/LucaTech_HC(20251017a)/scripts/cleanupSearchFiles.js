/**
 * 各コンテンツディレクトリのsearch.jsファイルから一行目を除いて削除するスクリプト
 * 
 * 使用方法:
 * node cleanupSearchFiles.js
 * 
 * 動作:
 * 1. contentsディレクトリ内のすべてのサブディレクトリを検索
 * 2. 各ディレクトリのsearch.jsファイルを特定
 * 3. 最初の行（var searchWords = ...）のみを残し、残りの行を削除
 * 4. 変更されたファイルの情報をコンソールに出力
 */

const fs = require('fs');
const path = require('path');

// 設定
const CONTENTS_DIR = path.join(__dirname, '..', 'contents');
const SEARCH_FILE_NAME = 'search.js';

/**
 * ディレクトリが存在するかチェック
 */
function checkDirectory(dirPath) {
    if (!fs.existsSync(dirPath)) {
        console.error(`エラー: contentsディレクトリが見つかりません: ${dirPath}`);
        process.exit(1);
    }
}

/**
 * ディレクトリ内のサブディレクトリを取得
 */
function getSubDirectories(dirPath) {
    try {
        return fs.readdirSync(dirPath, { withFileTypes: true })
            .filter(dirent => dirent.isDirectory())
            .map(dirent => dirent.name);
    } catch (error) {
        console.error(`エラー: ディレクトリの読み取りに失敗しました: ${dirPath}`, error.message);
        return [];
    }
}

/**
 * search.jsファイルから一行目のみを保持し、残りを削除
 */
function cleanupSearchFile(filePath) {
    try {
        // ファイルの存在確認
        if (!fs.existsSync(filePath)) {
            console.log(`スキップ: ファイルが存在しません - ${filePath}`);
            return false;
        }

        // ファイル内容を読み取り
        const content = fs.readFileSync(filePath, 'utf8');
        const lines = content.split('\n');

        // ファイルが空の場合
        if (lines.length === 0) {
            console.log(`スキップ: ファイルが空です - ${filePath}`);
            return false;
        }

        // 最初の行のみを保持
        const firstLine = lines[0];
        
        // 最初の行が var searchWords で始まることを確認
        if (!firstLine.trim().startsWith('var searchWords')) {
            console.log(`スキップ: ファイル形式が期待と異なります - ${filePath}`);
            return false;
        }

        // 最初の行のみでファイルを上書き
        fs.writeFileSync(filePath, firstLine + '\n', 'utf8');
        
        const deletedLines = lines.length - 1;
        console.log(`✅ 処理完了: ${filePath} (削除行数: ${deletedLines}行)`);
        return true;

    } catch (error) {
        console.error(`❌ エラー: ファイル処理に失敗しました - ${filePath}`, error.message);
        return false;
    }
}

/**
 * メイン処理
 */
function main() {
    console.log('='.repeat(60));
    console.log('search.js クリーンアップスクリプト開始');
    console.log('='.repeat(60));

    // contentsディレクトリの存在確認
    checkDirectory(CONTENTS_DIR);

    // サブディレクトリを取得
    const subDirectories = getSubDirectories(CONTENTS_DIR);
    
    if (subDirectories.length === 0) {
        console.log('処理対象のサブディレクトリが見つかりませんでした。');
        return;
    }

    console.log(`発見されたサブディレクトリ数: ${subDirectories.length}`);
    console.log('処理対象ディレクトリ:', subDirectories.join(', '));
    console.log('-'.repeat(60));

    let processedCount = 0;
    let successCount = 0;
    let skipCount = 0;
    let errorCount = 0;

    // 各サブディレクトリのsearch.jsファイルを処理
    subDirectories.forEach(subDir => {
        const searchFilePath = path.join(CONTENTS_DIR, subDir, SEARCH_FILE_NAME);
        processedCount++;

        console.log(`\n[${processedCount}/${subDirectories.length}] 処理中: ${subDir}/${SEARCH_FILE_NAME}`);
        
        const result = cleanupSearchFile(searchFilePath);
        if (result === true) {
            successCount++;
        } else if (result === false) {
            skipCount++;
        } else {
            errorCount++;
        }
    });

    // 結果サマリー
    console.log('\n' + '='.repeat(60));
    console.log('処理結果サマリー');
    console.log('='.repeat(60));
    console.log(`対象ファイル数: ${processedCount}`);
    console.log(`✅ 成功: ${successCount}ファイル`);
    console.log(`⏭️  スキップ: ${skipCount}ファイル`);
    console.log(`❌ エラー: ${errorCount}ファイル`);
    
    if (successCount > 0) {
        console.log('\n✨ クリーンアップが完了しました！');
    } else {
        console.log('\n⚠️  処理されたファイルがありませんでした。');
    }
}

// スクリプト実行
if (require.main === module) {
    main();
}

module.exports = {
    cleanupSearchFile,
    getSubDirectories,
    checkDirectory
};