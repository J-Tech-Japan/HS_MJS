using System.Management.Automation;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.PowerShell;

// アセンブリに関する一般情報は、以下の属性セットによって
// 制御されます。アセンブリに関連付けられている情報を変更するには、
// これらの属性値を変更します。
[assembly: AssemblyTitle("MJSワードプラグイン")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("MJSワードプラグイン")]
[assembly: AssemblyCopyright("Copyright c  2017")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// ComVisible を false に設定すると、その型はこのアセンブリ内で COM コンポーネントから
// 参照不可能になります。COM からこのアセンブリ内の型にアクセスする場合は、
// その型の ComVisible 属性を true に設定してください。
[assembly: ComVisible(false)]

// このプロジェクトが COM に公開される場合、次の GUID が typelib の ID になります
[assembly: Guid("efcb7755-f1d8-4bb1-b051-137af1a308da")]

// Wordのリボンにはリビジョンを除く3つの数字でバージョンを表示します
// [メジャー バージョン.マイナー バージョン.ビルド番号]
// 初期値は 3.1.0 です。
// 以下の数字はビルドする度に自動的にインクリメントされます。
// (PowerShellスクリプト IncrementMinorVersion.ps1 で管理されます。
// 末尾の数字はリビジョン番号です。Wordには表示されません。
[assembly: AssemblyVersion("3.1.7.0")]
[assembly: AssemblyFileVersion("3.1.7.0")]

// 通常はビルド前イベントコマンドラインに以下のコマンドが設定されています（リリースビルド毎にインクリメント）。
// if "$(ConfigurationName)"=="Release" powershell -ExecutionPolicy Bypass -File "$(ProjectDir)IncrementMinorVersion.ps1"

// デバッグビルドでもインクリメントしたい場合は、以下のコードに書き換えます。
// powershell -ExecutionPolicy Bypass -File "$(ProjectDir)IncrementMinorVersion.ps1"

// バージョンを3.1.0にリセットする場合は、以下のように書き換えます。
// powershell - ExecutionPolicy Bypass - File "$(ProjectDir)IncrementMinorVersion.ps1" reset
