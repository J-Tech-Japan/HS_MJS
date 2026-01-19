var searchWords = $('<div class="search"><div id="MAS12001"><div class="search_breadcrumbs">導入処理（部門設定） &gt; はじめに</div></div><div id="MAS12003"><div class="search_breadcrumbs">導入処理（部門設定） &gt; 部門関係 &gt; 配賦情報を登録する</div><div class="search_title">配賦情報を登録する</div><div class="displayText">配賦情報を登録する「配賦」とは、賃借料や水道光熱費などの複数の部門で共通に負担すべき金額を各部門に割り振る（按分する）ことです。配賦情報を登録することで、各部門への按分を自動で行い ...</div><div class="search_word">配賦情報を登録する「配賦」とは、賃借料や水道光熱費などの複数の部門で共通に負担すべき金額を各部門に割り振る（按分する）ことです。配賦情報を登録することで、各部門への按分を自動で行います。 配賦情報の登録と処理の流れ処理の流れ配賦の採用初めて配賦情報を登録する場合は、『財務基本情報登録』の［処理区分：科目情報］で配賦を採用します。⇒『部門』参照配賦情報の登録配賦の条件を登録します。人数や床面積などの任意の数値を基準にして配賦を行う場合は、先に配賦基準値を登録します。⇒『配賦基準値を登録する』参照マスター更新『マスター更新』を行うことで、登録した配賦情報に従い配賦が行われます。配賦結果を出力し、登録した配賦情報が正しく反映されるか確認します。⇒『マスター更新を行う』参照 『固定資産管理』を使用していて、減価償却費の部門配賦を『固定資産管理』側で設定する場合は、［設定］＞［導入処理］＞［資産管理］＞［配賦パターン登録］で行います。『配賦パターン登録』については、次を参照してください。⇒『配賦パターン情報を登録する』参照 配賦結果の反映方法について配賦結果は、次の2通りの方法でデータに反映させることができます。配賦結果の反映方法は、［財務基本情報登録］＞［処理区分：科目情報］＞［部門］の［配賦採用区分］で設定します。 帳票上でのみ管理する仕訳を作成せず、帳票を表示するときに出力条件で指定し、集計して確認します。✔をはずして出力すると、配賦前の金額を確認することもできます。  自動仕訳を作成する配賦元の金額を相殺する仕訳を自動で作成します。仕訳の日付は末日で作成されます。『部門別集計表』などの帳票には、配賦結果のみが表示されます。 広告宣伝費を営業一課と営業二課に配賦する場合の自動仕訳の例借方科目借方部門貸方科目貸方部門消費税金額摘要広告宣伝費営業一課複合 40625,000部門配賦自動仕訳／［配賦パターン名称］広告宣伝費営業二課複合 40375,000部門配賦自動仕訳／［配賦パターン名称］複合 広告宣伝費共通部門401,000,000部門配賦自動仕訳／［配賦パターン名称］ 配賦元情報に設定した勘定科目に補助が設定されている場合、自動仕訳の補助には［諸口］がセットされます。 自動仕訳の摘要欄には、配賦パターン名称を表示させることができます。設定は［財務基本情報登録］＞［処理区分：科目情報］＞［部門］の［配賦摘要区分］で行います。</div></div><div id="MAS12004"><div class="search_breadcrumbs">導入処理（部門設定） &gt; 部門関係 &gt; 配賦情報を登録する &gt; 配賦パターンを登録する</div><div class="search_title">配賦パターンを登録する</div><div class="displayText">配賦パターンを登録する　［設定］＞［マスター登録処理］＞［部門関係］＞［部門配賦情報］＞［配賦情報］タブ実績金額を配賦するパターンを登録します。予算の配賦情報の登録については、次を ...</div><div class="search_word">配賦パターンを登録する　［設定］＞［マスター登録処理］＞［部門関係］＞［部門配賦情報］＞［配賦情報］タブ実績金額を配賦するパターンを登録します。予算の配賦情報の登録については、次を参照してください。⇒『予算の配賦情報を登録する』参照 人数や床面積などの任意の数値を基準にして配賦を行う場合は、先に配賦基準値を登録します。配賦基準値の登録については、次を参照してください。⇒『配賦基準値を登録する』参照   配賦パターンの登録順について配賦は、マスターの更新時にすべてのパターンがNo.順に実行されます。配賦を2段階以上に分けて行う場合は、実行する順番に配賦パターンNo.を登録します。  配賦パターン任意のコードと名称を入力します。コードは6桁、名称は30文字以内で入力します。作成済みの配賦パターンをコピーする場合や、配賦パターンNo.を変更する場合は、画面右上の［…］をクリックして表示される［配賦パターンコピー］から行います。配賦パターンコピーについては、次を参照してください。⇒『配賦パターンをコピーする』参照適用期間配賦パターンを使用する期間を指定する場合に入力します。登録済みの配賦パターンの適用期間は、画面右上の［…］をクリックして表示される［配賦適用期間確認］から一覧で確認できます。［配賦適用期間確認画面］ 配賦元情報 部門配賦元金額が発生する部門（配賦元金額を計上している部門）を選択します。科目指定方法配賦元の科目の指定方法を選択します。［単一］１つの科目を配賦する場合に指定します。［単一］の場合は、配賦元と配賦先で異なる科目を選択できます。たとえば、広告宣伝費に発生した金額を、別の科目の販売促進費として配賦できます。［科目範囲指定］［複数指定］複数の科目を配賦する場合に指定します。［科目範囲指定］［複数指定］の場合は、配賦元と配賦先は同じ科目になります。たとえば、配賦元の科目に「4432 地代家賃」と「4434 リース料」の複数を指定した場合、「4432 地代家賃」は「4432 地代家賃」に、「4434 リース料」は「4434 リース料」に配賦されます。開始科目終了科目科目［科目指定方法］に従って、配賦元金額が発生する科目を指定します。［部門］で選択した部門分類（P/L・B/S）の科目が指定できます。金額相殺区分［科目指定方法］が［単一］かつ、実在科目を指定している場合に指定します。配賦元の科目とは異なる科目を配賦先に指定する場合、配賦元の金額の相殺を、どちらの科目で行うか選択します。初期設定は［配賦先］です。相殺区分の考え方については、次を参照してください。⇒『配賦を相殺する科目について』参照［科目範囲指定］［複数選択］の場合や、［単一］で合計科目を指定している場合は、配賦元科目と配賦先科目は同一になります。  配賦基準情報 基準科目配賦の基準を選択します。基準は3つまで指定できます。［通常科目］特定の勘定科目（実在科目、合計科目、非会計科目）を指定し、その当月の実績（発生高）の比率で按分します。たとえば、営業部全体の広告宣伝費を、各営業課の売上高の比率に応じて按分したいときなどに使用します。［配賦基準値］人数や床面積などの任意の数値を基準にし、その比率で按分します。たとえば、本社ビルにかかる水道光熱費を、本社ビル内に存在する部門の床面積で按分したいときなどに使用します。［配賦基準値］タブから事前に数値を登録します。登録方法は次を参照してください。⇒『配賦基準値を登録する』参照基準が1つだけの場合は、［配賦割合］に「100」を入力します。基準が複数の場合は、［配賦割合］の合計が「100」になるように割合を入力します。マイナス配賦基準値の配賦計算方法［基準科目］がマイナスのときに、その部門を配賦の計算の対象外にする（0として計算する）場合は［配賦計算に含めない］を選択します。  配賦先情報 科目［科目指定方法］が［単一］の場合に、配賦結果を計上する実在科目を指定します。配賦元の科目を範囲または複数指定している場合は、配賦先の科目を指定できません（配賦元と同一の科目に配賦します）。組織部門配賦結果を計上する組織（部門体系）と部門を指定します。合計部門を指定する場合は、［組織］を［組織指定なし］以外にします。合計部門を指定した場合、その配下のすべての部門が配賦の対象となります。選択した組織の適用期間と配賦パターンの適用期間とで期間が異なる場合、重複した期間がある期間のみ配賦が行われます。  自動仕訳情報自動仕訳を起票するときの仕訳の内容を設定します。［財務基本情報登録］＞［処理区分：科目情報］＞［部門］＞［その他情報］＞［配賦採用区分］を［採用する（自動仕訳）］にしている場合に表示されます。  配賦結果を自動仕訳で記帳自動仕訳を作成する必要がない場合は、［記帳しない］を選択します。自動仕訳を作成しない場合の配賦結果は、帳票の集計時に確認できます。配賦結果の反映方法については、次を参照してください。⇒『配賦結果の反映方法について』参照相手科目コード自動仕訳の相手科目を「複合」から変更する場合に、相手科目に使用する勘定科目を選択します。消費税コード自動仕訳に使用する消費税コードを指定します。配賦先科目に応じた消費税コードを選択します。なお、［相手科目コード］が「複合」「資金複合」の場合は、［40 不課税（精算取引）］が使用されます。消費税率コード自動仕訳に使用する消費税率を選択します。基本的には［0 標準税率］を設定します。［0 標準税率］を設定すると、『マスター更新』を行ったときに、仕訳の日付に応じた消費税率がセットされます。開始伝票No.自動仕訳の開始伝票No.を指定します。伝票No変位［開始伝票No.］以降に発番する伝票No.の間隔を［伝票No変位］に登録します。（例）［開始伝票No.］を1001とする場合［0］を入力する場合作成される仕訳の伝票番号は、［開始伝票No.］で登録した伝票No.で固定されます。伝票番号はすべて「1001」になります。［1］を入力する場合作成される仕訳の伝票番号は、［開始伝票No.］で登録した伝票番号から1ずつ加算されます。伝票番号は「1001、1002、1003、1004…」となります。［5］を入力する場合作成される仕訳の伝票番号は、［開始伝票No.］で登録した伝票番号から5ずつ加算されます。伝票番号は「1001、1006、1011、1016…」となります。 配賦を相殺する科目について［科目指定方法］を［単一］にしており、配賦元の科目とは異なる科目を配賦先に指定する場合、配賦元の金額の相殺を配賦元と配賦先のどちらの科目で行うか選択します。相殺方法は、［配賦元情報］の［金額相殺区分］で指定します。相殺方法をどちらにするかは、確認に用いる帳票や、自社の運用、作成する配賦パターンなどに応じて選択してください。 共通部門に計上した広告宣伝費を、各部門の販売促進費として配賦する場合の例配賦元の科目で相殺する場合 借方貸方金額配賦前広告宣伝費（共通部門）現金1,000,000 配賦後販売促進費（営業一課）広告宣伝費（共通部門）600,000 販売促進費（営業二課）広告宣伝費（共通部門）400,000配賦元の科目（広告宣伝費）の残高は、相殺されて0円になります。全社を対象にした帳票（配賦前の帳票）と、配賦結果を反映させた部門別帳票とで、各科目の残高が変わります。配賦前の帳票：広告宣伝費1,000,000円　販売促進費　　　　 0円配賦後の帳票：広告宣伝費　　　　 0円　販売促進費1,000,000円 配賦先の科目で相殺する場合 借方貸方金額配賦前広告宣伝費（共通部門）現金1,000,000 配賦後販売促進費（営業一課）販売促進費（共通部門）600,000 販売促進費（営業二課）販売促進費（共通部門）400,000配賦元の科目（広告宣伝費）の残高は相殺されずに残り、各部門の販売促進費が共通部門の販売促進費と相殺されて残高が0円となります。全社を対象にした帳票（配賦前の帳票）と、配賦結果を反映させた部門別帳票とで、各科目の残高は変わりません。配賦前の帳票：広告宣伝費1,000,000円　販売促進費　　　　 0円配賦後の帳票：広告宣伝費1,000,000円　販売促進費　　　　 0円  配賦パターンをコピーする登録済みの配賦パターンをコピーして新規に配賦パターンを作成する場合は、画面右上の［…］をクリックして表示される［配賦パターンコピー］画面から行います。 配賦パターンのコピーは、実績と予算の間でコピーすることもできます。また、登録済みの配賦パターンのパターンNo.を変更する場合にも使用します。  配賦区分コピー元とコピー先の配賦区分（実績または予算）を選択します。コピー方法配賦パターンのコピー方法を選択します。［コピー先で指定したパターンNo.より連番でコピーする］コピー元の配賦パターンをコピーして、指定したパターンNo.から連番で配賦パターンを作成します。コピー先に同じパターンNo.がすでに存在する場合は、重複しないもののみコピーするか、すべて上書きするかを次の画面で指定します。［配賦パターンNo.］の［コピー先］に開始No.を指定します。［コピー元のパターンNo.でコピーする］コピー元の配賦パターンをコピーして、同じNo.にコピーします。コピー元とコピー先とで配賦区分が異なる場合に使用できます。［パターンNo.を変更する］登録済みの配賦パターンのNo.を変更します。コピー元とコピー先とで配賦区分が同じ場合に使用できます。配賦パターンNo.コピー元の配賦パターンNo.を指定します。［コピー方法］を［コピー先で指定したパターンNo.より連番でコピーする］にした場合は、［コピー先］に開始No.を指定します。 登録した配賦パターンを一覧で確認する登録した配賦パターンの一覧は、CSVファイルに出力するほか、画面右上の［…］をクリックして表示される［配賦適用期間確認画面］からも確認できます。 配賦パターンを削除する登録した配賦パターンを削除する場合は、画面右上の［…］をクリックして表示される［配賦パターン一括削除］画面から行います。</div></div><div id="MAS12008"><div class="search_breadcrumbs">導入処理（部門設定） &gt; 部門関係 &gt; 配賦情報を登録する &gt; 配賦基準値を登録する</div><div class="search_title">配賦基準値を登録する</div><div class="displayText">配賦基準値を登録する　［設定］＞［マスター登録処理］＞［部門関係］＞［部門配賦情報］＞［配賦基準値］タブ人数、床面積、稼働時間などの任意の数値を基準にして配賦を行う場合に、配賦基準 ...</div><div class="search_word">配賦基準値を登録する　［設定］＞［マスター登録処理］＞［部門関係］＞［部門配賦情報］＞［配賦基準値］タブ人数、床面積、稼働時間などの任意の数値を基準にして配賦を行う場合に、配賦基準値を登録します。配賦基準値の登録は、［配賦情報］タブで配賦パターンを登録する前に行います。 配賦基準値は、CSVファイルをインポートして登録することもできます。CSVファイルをインポートして登録する方法については、次を参照してください。⇒『CSVファイルをインポートして配賦基準値を登録する』参照  基準科目コード任意のコードと名称を入力します。コードは6桁、名称は30文字以内で入力します。対象期間配賦基準値が１年間の会計期間を通して同じ場合は［年間］、月別に設定する場合は［月別］を選択します。 部門コード［新規行追加］をクリックして、部門を追加します。配賦基準値部門ごとに基準となる数値を13桁以内で入力します。小数点以下は、第4位まで入力できます。 配賦基準値をコピーする登録済みの配賦基準値をコピーします。別の基準科目コードへの配賦基準値のコピーの他に、特定の月の配賦基準値を同じ基準科目コード内の別の月へコピーすることもできます。 配賦基準値のコピーは、画面右上の［…］をクリックして表示される［配賦基準値コピー］画面から行います。  コピー元コピーする基準科目コードを選択します。［対象期間］が［月別］の基準科目コードを選択した場合は、コピー元の月を入力します。決算月は、決算月［91］、決算月2［92］、決算月3［93］と入力します。コピー先コピー先の基準科目コードを指定します。コピー元と［対象期間］が同じ基準科目を指定できます。また、コピー先は、複数選択できます。［月別］の場合は、［コピー開始月］［コピー終了月］を指定します。決算月は、決算月［91］、決算月2［92］、決算月3［93］と入力します。 CSVファイルをインポートして配賦基準値を登録するCSVファイルをインポートして配賦基準値を登録します。登録する基準科目に、すでに配賦基準値が登録されている場合は、上書きされます。 ①　画面右上の［…］から［配賦基準値インポート］をクリックします。②　［CSVテンプレート出力］をクリックして、入力用のCSVファイルをダウンロードします。なお、CSVファイルには、登録済みの配賦基準値は出力されません。③　ダウンロードしたCSVファイルに、次のように数値や文字列を入力します。CSVファイルの作成例（年間の配賦基準値を登録する場合） 文字列を入力するときは半角のダブルクォーテーション「”」で囲みます。 CSVファイルのデータ項目項目内容基準科目コード基準科目のコードを6桁以内で入力します。すでに登録済みの基準科目コードを入力した場合は、内容が上書きされます。基準科目名称基準科目の名称を30文字以内で入力します。月別管理区分年間の場合は「0」、月別の場合は「1」を入力します。月別管理名称［月別管理区分］に合わせて、月別管理名称の「年間」または「月別」を入力します。実際の取込時は［月別管理区分］をもとにして内容が判断されるため、［月別管理名称］を入力しなくても取込できます。コード部門コードを入力します。部門名称部門名称を入力します。実際の取込時は［コード］をもとにして部門が判断されるため、［部門名称］を入力しなくても取込できます。配賦基準値［月別管理区分］が［年間］の場合に、年間の基準値を入力します。XXXX年X月～［月別管理区分］が［月別］の場合に、月別の基準値を入力します。採用していない決算月の入力は不要です。 ④　CSVファイルを保存して閉じます。⑤　［配賦基準値インポート］画面の［ファイル選択］をクリックして、作成したCSVファイルを選択します。⑥　［インポート］をクリックします。配賦基準値がインポートされます。 インポートでエラーが発生すると、次のメッセージが表示されます。［ログ表示］をクリックして内容を確認し、CSVファイルを修正してから再度インポートします。</div></div><div id="MAS12011"><div class="search_breadcrumbs">導入処理（部門設定） &gt; 部門関係 &gt; 配賦情報を登録する &gt; 予算の配賦情報を登録する</div><div class="search_title">予算の配賦情報を登録する</div><div class="displayText">予算の配賦情報を登録する　［設定］＞［マスター登録処理］＞［部門関係］＞［部門配賦情報］＞［予算配賦情報］タブ予算を配賦するための情報を登録します。予算の配賦は、［設定］＞［導入処 ...</div><div class="search_word">予算の配賦情報を登録する　［設定］＞［マスター登録処理］＞［部門関係］＞［部門配賦情報］＞［予算配賦情報］タブ予算を配賦するための情報を登録します。予算の配賦は、［設定］＞［導入処理］＞［会社情報］＞［マスター採用情報登録］で予算を採用している場合に使用できます。 ［予算配賦情報］タブの登録項目は、［配賦情報］タブと同じです。項目の説明については、次を参照してください。⇒『配賦パターンを登録する』参照 なお、予算の配賦では自動仕訳を行わないため、自動仕訳に関する設定項目はありません。予算の配賦結果は、『予算登録』で確認します。『予算登録』については、次を参照してください。⇒『予算登録の概要』参照</div></div><div id="MAS12012"><div class="search_breadcrumbs">導入処理（部門設定） &gt; 部門関係 &gt; 配賦情報を登録する &gt; 配賦処理を実行する</div><div class="search_title">配賦処理を実行する</div><div class="displayText">配賦処理を実行する　［財務大将］＞［日常処理］＞［マスター］＞［マスター更新］登録した配賦情報は、『マスター更新』でデータに反映させます。 マスター更新を行うときは、［配賦結果CS ...</div><div class="search_word">配賦処理を実行する　［財務大将］＞［日常処理］＞［マスター］＞［マスター更新］登録した配賦情報は、『マスター更新』でデータに反映させます。 マスター更新を行うときは、［配賦結果CSV出力］に✔をつけて配賦結果を確認し、配賦金額に相違がある場合は配賦情報を見直します。 ［マスター更新］画面 マスター更新の操作方法については、次を参照してください。⇒『マスター更新を行う』参照  配賦結果を削除する配賦結果をすべて削除する場合は、削除したい配賦パターンの［適用期間］を、マスター更新する期間に重複しない期間に変更してから、再度マスター更新を行います。 『マスター更新』は、［詳細設定］を［ON］にして、［更新済みの月も更新する］に✔をつけて行います。   前年の配賦結果を当期の期首残高から削除する翌期以降の残高登録において、期首残高を部門配賦前の金額で開始したい場合は、『勘定科目残高』の［一括削除］で、［配賦金額のみ削除］を［適用する］にします。 『勘定科目残高』の［一括削除］画面 『勘定科目残高』の操作方法については、次を参照してください。⇒『登録した残高を削除する』参照</div></div><div id="MAS12015"><div class="search_breadcrumbs">導入処理（部門設定） &gt; 部門関係 &gt; 部門出力順序を登録する</div><div class="search_title">部門出力順序を登録する</div><div class="displayText">部門出力順序を登録する　［設定］＞［マスター登録処理］＞［部門関係］＞［部門出力順序設定］部門別集計表・部門別財務報告書を表示・印刷するときの、部門の並び順を登録します（部門出力順 ...</div><div class="search_word">部門出力順序を登録する　［設定］＞［マスター登録処理］＞［部門関係］＞［部門出力順序設定］部門別集計表・部門別財務報告書を表示・印刷するときの、部門の並び順を登録します（部門出力順序）。部門出力順序は、1つの帳票に対して、用途に応じた複数の出力パターンを登録できます。 『月次管理表』の［部門別集計表］で部門出力順序を選択して出力した結果の例『部門出力順序設定』の画面説明は次のとおりです。 1部門出力順序メニュー『部門出力順序登録』で登録済みの出力パターンが表示されます。2部門出力順序登録エリア部門の並び順や装飾などを登録するエリアです。登録項目の説明については、次を参照してください。⇒『部門の表示方法を変更する』参照3所属体系一覧『部門所属体系登録』に登録されている所属体系が一覧で表示されます。ここに表示されている部門を、部門出力順序登録エリアに追加して、出力順序を登録します。［所属パターン］から所属体系のパターンを切り替えます。4［…］［パターンコピー］登録済みの部門出力順序のパターンをコピーして新しいパターンを登録します。なおコピー先で登録済みのパターンを選択した場合、部門出力順序が上書きされます。［出力順序登録チェック］部門出力順序に未登録の部門や、重複して登録されている部門がないかをチェックします。⇒『部門の未登録や重複を確認する』参照  部門の出力順序を登録する部門出力順序に部門を追加し、帳票のレイアウトを修正するまでの主な流れは次のとおりです。①　出力パターンを新規に作成する場合は、［出力パターン］に未使用のコードとパターン名を入力するか、［…］の［パターンコピー］で登録済みの帳票をコピーして作成します。②　部門比を算出するための分母となる部門を［分母部門］で選択します。［部門比］の項目名称を変更する場合は、［タイトル］を入力します。③　［新規順序行追加］をクリック、または部門体系一覧から部門出力順序に追加する部門を任意の行にドラッグ&amp;ドロップします。 部門出力順序に未登録の部門を絞り込んで表示する場合は、［未登録絞込］から行います。削除したい場合は、✔をつけて［削除］をクリック、または体系一覧にドラッグ&amp;ドロップします。追加した部門を並び替える場合は、行の左端をクリックして任意の行にドラッグ＆ドロップします。④　続けて、部門の表示方法などを必要に応じて修正します。修正方法については、次を参照してください。⇒『部門の表示方法を変更する』参照⑤　すべての設定が完了したら、［更新］をクリックします。⑥　［…］の［出力順序登録チェック］をクリックして、部門出力順序に未登録の部門や、重複して登録されている部門がないかをチェックします。⇒『部門の未登録や重複を確認する』参照⑦　部門出力順序の登録後は、チェックリストや実際の帳票で内容を確認します。チェックリストで確認する場合は、画面右上のからCSVファイルを出力します。</div></div><div id="MAS12016"><div class="search_breadcrumbs">導入処理（部門設定） &gt; 部門関係 &gt; 部門出力順序を登録する &gt; 部門の表示方法を変更する</div><div class="search_title">部門の表示方法を変更する</div><div class="displayText">部門の表示方法を変更する部門出力順序に追加した部門の表示方法などを変更します。 操作［…］から強制空白の設定や行の挿入・削除を行います。［強制空白］選択中の行を空白行に変更します。 ...</div><div class="search_word">部門の表示方法を変更する部門出力順序に追加した部門の表示方法などを変更します。 操作［…］から強制空白の設定や行の挿入・削除を行います。［強制空白］選択中の行を空白行に変更します。［挿入］選択中の行の上に新規行を追加します。［削除］選択中の行を削除します。部門コード部門名称をクリックして部門を選択します。部門比分母コード部門比分母名称選択中の部門に対して、個別に部門比分母を設定する場合に、分母にする部門を選択します。指定がない場合は、部門出力順序登録エリアの［共通部門比］で指定した部門が部門比の分母になります。装飾区分画面表示・印刷時の装飾を選択します。設定例</div></div><div id="MAS12017"><div class="search_breadcrumbs">導入処理（部門設定） &gt; 部門関係 &gt; 部門出力順序を登録する &gt; 部門の未登録や重複を確認する</div><div class="search_title">部門の未登録や重複を確認する</div><div class="displayText">部門の未登録や重複を確認するチェックログを出力し、部門出力順序に未登録の部門、重複して登録されている部門がないかを確認します。①　画面右上の［…］の［出力順序登録チェック］をクリッ ...</div><div class="search_word">部門の未登録や重複を確認するチェックログを出力し、部門出力順序に未登録の部門、重複して登録されている部門がないかを確認します。①　画面右上の［…］の［出力順序登録チェック］をクリックします。②　チェック条件を指定し、［チェック］をクリックします。出力パターンチェックの対象とする出力パターンを指定します。部門マスター登録日新しい部門を登録した場合など、マスターの登録日でチェックの対象を絞り込む場合は、年月日を入力します。指定がない場合は空欄にします。適用期間部門の適用期間でチェックの対象を絞り込む場合は、年月日を入力します。指定がない場合は空欄にします。未登録と重複している部門のチェックが始まります。チェックが完了すると、次のメッセージが表示されます。［ログ表示］をクリックすると、別ウィンドウに［処理ログ表示］が表示されます。③　［処理ログ表示］タブを閉じて『部門出力順序設定』に戻ります。④　未登録の部門や、重複して登録されている部門がある場合は修正し、［更新］をクリックします。なお、部門が重複している場合は対象の部門の背景がピンク色で表示されます。未登録の部門の場合は、ログで確認します。</div></div></div>');
// ========================================
// グローバル変数定義
// ========================================

// MutationObserver機能用のグローバル変数
var currentSearchValue = ""; // 現在の検索キーワード
var mutationObserver = null; // MutationObserverインスタンス
var debounceTimer = null; // DOM変更用のデバウンスタイマー

// jQueryセレクタキャッシュ
var $cachedElements = {
    iframe: null,
    searchField: null,
    searchResultItemsBlock: null,
    searchResultsEnd: null,
    searchMsg: null,
    searchInput: null
};

// セレクタマップ（キャッシュ機構の効率化）
var selectorMap = {
    iframe: "iframe.topic",
    searchField: ".wSearchField",
    searchResultItemsBlock: ".wSearchResultItemsBlock",
    searchResultsEnd: ".wSearchResultsEnd",
    searchMsg: "#searchMsg",
    searchInput: ".search-input"
};

// HTMLエンティティ復元用の正規表現（効率化のためグローバル定義）
var htmlEntityRegexes = {
    nbsp: /&nbsp;(?=[^<>]*<)/gm,
    gt: /&gt;(?=[^<>]*<)/gm,
    lt: /&lt;(?=[^<>]*<)/gm,
    quot: /&quot;(?=[^<>]*<)/gm,
    amp: /&amp;(?=[^<>]*<)/gm
};

// 文字変換マップ（効率化のため初期化時に正規表現を構築）
var characterMappings = (function () {
    // 文字変換用の配列定義（配列リテラルに変更）
    var wide = ["０", "１", "２", "３", "４", "５", "６", "７", "８", "９", "ａ", "ｂ", "ｃ", "ｄ", "ｅ", "ｆ", "ｇ", "ｈ", "ｉ", "ｊ", "ｋ", "ｌ", "ｍ", "ｎ", "ｏ", "ｐ", "ｑ", "ｒ", "ｓ", "ｔ", "ｕ", "ｖ", "ｗ", "ｘ", "ｙ", "ｚ", "ガ", "ギ", "グ", "ゲ", "ゴ", "ザ", "ジ", "ズ", "ゼ", "ゾ", "ダ", "ヂ", "ヅ", "デ", "ド", "バ", "ビ", "ブ", "ベ", "ボ", "パ", "ピ", "プ", "ペ", "ポ", "。", "「", "」", "、", "ヲ", "ァ", "ィ", "ゥ", "ェ", "ォ", "ャ", "ュ", "ョ", "ッ", "ー", "ア", "イ", "ウ", "エ", "オ", "カ", "キ", "ク", "ケ", "コ", "サ", "シ", "ス", "セ", "ソ", "タ", "チ", "ツ", "テ", "ト", "ナ", "ニ", "ヌ", "ネ", "ノ", "ハ", "ヒ", "フ", "ヘ", "ホ", "マ", "ミ", "ム", "メ", "モ", "ヤ", "ユ", "ヨ", "ラ", "リ", "ル", "レ", "ロ", "ワ", "ン"];
    var narrow = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "ｶﾞ", "ｷﾞ", "ｸﾞ", "ｹﾞ", "ｺﾞ", "ｻﾞ", "ｼﾞ", "ｽﾞ", "ｾﾞ", "ｿﾞ", "ﾀﾞ", "ﾁﾞ", "ﾂﾞ", "ﾃﾞ", "ﾄﾞ", "ﾊﾞ", "ﾋﾞ", "ﾌﾞ", "ﾍﾞ", "ﾎﾞ", "ﾊﾟ", "ﾋﾟ", "ﾌﾟ", "ﾍﾟ", "ﾎﾟ", "｡", "｢", "｣", "､", "ｦ", "ｧ", "ｨ", "ｩ", "ｪ", "ｫ", "ｬ", "ｭ", "ｮ", "ｯ", "ｰ", "ｱ", "ｲ", "ｳ", "ｴ", "ｵ", "ｶ", "ｷ", "ｸ", "ｹ", "ｺ", "ｻ", "ｼ", "ｽ", "ｾ", "ｿ", "ﾀ", "ﾁ", "ﾂ", "ﾃ", "ﾄ", "ﾅ", "ﾆ", "ﾇ", "ﾈ", "ﾉ", "ﾊ", "ﾋ", "ﾌ", "ﾍ", "ﾎ", "ﾏ", "ﾐ", "ﾑ", "ﾒ", "ﾓ", "ﾔ", "ﾕ", "ﾖ", "ﾗ", "ﾘ", "ﾙ", "ﾚ", "ﾛ", "ﾜ", "ﾝ"];
    var highlight = ["(?:０|0)", "(?:１|1)", "(?:２|2)", "(?:３|3)", "(?:４|4)", "(?:５|5)", "(?:６|6)", "(?:７|7)", "(?:８|8)", "(?:９|9)", "(?:ａ|a)", "(?:ｂ|b)", "(?:ｃ|c)", "(?:ｄ|d)", "(?:ｅ|e)", "(?:ｆ|f)", "(?:ｇ|g)", "(?:ｈ|h)", "(?:ｉ|i)", "(?:ｊ|j)", "(?:ｋ|k)", "(?:ｌ|l)", "(?:ｍ|m)", "(?:ｎ|n)", "(?:ｏ|o)", "(?:ｐ|p)", "(?:ｑ|q)", "(?:ｒ|r)", "(?:ｓ|s)", "(?:ｔ|t)", "(?:ｕ|u)", "(?:ｖ|v)", "(?:ｗ|w)", "(?:ｘ|x)", "(?:ｙ|y)", "(?:ｚ|z)", "(?:ガ|ｶﾞ)", "(?:ギ|ｷﾞ)", "(?:グ|ｸﾞ)", "(?:ゲ|ｹﾞ)", "(?:ゴ|ｺﾞ)", "(?:ザ|ｻﾞ)", "(?:ジ|ｼﾞ)", "(?:ズ|ｽﾞ)", "(?:ゼ|ｾﾞ)", "(?:ゾ|ｿﾞ)", "(?:ダ|ﾀﾞ)", "(?:ヂ|ﾁﾞ)", "(?:ヅ|ﾂﾞ)", "(?:デ|ﾃﾞ)", "(?:ド|ﾄﾞ)", "(?:バ|ﾊﾞ)", "(?:ビ|ﾋﾞ)", "(?:ブ|ﾌﾞ)", "(?:ベ|ﾍﾞ)", "(?:ボ|ﾎﾞ)", "(?:パ|ﾊﾟ)", "(?:ピ|ﾋﾟ)", "(?:プ|ﾌﾟ)", "(?:ペ|ﾍﾟ)", "(?:ポ|ﾎﾟ)", "(?:。|｡)", "(?:「|｢)", "(?:」|｣)", "(?:、|､)", "(?:ヲ|ｦ)", "(?:ァ|ｧ)", "(?:ィ|ｨ)", "(?:ゥ|ｩ)", "(?:ェ|ｪ)", "(?:ォ|ｫ)", "(?:ャ|ｬ)", "(?:ュ|ｭ)", "(?:ョ|ｮ)", "(?:ッ|ｯ)", "(?:ー|ｰ)", "(?:ア|ｱ)", "(?:イ|ｲ)", "(?:ウ|ｳ)", "(?:エ|ｴ)", "(?:オ|ｵ)", "(?:カ|ｶ)", "(?:キ|ｷ)", "(?:ク|ｸ)", "(?:ケ|ｹ)", "(?:コ|ｺ)", "(?:サ|ｻ)", "(?:シ|ｼ)", "(?:ス|ｽ)", "(?:セ|ｾ)", "(?:ソ|ｿ)", "(?:タ|ﾀ)", "(?:チ|ﾁ)", "(?:ツ|ﾂ)", "(?:テ|ﾃ)", "(?:ト|ﾄ)", "(?:ナ|ﾅ)", "(?:ニ|ﾆ)", "(?:ヌ|ﾇ)", "(?:ネ|ﾈ)", "(?:ノ|ﾉ)", "(?:ハ|ﾊ)", "(?:ヒ|ﾋ)", "(?:フ|ﾌ)", "(?:ヘ|ﾍ)", "(?:ホ|ﾎ)", "(?:マ|ﾏ)", "(?:ミ|ﾐ)", "(?:ム|ﾑ)", "(?:メ|ﾒ)", "(?:モ|ﾓ)", "(?:ヤ|ﾔ)", "(?:ユ|ﾕ)", "(?:ヨ|ﾖ)", "(?:ラ|ﾗ)", "(?:リ|ﾘ)", "(?:ル|ﾙ)", "(?:レ|ﾚ)", "(?:ロ|ﾛ)", "(?:ワ|ﾜ)", "(?:ン|ﾝ)"];

    // 全角→半角の変換マップを作成
    var wideToNarrowMap = {};
    var narrowToHighlightMap = {};

    for (var i = 0; i < wide.length; i++) {
        wideToNarrowMap[wide[i]] = narrow[i];
        // 半角文字をキーにしてハイライトパターンをマップ
        narrowToHighlightMap[narrow[i]] = highlight[i];
    }

    // 全角文字の正規表現パターンを作成（エスケープ処理を含む）
    var wideCharsPattern = wide.map(function (char) {
        return char.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }).join('|');

    var wideToNarrowRegex = new RegExp(wideCharsPattern, 'g');

    return {
        wideToNarrowMap: wideToNarrowMap,
        narrowToHighlightMap: narrowToHighlightMap,
        wideToNarrowRegex: wideToNarrowRegex,
        // 効率的な変換関数
        convertWideToNarrow: function (text) {
            return text.replace(wideToNarrowRegex, function (match) {
                return wideToNarrowMap[match] || match;
            });
        }
    };
})();

// ========================================
// 基本ユーティリティ関数（依存なし）
// ========================================

// セレクタ用のエスケープ処理
function selectorEscape(val) {
    return val.replace(/[-\/\\^$*+?.()|[\]{}\!]/g, '\\$&');
}

// 文字列を正規化（全角→半角カナ変換、小文字化）
function normalizeForSearch(text) {
    var normalized = text.toLowerCase();
    return characterMappings.convertWideToNarrow(normalized);
}

// HTMLエンティティを復元
function decodeHtmlEntities(html) {
    return html
        .replace(htmlEntityRegexes.nbsp, "　")
        .replace(htmlEntityRegexes.gt, ">")
        .replace(htmlEntityRegexes.lt, "<")
        .replace(htmlEntityRegexes.quot, '"')
        .replace(htmlEntityRegexes.amp, "&");
}

// ========================================
// キャッシュ管理関数
// ========================================

// キャッシュを初期化
function initializeCachedElements() {
    for (var key in selectorMap) {
        if (selectorMap.hasOwnProperty(key)) {
            if (key === 'searchInput') {
                $cachedElements[key] = $(selectorMap[key], document);
            } else {
                $cachedElements[key] = $(selectorMap[key]);
            }
        }
    }
}

// キャッシュされた要素を取得（存在チェック付き）
function getCachedElement(key) {
    if (!$cachedElements[key] || $cachedElements[key].length === 0) {
        // キャッシュが無効な場合は再取得
        if (selectorMap.hasOwnProperty(key)) {
            if (key === 'searchInput') {
                $cachedElements[key] = $(selectorMap[key], document);
            } else {
                $cachedElements[key] = $(selectorMap[key]);
            }
        }
    }
    return $cachedElements[key];
}

// ========================================
// 検索処理関連関数（依存関係順）
// ========================================

// 検索語を正規化してエスケープ
function prepareSearchWords(searchValue) {
    // 全角・半角スペースを統一して連続スペースを1つにまとめる
    var searchWordTmp = searchValue.replace(/[　\s]+/g, " ").trim().toLowerCase();
    searchWordTmp = characterMappings.convertWideToNarrow(searchWordTmp);

    var searchWord = searchWordTmp.split(" ");
    for (var i = 0; i < searchWord.length; i++) {
        searchWord[i] = selectorEscape(searchWord[i].replace(/>/g, "&gt;").replace(/</g, "&lt;"));
    }
    return searchWord;
}

// ハイライト用の正規表現パターンを生成
function createHighlightPattern(searchWords) {
    // 各検索ワードを全角・半角両対応のパターンに変換
    var patterns = searchWords.map(function(word) {
        // 2文字パターンを優先的にマッチさせる正規表現を作成
        var result = '';
        var pos = 0;
        
        while (pos < word.length) {
            var matched = false;
            
            // 2文字の組み合わせを試行
            if (pos + 1 < word.length) {
                var twoChar = word.substring(pos, pos + 2);
                var pattern = characterMappings.narrowToHighlightMap[twoChar];
                if (pattern) {
                    result += pattern;
                    pos += 2;
                    matched = true;
                }
            }
            
            // 1文字で処理
            if (!matched) {
                var char = word.charAt(pos);
                var pattern = characterMappings.narrowToHighlightMap[char];
                result += pattern || char.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                pos++;
            }
        }
        
        return result;
    });
    return patterns.join("|");
}

// iframeのbody要素を取得
function getIframeBody() {
    var $iframe = getCachedElement('iframe');
    if ($iframe.length === 0) return null;
    return $iframe.contents().find("body");
}

// iframeコンテンツにハイライトを適用
function applyHighlight(searchValue) {
    var $body = getIframeBody();
    if (!$body) return;

    var searchWords = prepareSearchWords(searchValue);
    var highlightPattern = createHighlightPattern(searchWords);
    var reg = new RegExp("(" + highlightPattern + ")(?=[^<>]*<)", "gmi");
    var html = $body.html();
    var decodedHtml = decodeHtmlEntities(html);
    var highlightedHtml = decodedHtml.replace(reg, "<font class='keyword' style='color:rgb(0, 0, 0); background-color:rgb(252, 255, 0);'>$1</font>");

    $body.html(highlightedHtml);
}

// キーワードのハイライトを削除
function removeHighlight() {
    var $body = getIframeBody();
    if (!$body) return;

    $body.find(".keyword").each(function () {
        var $this = $(this);
        $this.replaceWith($this.contents());
    });
}

// 検索結果をクリア
function clearSearchResults() {
    var $searchResultItemsBlock = getCachedElement('searchResultItemsBlock');
    var $searchResultsEnd = getCachedElement('searchResultsEnd');
    var $searchMsg = getCachedElement('searchMsg');

    $searchResultItemsBlock.html("");
    $searchResultsEnd.addClass("rh-hide");
    $searchResultsEnd.attr("hidden", "");
    $searchMsg.html("2つ以上の語句を入力して検索する場合は、スペース（空白）で区切ります。");
    removeHighlight();
    currentSearchValue = "";
    disconnectMutationObserver();
}

// ========================================
// MutationObserver関連関数
// ========================================

// キーワード要素のみの追加かどうかを判定
function isOnlyKeywordAddition(mutation) {
    if (mutation.addedNodes.length !== 1) {
        return false;
    }
    
    var node = mutation.addedNodes[0];
    if (node.nodeType !== Node.ELEMENT_NODE) {
        return false;
    }
    
    // 追加されたノードがキーワード要素か、キーワード要素を含むか
    return (node.classList && node.classList.contains('keyword')) ||
           (node.querySelector && node.querySelector('.keyword') !== null);
}

// デバウンスされた再ハイライト処理
function debouncedReHighlight() {
    if (debounceTimer) {
        clearTimeout(debounceTimer);
    }

    debounceTimer = setTimeout(function () {
        reHighlightAfterDomChange();
    }, 500);
}

// DOM変更後の再ハイライト処理
function reHighlightAfterDomChange() {
    if (!currentSearchValue || currentSearchValue.trim() === "") {
        return;
    }

    try {
        disconnectMutationObserver();
        removeHighlight();
        applyHighlight(currentSearchValue);

        setTimeout(function () {
            setupMutationObserver();
        }, 100);

        console.debug("DOM変更後に検索語を再ハイライトしました");
    } catch (error) {
        console.warn("DOM変更後の再ハイライトに失敗しました:", error);
        setupMutationObserver();
    }
}

// iframeコンテンツ用のMutationObserverをセットアップ
function setupMutationObserver() {
    disconnectMutationObserver();

    try {
        var $iframe = getCachedElement('iframe');
        if ($iframe.length === 0) return;

        var iframeDocument = $iframe[0].contentDocument || $iframe[0].contentWindow.document;
        if (!iframeDocument || !iframeDocument.body) return;

        mutationObserver = new MutationObserver(function (mutations) {
            if (!currentSearchValue || currentSearchValue.trim() === "") {
                return;
            }

            var shouldReHighlight = mutations.some(function(mutation) {
                if (mutation.type === 'characterData') {
                    return true;
                }
                
                if (mutation.type === 'childList') {
                    // キーワード要素自身の追加は無視
                    if (isOnlyKeywordAddition(mutation)) {
                        return false;
                    }
                    
                    // ノードの追加または削除があれば再ハイライト
                    return mutation.addedNodes.length > 0 || mutation.removedNodes.length > 0;
                }
                
                return false;
            });

            if (shouldReHighlight) {
                debouncedReHighlight();
            }
        });

        mutationObserver.observe(iframeDocument.body, {
            childList: true,
            subtree: true,
            characterData: true,
            attributes: false
        });

        console.debug("iframeコンテンツ用のMutationObserverをセットアップしました");
    } catch (error) {
        console.warn("MutationObserverのセットアップに失敗しました:", error);
    }
}

// MutationObserverを切断
function disconnectMutationObserver() {
    if (mutationObserver) {
        mutationObserver.disconnect();
        mutationObserver = null;
    }

    if (debounceTimer) {
        clearTimeout(debounceTimer);
        debounceTimer = null;
    }
}

// ========================================
// jQueryカスタムセレクタ
// ========================================

// カスタムの:contains()セレクタ（正規化された検索用）
$.expr[':'].containsNormalized = function (elem, index, match) {
    var normalizedElemText = normalizeForSearch($(elem).text());
    var normalizedSearchText = normalizeForSearch(match[3]);
    return normalizedElemText.indexOf(normalizedSearchText) >= 0;
};

// ========================================
// jQueryイベントハンドラー
// ========================================

$(function () {
    // 要素のキャッシュを初期化
    initializeCachedElements();

    $(document).on("click", "ul.toc li.book", function () {
        if ($(this).children("a[href='#'],a[href='javascript:void 0;']").length == 0) {
            $(this).children("a").each(function () {
                location.href = location.href.replace(location.hash, "") + "#t=" + $(this).attr("href");
            });
        }
    });

    getCachedElement('searchField').each(function () {
        $(this).off();
    });

    $(document).on("keyup", ".wSearchField", function () {
        var $searchResultItemsBlock = getCachedElement('searchResultItemsBlock');
        var $searchResultsEnd = getCachedElement('searchResultsEnd');
        var $searchMsg = getCachedElement('searchMsg');
        var searchValue = $(this).val();

        // trim()を追加してスペースのみの入力も空として扱う
        if (searchValue.trim() === "") {
            clearSearchResults();
            return;
        }

        $searchMsg.html("");
        currentSearchValue = searchValue; // 現在の検索値を保存
        
        // スペースを正規化
        var searchWordTmp = searchValue.replace(/[　\s]+/g, " ").trim();
        // 正規化（全角→半角カナ、小文字化）
        searchWordTmp = normalizeForSearch(searchWordTmp);

        // 正規化後も空文字列チェックを追加
        if (searchWordTmp === "") {
            clearSearchResults();
            return;
        }

        var searchWord = searchWordTmp.split(" ");
        var searchQuery = searchWord.map(function(word) {
            return ":containsNormalized(" + word + ")";
        }).join("");

        var findItems = searchWords.find(".search_word" + searchQuery);
        if (findItems.length !== 0) {
            $searchResultsEnd.removeClass("rh-hide");
            $searchResultsEnd.removeAttr("hidden");
            $searchResultItemsBlock.html("");
            findItems.each(function () {
                var displayText = $(this).parent().find(".displayText").text();
                var parentId = $(this).parent().attr("id");
                var searchTitle = $(this).parent().find(".search_title").html();
                
                var resultHtml = "<div class='wSearchResultItem'>" +
                    "<a class='nolink' href='./" + parentId + ".html'>" +
                    "<div class='wSearchResultTitle'>" + searchTitle + "</div>" +
                    "</a>" +
                    "<div class='wSearchContent'>" +
                    "<span class='wSearchContext'>" + displayText + "</span>" +
                    "</div></div>";
                
                $searchResultItemsBlock.append($(resultHtml));
            });
            removeHighlight();
            applyHighlight(currentSearchValue);
        }
        else {
            removeHighlight();
            $searchResultsEnd.addClass("rh-hide");
            $searchResultsEnd.attr("hidden", "");
            $searchResultItemsBlock.html("");
            var displayText = "検索条件に一致するトピックはありません。";
            $searchResultItemsBlock.append($("<div class='wSearchResultItem'><div class='wSearchContent'><span class='wSearchContext'>" + displayText + "</span></div></div>"));
        }
    });

    getCachedElement('iframe').on("load", function () {
        var $searchInput = getCachedElement('searchInput');
        var $searchField = getCachedElement('searchField');

        if ($searchInput.is(":not(.rh-hide)") && ($searchField.val() != "")) {
            var searchValue = $searchField.val();
            currentSearchValue = searchValue; // 現在の検索値を保存
            applyHighlight(searchValue);
        }

        // iframeコンテンツ変更用のMutationObserverをセットアップ
        setupMutationObserver();
    });
});
