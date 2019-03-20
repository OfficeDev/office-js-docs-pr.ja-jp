---
title: Excel JavaScript API の要件セット
description: ''
ms.date: 03/19/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: e7d9eac6d06fdce8936e92a001ff213b04a50bc1
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691231"
---
# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

Excel アドインは、Office for Windows 2016 以降、Office for iPad、Office for Mac、Office Online など、複数のバージョンの Office で機能します。 次の表に、Excel の要件セット、各要件セットをサポートする Office ホスト アプリケーション、それらのアプリケーションのビルド バージョンまたはビルド番号を記載します。

> [!NOTE]
> 番号付きの要件セットで API を使用するには、CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js で**実稼働**ライブラリを参照してください。
>
> プレビューの API の使用に関する詳細については、この記事の「[Excel JavaScript プレビュー API](#excel-javascript-preview-apis)」セクションを参照してください。

|  要件セット  |  Office 365 for Windows  |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| プレビュー  | プレビュー API を試すには、最新版 Office を使用してください (場合によっては、[Office Insider プログラム](https://products.office.com/office-insider)に参加する必要があります) |
| ExcelApi1.8  | バージョン 1808 (ビルド 10730.20102) 以降 | 2.17 以降 | 16.17 以降 | 2018 年 9 月 | 間もなく提供開始 |
| ExcelApi1.7  | バージョン 1801 (ビルド 9001.2171) 以降   | 2.9 以降 | 16.9 以降 | 2018 年 4 月 | 間もなく提供開始 |
| ExcelApi1.6  | バージョン 1704 (ビルド 8201.2001) 以降   | 2.2 以降 |15.36 以降| 2017 年 4 月 | 近日公開|
| ExcelApi1.5  | バージョン 1703 (ビルド 8067.2070) 以降   | 2.2 以降 |15.36 以降| 2017 年 3 月 | 近日公開|
| ExcelApi1.4  | バージョン 1701 (ビルド 7870.2024) 以降   | 2.2 以降 |15.36 以降| 2017 年 1 月 | 近日公開|
| ExcelApi1.3  | バージョン 1608 (ビルド 7369.2055) 以降 | 1.27 以降 |  15.27 以降| 2016 年 9 月 | バージョン 1608 (ビルド 7601.6800) 以降|
| ExcelApi1.2  | バージョン 1601 (ビルド 6741.2088) 以降 | 1.21 以降 | 15.22 以降| 2016 年 1 月 ||
| ExcelApi1.1  | バージョン 1509 (ビルド 4266.1001) 以降 | 1.19 以降 | 15.20 以降| 2016 年 1 月 ||

> [!NOTE]
> MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。 このバージョンには、ExcelApi 1.1 の要件セットのみが含まれています。

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。 次の表は、現在プレビューで利用できる API の一覧です。 プレビュー API についてフィードバックを提供するには、その API が記載されている Web ページの最後にあるフィードバック メカニズムを使用してください。

> [!NOTE]
> プレビュー API は変更されることがあります。運用環境での使用は意図されていません。 試用はテスト環境と開発環境に限定することをお勧めします。 運用環境やビジネス上重要なドキュメントでプレビュー API を使用しないでください。
>
> プレビュー API を使用するには、CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js で**ベータ** ライブラリを参照する必要があります。場合によっては、Office Insider プログラムに参加し、十分に新しい Office ビルドを入手する必要があります。

400 以上の新しい Excel API が現在プレビュー中です。 最初の表には API が簡潔にまとめられています。後続の表は詳しい一覧になっています。 新しい機能を試し、フィードバックを投稿してください。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| スライサー | テーブルやピボットテーブルにスライサーを挿入し、構成します。 | [Slicer](/javascript/api/excel/excel.slicer) |
| コメント | コメントを追加、編集、削除します。 | [Comment](/javascript/api/excel/excel.comment)、[CommentCollection](/javascript/api/excel/excel.commentcollection) |
| 図形 | 画像、幾何学な図形、テキスト ボックスを挿入、位置変更、書式設定します。 | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape) [Image](/javascript/api/excel/excel.image) |
| 新しいグラフ | 新しくサポートされたグラフである、マップ、箱ひげ図、ウォーターフォール、サンバースト、パレート、 じょうごをお試しください。 | [Chart](/javascript/api/excel/excel.charttype) |
| オート フィルター | 範囲にフィルターを追加します。 | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| エリア | 連続していない範囲をサポートします。 | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| 特別なセル | ある範囲内に日付、コメント、数式を含むセルを取得します。 | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| 検索 | ある範囲またはワークシート内で値や数式を見つけます。 | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| コピーと貼り付け | 範囲間で値、書式、数式をコピーします。 | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| 範囲の形式 | 範囲の形式の新しい機能です。 | [Range](/javascript/api/excel/excel.rangeformat) |
| ブックを保存して閉じる | ブックを保存して閉じます。  | [Workbook](/javascript/api/excel/excel.workbook) |
| ブックを挿入する | あるブックを別のブックに挿入します。  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| 計算 | Excel 計算エンジンを細かく操作できます。 | [Application](/javascript/api/excel/excel.application) |

以下は、プレビュー中の API の完全な一覧です。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|最後にブックを完全に再計算した Excel 計算エンジンのバージョンに関する数字を返します。 読み取り専用です。|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|アプリケーションの計算状態を示す CalculationState を返します。 詳細については、Excel.CalculationState をご覧ください。 読み取り専用です。|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|反復計算の設定を返します。|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|次の "context.sync()" が呼び出されるまで画面の更新を一時停止します。|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|ある範囲に AutoFilter を適用し、列インデックスやフィルター条件が指定されている場合、列にフィルターを適用します。|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|AutoFilter にフィルターが含まれている場合、条件を消去します。|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|AutoFilter が適用される範囲を表す Range オブジェクトを返します。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Range オブジェクトが AutoFilter に関連付けられている場合、このメソッドによってそのオブジェクトが返されます。|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|自動的にフィルターが適用された範囲のすべてのフィルター条件を保持する配列です。 読み取り専用です。|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|AutoFilter が有効かどうかを示します。 読み取り専用です。|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|AutoFilter にフィルター条件が与えられているかどうかを示します。 読み取り専用です。|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|その範囲で現在指定されている Autofilter オブジェクトを適用します。|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|範囲の AutoFilter を削除します。|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)||
||[style](/javascript/api/excel/excel.cellborder#style)||
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)||
||[weight](/javascript/api/excel/excel.cellborder#weight)||
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)||
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)||
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)||
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)||
||[left](/javascript/api/excel/excel.cellbordercollection#left)||
||[right](/javascript/api/excel/excel.cellbordercollection#right)||
||[top](/javascript/api/excel/excel.cellbordercollection#top)||
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)||
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)||
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)||
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)||
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)||
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)||
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)||
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)||
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)||
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)||
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)||
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)||
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)||
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)||
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)||
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)||
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)||
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)||
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)||
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)||
||[borders](/javascript/api/excel/excel.cellpropertiesformat#borders)||
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)||
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)||
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)||
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)||
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)||
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)||
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)||
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)||
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)||
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)||
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)||
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|新しいブックを作成し、開きます。  任意で、base64 でエンコードした .xlsx ファイルでブックにデータを事前入力できます。|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)||
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)||
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|変更後の値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|変更前の値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|変更後の値の型を表します。|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|変更前の値の型を表します。|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Excel UI でグラフをアクティブにします。|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|ピボット グラフのオプションをカプセル化します。 読み取り専用です。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|グラフの配色を表す整数型の値を設定または返します。 読み取り/書き込み可能。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|グラフのグラフ エリアの角が丸くなる場合、true となります。 読み取り/書き込み可能。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表します。|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|ヒストグラム図やパレート図でビンのオーバーフローが有効になっているかどうかを設定または返します。 読み取り/書き込み可能。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|ヒストグラム図やパレート図でビンのアンダーフローが有効になっているかどうかを設定または返します。 読み取り/書き込み可能。|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|ヒストグラム図やパレート図のビン数を設定または返します。 読み取り/書き込み可能。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|ヒストグラム図やパレート図のビンのオーバーフロー値を設定または返します。 読み取り/書き込み可能。|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|ヒストグラム図やパレート図のビンの種類を設定または返します。 読み取り/書き込み可能。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|ヒストグラム図やパレート図のビンのアンダーフロー値を設定または返します。 読み取り/書き込み可能。|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|ヒストグラム図やパレート図のビンの幅値を設定または返します。 読み取り/書き込み可能。|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|箱ひげ図の四分位数計算の種類を設定または返します。 読み取り/書き込み可能。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|箱ひげ図に内側のポイントが表示されているかどうかを設定または返します。 読み取り/書き込み可能。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|箱ひげ図に平均線が表示されているかどうかを設定または返します。 読み取り/書き込み可能。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|箱ひげ図に平均マーカーが表示されているかどうかを設定または返します。 読み取り/書き込み可能。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|箱ひげ図に特異ポイントが表示されているかどうかを設定または返します。 読み取り/書き込み可能。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表します。|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|誤差範囲に終点のスタイルのキャップがあるかどうかを表します。|
||[include](/javascript/api/excel/excel.charterrorbars#include)|含める誤差範囲の部分を表します。 詳細については、Excel.ChartErrorBarsInclude をご覧ください。|
||[format](/javascript/api/excel/excel.charterrorbars#format)|グラフの誤差範囲の書式設定を表します。|
||[type](/javascript/api/excel/excel.charterrorbars#type)|誤差範囲でマークされる範囲を表します。 詳細については、Excel.ChartErrorBarsType をご覧ください。|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|誤差範囲が表示されているかどうかを表します。|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|グラフの線の書式設定を表します。|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|リージョン マップ グラフの系列マップ ラベル方法を設定または返します。 読み取り/書き込み可能。|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|リージョン マップ グラフの系列マップ領域を設定または返します。 読み取り/書き込み可能。|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|リージョン マップ グラフの系列投影タイプを設定または返します。 読み取り/書き込み可能。|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|ピボットグラフに軸フィールド ボタンを表示するかどうかを表します。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|ピボットグラフに凡例フィールド ボタンを表示するかどうかを表します。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|ピボットグラフにレポート フィルター フィールド ボタンを表示するかどうかを表します。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|ピボットグラフに値表示フィールド ボタンを表示するかどうかを表します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|指定されたグラフ グループ内のバブルのスケール ファクターを設定または返します。 既定のサイズのパーセンテージに対応する 0 (ゼロ) から 300 までの整数値とすることができます。 バブル チャートにのみ適用されます。 読み取り/書き込み可能。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|リージョン マップ グラフ系列の最大値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|リージョン マップ グラフ系列の最大値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|リージョン マップ グラフ系列の最大値を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|リージョン マップ グラフ系列の中間値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|リージョン マップ グラフ系列の中間値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|リージョン マップ グラフ系列の中間値を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|リージョン マップ グラフ系列の最小値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|リージョン マップ グラフ系列の最小値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|リージョン マップ グラフ系列の最小値を設定または返します。 読み取り/書き込み可能。|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|リージョン マップ グラフのグラデーション スタイルを設定または返します。 読み取り/書き込み可能。|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|系列の負のデータ ポイントに対して塗りつぶしの色を設定または返します。 読み取り/書き込み可能。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|ツリーマップ グラフの系列上位ラベル方法領域を設定または返します。 読み取り/書き込み可能。|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|ヒストグラム図とパレート図にのみ、ビンのオプションをカプセル化します。 読み取り専用です。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|箱ひげ図グラフのオプションをカプセル化します。 読み取り専用です。|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|マップ グラフのオプションをカプセル化します。 読み取り専用です。|
||[xerrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
||[yerrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|ウォーターフォール図にコネクタの線が表示されているかどうかを設定または返します。 読み取り/書き込み可能。|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|Microsoft Excel で系列のデータラベルごとに引出線が表示されている場合、true となります。 読み取り/書き込み可能。|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|補助のグラフ (円または縦棒) が付いた円グラフを 2 つの部分に区切るしきい値を設定または返します。 読み取り/書き込み可能。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)||
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)||
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)||
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|コンテンツを取得または設定します。|
||[delete()](/javascript/api/excel/excel.comment#delete--)|コメント スレッドを削除します。|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|コメントの場所を取得します。|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|コメントの作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.comment#authorname)|コメントの作成者の名前を取得します。|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|コメントの作成日時を取得します。 コメントがメモから変換された場合は、null が返されます。この場合、コメントに作成日時はありません。|
||[id](/javascript/api/excel/excel.comment#id)|コメント ID を表します。 読み取り専用です。|
||[isParent](/javascript/api/excel/excel.comment#isparent)|コメント スレッドまたは返信かどうかを表します。 ここでは常に true を返します。 読み取り専用です。|
||[replies](/javascript/api/excel/excel.comment#replies)|コメントに関連付けられている返信オブジェクトのコレクションを表します。 読み取り専用です。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: "Plain")](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|セルの一とコンテンツに基づいて新しいコメント (コメント スレッド) を作成します。 場所が 1 つのセルより大きい場合、無効な引数がスローされます。|
||[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|セルの一とコンテンツに基づいて新しいコメント (コメント スレッド) を作成します。 場所が 1 つのセルより大きい場合、無効な引数がスローされます。|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|コレクションに含まれるコメントの数を取得します。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|その ID で識別されるコメントを返します。 読み取り専用です。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|コレクション内の位置に基づいてコメントを取得します。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|コレクション内の特定のセルに関するコメントを取得します。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|コレクション内のその返信 ID に関連付けられているコメントを取得します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.commentcollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.commentcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|コンテンツを取得または設定します。|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|コメント返信を削除します。|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|コメント返信の場所を取得します。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|この返信の親コメントを取得します。|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|コメント返信の作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|コメント返信の作成者の名前を取得します。|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|コメント返信の作成日時を取得します。|
||[id](/javascript/api/excel/excel.commentreply#id)|コメント返信 ID を表します。 読み取り専用です。|
||[isParent](/javascript/api/excel/excel.commentreply#isparent)|コメント スレッドまたは返信かどうかを表します。 ここでは常に false を返します。 読み取り専用です。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: "Plain")](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|コメントのコメント返信を作成します。|
||[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|コメントのコメント返信を作成します。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|コレクションのコメント返信数を取得します。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|その ID で識別されるコメント返信を返します。 読み取り専用です。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|コレクション内の位置に基づいてコメント返信を取得します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.commentreplycollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|1 つまたは複数の長方形範囲で構成され、条件付き書式が適用された RangeAreas を返します。 読み取り専用です。|
|[CustomFunctionEventArgs](/javascript/api/excel/excel.customfunctioneventargs)|[higherTicks](/javascript/api/excel/excel.customfunctioneventargs#higherticks)||
||[lowerTicks](/javascript/api/excel/excel.customfunctioneventargs#lowerticks)||
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|1 つまたは複数の長方形範囲で構成され、無効なセル値を含む RangeAreas を返します。 すべてのセル値が有効な場合、この関数からは ItemNotFound エラーがスローされます。|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|1 つまたは複数の長方形範囲で構成され、無効なセル値を含む RangeAreas を返します。 すべてのセル値が有効な場合、この関数からは null が返されます。|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|リッチな値にリッチなフィルターを適用する場合、フィルターによって使用されるプロパティです。|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|図形 ID を返します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|幾何学的図形の Shape オブジェクトを返します。 読み取り専用です。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|図形グループの図形の数を返します。 読み取り専用です。|
||[getItem(name: string)](/javascript/api/excel/excel.groupshapecollection#getitem-name-)|名前を使用して、その図形を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|コレクション内の位置に基づいて図形を取得します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.groupshapecollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|ワークシートの中央フッターを取得または設定します。|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|ワークシートの中央ヘッダーを取得または設定します。|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|ワークシートの左フッターを取得または設定します。|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|ワークシートの左ヘッダーを取得または設定します。|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|ワークシートの右フッターを取得または設定します。|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|ワークシートの右ヘッダーを取得または設定します。|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|偶数/奇数または最初のページが指定されていない場合にすべてのページに使用される汎用ヘッダー/フッター。|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|偶数ページに使用するヘッダー/フッター。奇数ページには奇数のヘッダー/フッターを指定する必要があります。|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|最初のページに使用するヘッダー/フッター。その他すべてのページには汎用または偶数/奇数のヘッダー/フッターが使用されます。|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|奇数ページに使用するヘッダー/フッター。偶数ページには偶数のヘッダー/フッターを指定する必要があります。|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|設定されているヘッダー/フッターの状態を取得または設定します。 詳細については、Excel.HeaderFooterState をご覧ください。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|ワークシートのページ レイアウト オプションに設定されているページ余白に合わせてヘッダー/フッターの位置が調整されているかどうかを示すフラグを取得または設定します。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|ワークシートのページ レイアウト オプションに設定されているページ パーセンテージ スケールによってヘッダー/フッターが調整されているかどうかを示すフラグを取得または設定します。|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|画像の形式を返します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.image#id)|画像オブジェクトの図形 ID を表します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.image#shape)|画像に関連付けられた Shape オブジェクトを返します。 読み取り専用です。|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|Excel で反復計算を使用して循環参照を解決する場合、true となります。|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|循環参照は Excel の反復計算によって解決されます。その反復計算間の変化の最大値を設定または返します。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Excel で循環参照の解決に使用できる、最大反復回数を設定または返します。|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|指定された線の始点の矢印の長さを表します。|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|指定された線の始点の矢印のスタイルを表します。|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|指定された線の始点の矢印の幅を表します。|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|指定されたコネクタの始点を指定された図形に接続します。|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|指定されたコネクタの終点を指定された図形に接続します。|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|線のコネクタの種類を表します。|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|指定されたコネクタの始点を図形から切り離します。|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|指定されたコネクタの終点を図形から切り離します。|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|指定された線の終点の矢印の長さを表します。|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|指定された線の終点の矢印のスタイルを表します。|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|指定された線の終点の矢印の幅を表します。|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|指定された線の始点が接続されている図形を表します。 読み取り専用です。|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|コネクタの始点が接続されている結合点を表します。 読み取り専用です。 線の始点がどの図形にも接続されていない場合は、null を返します。|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|指定された線の終点が接続されている図形を表します。 読み取り専用です。|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|コネクタの終点が接続されている結合点を表します。 読み取り専用です。 線の終点がどの図形にも接続されていない場合は、null を返します。|
||[id](/javascript/api/excel/excel.line#id)|図形 ID を表します。 読み取り専用です。|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|指定された線の始点が図形に接続されているかどうかを指定します。 読み取り専用です。|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|指定された線の終点が図形に接続されているかどうかを指定します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.line#shape)|線に関連付けられた Shape オブジェクトを返します。 読み取り専用です。|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[source](/javascript/api/excel/excel.listdatavalidation#source)|データの入力規則のリストのソース。|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|改ページ オブジェクトを削除します。|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|改ページの後の最初のセルを取得します。|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|改ページの列インデックスを表します。|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|改ページの行インデックスを表します。|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|指定された範囲の左上セルの前に改ページを追加します。|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|コレクション内の改ページの数を取得します。|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|インデックス経由で改ページ オブジェクトを取得します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.pagebreakcollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|コレクション内の手動改ページをすべてリセットします。|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|ワークシートの白黒印刷オプションを取得または設定します。|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|ポイント単位印刷に使用するワークシートの下部ページ余白を取得または設定します。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|ワークシートの [ページ中央] の [水平] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を水平に設定するかどうかが決定されます。|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|ワークシートの [ページ中央] の [垂直] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を垂直に設定するかどうかが決定されます。|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|ワークシートの下書きモード オプションを取得または設定します。 true の場合、グラフィックスなしでシートが印刷されます。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|印刷するワークシートの最初のページ番号を取得または設定します。 null 値は "自動" ページ番号を表します。|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|印刷時に使用するワークシートのフッター余白 (ポイント数) を取得または設定します。|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|ワークシートの印刷範囲を表し、1 つまたは複数の長方形範囲で構成される RangeAreas オブジェクトを取得します。 印刷範囲がない場合、ItemNotFound エラーがスローされます。|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|ワークシートの印刷範囲を表し、1 つまたは複数の長方形範囲で構成される RangeAreas オブジェクトを取得します。 印刷範囲がない場合、null オブジェクトが返されます。|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|タイトル列を表す範囲オブジェクトを取得します。|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|タイトル列を表す範囲オブジェクトを取得します。 設定されていない場合、null オブジェクトが返されます。|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|タイトル行を表す範囲オブジェクトを取得します。|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|タイトル行を表す範囲オブジェクトを取得します。 設定されていない場合、null オブジェクトが返されます。|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|印刷時に使用するワークシートのヘッダー余白 (ポイント数) を取得または設定します。|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|印刷時に使用するワークシートの左余白 (ポイント数) を取得または設定します。|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|ワークシートのページの向きを取得または設定します。|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|ワークシートのページの用紙サイズを取得または設定します。|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|印刷時、ワークシートのコメントを表示するかどうかを取得または設定します。|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|ワークシートの印刷エラー オプションを取得または設定します。|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|ワークシートの印刷目盛線フラグを取得または設定します。 このフラグによって、目盛線を印刷するかどうかが決定されます。|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|ワークシートの見出し印刷フラグを取得または設定します。 このフラグによって、見出しを印刷するかどうかが決定されます。|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|ワークシートのページ印刷順序オプションを取得または設定します。 これによって、印刷されるページ番号の処理に使用する順序が指定されます。|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|ワークシートのヘッダーとフッターの構成。|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|印刷時に使用するワークシートの右余白 (ポイント数) を取得または設定します。|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|ワークシートの印刷範囲を設定します。|
||[setPrintMargins(unit: "Points" \| "Inches" \| "Centimeters", marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|ワークシートのページ余白を単位で設定します。|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|ワークシートのページ余白を単位で設定します。|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|セルを含む列を、印刷時、ワークシートの各ページの左で繰り返すように設定します。|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|セルを含む行を、印刷時、ワークシートの各ページの上で繰り返すように設定します。|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|印刷時に使用するワークシートの上余白 (ポイント数) を取得または設定します。|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|ワークシートの拡大印刷オプションを取得または設定します。|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|印刷時に使用するように指定された単位でページ レイアウトの下余白を表します。|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|印刷時に使用するように指定された単位でページ レイアウトのフッター余白を表します。|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|印刷時に使用するように指定された単位でページ レイアウトのヘッダー余白を表します。|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|印刷時に使用するように指定された単位でページ レイアウトの左余白を表します。|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|印刷時に使用するように指定された単位でページ レイアウトの右余白を表します。|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|印刷時に使用するように指定された単位でページ レイアウトの上余白を表します。|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|横方向に合わせるページ数。 パーセンテージ スケールが使用される場合、この値には null を指定できます。|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|印刷ページのスケール値は 10 から 400 までです。 縦または横方向にページを合わせるように指定されている場合、この値には null を指定できます。|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|縦方向に合わせるページ数。 パーセンテージ スケールが使用される場合、この値には null を指定できます。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortby: "Ascending" \| "Descending", valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|与えられた範囲で、指定された値に基づいて PivotField を並べ替えます。 この範囲によって、並べ替えに使用する特定の値が定義されます。|
||[sortByValues(sortby: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|与えられた範囲で、指定された値に基づいて PivotField を並べ替えます。 この範囲によって、並べ替えに使用する特定の値が定義されます。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|更新時またはフィールドの削除時に書式が自動的に設定される場合、true となります。|
||[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|UI のフィールド リストを表示するか、非表示にする場合、true となります。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|指定された dataHierarchy、rowItems、columnItems の交差の値を含む、PivotTable のデータ本体のセルを取得します。|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|PivotTable 内で指定された範囲の値を計算するために使用される DataHierarchy を取得します。|
||[getPivotItems(axis: "Unknown" \| "Row" \| "Column" \| "Data" \| "Filter", cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|PivotTable 内で指定された範囲の値を構成する PivotItems を軸から取得します。|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|PivotTable 内で指定された範囲の値を構成する PivotItems を軸から取得します。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|ピボット、並べ替え、ページ フィールド項目の変更などの操作によってレポートが更新または再計算されたとき、書式設定が保存される場合、true となります。|
||[setAutosortOnCell(cell: Range \| string, sortby: "Ascending" \| "Descending")](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|並べ替えにすべての条件とコンテキストを自動選択するよう、指定されたセルを使用し、自動並べ替えを設定します。|
||[setAutosortOnCell(cell: Range \| string, sortby: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|並べ替えにすべての条件とコンテキストを自動選択するよう、指定されたセルを使用し、自動並べ替えを設定します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|並べ替え時、PivotTable でカスタム リストを使用する場合、true となります。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|並べ替え時、PivotTable でカスタム リストを使用する場合、true となります。|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|PivotTableStyle を削除します。|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|すべてのスタイル要素のコピーでこの PivotTableStyle の複製を作成します。|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|PivotTableStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|true は、この PivotTableStyle オブジェクトが読み取り専用であることを意味します。 読み取り専用です。|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|指定された名前で空の PivotTableStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|コレクションに含まれる PivotTableStyle の数を取得します。|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|親オブジェクトのスコープに対する既定の PivotTableStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|名前に基づいて PivotTableStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|名前に基づいて PivotTableStyle を取得します。 PivotTableStyle が存在しない場合は、null オブジェクトを返します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.pivottablestylecollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の PivotTableStyle を設定します。|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: "FillDefault" \| "FillCopy" \| "FillSeries" \| "FillFormats" \| "FillValues" \| "FillDays" \| "FillWeekdays" \| "FillMonths" \| "FillYears" \| "LinearTrend" \| "GrowthTrend" \| "FlashFill")](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|現在の範囲から対象の範囲までの範囲に値を設定します。|
||[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|現在の範囲から対象の範囲までの範囲に値を設定します。|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|データ型を含む範囲セルをテキストに変換します。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|ワークシート内で範囲セルをリンク付きデータ型に変換します。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|ソース範囲または RangeAreas から現在の範囲にセル データまたは書式設定をコピーします。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|ソース範囲または RangeAreas から現在の範囲にセル データまたは書式設定をコピーします。|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|指定された条件に基づいて指定された文字列を見つけます。|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|指定された条件に基づいて指定された文字列を見つけます。|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|現在の範囲に対してフラッシュ フィルを実行します。フラッシュ フィルでは、パターンを感知して自動的にデータが設定されるので、範囲は単一列範囲で、かつパターンを検出できるように周囲にデータが存在する必要があります。|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|2D 配列を返します。各セルのフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|一次元配列を返します。各列のフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。  指定された列内の列間で一貫性のないプロパティについては、null が返されます。|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|一次元配列を返します。各行のフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。  指定された行内の列間で一貫性のないプロパティについては、null が返されます。|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表し、1 つまたは複数の長方形範囲で構成される RangeAreas オブジェクトを取得します。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表し、1 つまたは複数の長方形範囲で構成される RangeAreas オブジェクトを取得します。|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表し、1 つまたは複数の範囲で構成される、RangeAreas オブジェクトを取得します。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表し、1 つまたは複数の範囲を構成する RangeAreas オブジェクトを取得します。|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。 読み取り専用です。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 読み取り専用です。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。 読み取り専用です。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 読み取り専用です。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|範囲と重なるテーブルの集まりを範囲限定で取得します。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|各セルのデータ型の状態を表します。 読み取り専用です。|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|列によって指定される範囲から重複する値を削除します。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|現在の範囲内で、指定された条件に基づき、指定された文字列を検索し、置換します。|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][] \| OfficeExtension.ClientResult<SettableCellProperties[][]>)](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|セル プロパティの 2D 配列に基づいて範囲を更新します。フォント、塗りつぶし、罫線、配置などをカプセル化します。|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[] \| OfficeExtension.ClientResult<SettableColumnProperties[]>)](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|列プロパティの一次元配列に基づいて範囲を更新します。フォント、塗りつぶし、罫線、配置などをカプセル化します。|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|次の再計算が発生したときに再計算する範囲を設定します。|
||[setRowProperties(rowPropertiesData: SettableRowProperties[] \| OfficeExtension.ClientResult<SettableRowProperties[]>)](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|行プロパティの一次元配列に基づいて範囲を更新します。フォント、塗りつぶし、罫線、配置などをカプセル化します。|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|RangeAreas のすべてのセルを計算します。|
||[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.rangeareas#clear-applyto-)|この RangeAreas オブジェクトを構成する各領域で値、フォーマット、塗りつぶし、罫線などを消去します。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|この RangeAreas オブジェクトを構成する各領域で値、フォーマット、塗りつぶし、罫線などを消去します。|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|RangeAreas 内でデータ型を含むすべてのセルをテキストに変換します。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|RangeAreas 内のすべてのセルをリンク付きデータ型に変換します。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|ソース範囲または RangeAreas から現在の RangeAreas にセル データまたは書式設定をコピーします。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|ソース範囲または RangeAreas から現在の RangeAreas にセル データまたは書式設定をコピーします。|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|RangeAreas の列全体を表す RangeAreas オブジェクトを返します (たとえば、現在の RangeAreas がセル "B4:E11, H2" を表す場合、列 "B:E, H:H" を表す RangeAreas が返されます)。|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|RangeAreas の行全体を表す RangeAreas オブジェクトを返します (たとえば、現在の RangeAreas がセル "B4:E11" を表す場合、行 "4:11" を表す RangeAreas が返されます)。|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|指定した範囲または RangeAreas の交差を表す RangeAreas オブジェクトを返します。 交差が見つからない場合、ItemNotFound エラーがスローされます。|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|指定した範囲または RangeAreas の交差を表す RangeAreas オブジェクトを返します。 交差が見つからない場合、null オブジェクトが返されます。|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|特定の行と列のオフセットによってシフトされる RangeAreas オブジェクトを返します。 返される RangeAreas のディメンションは元のオブジェクトと一致します。 結果の RangeAreas がワークシート グリッドの境界線の外にはみ出る場合、エラーがスローされます。|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表す RangeAreas オブジェクトを返します。 条件に一致する特別なセルが見つからない場合、エラーがスローされます。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表す RangeAreas オブジェクトを返します。 条件に一致する特別なセルが見つからない場合、エラーがスローされます。|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表す RangeAreas オブジェクトを返します。 条件に一致する特別なセルが見つからない場合、null オブジェクトを返します。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表す RangeAreas オブジェクトを返します。 条件に一致する特別なセルが見つからない場合、null オブジェクトを返します。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|この RangeAreas オブジェクトの範囲と重なるテーブルの集まりを範囲限定で返します。|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|RangeAreas オブジェクトの個別の長方形範囲の全使用済み領域を構成する使用済み RangeAreas を返します。|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|RangeAreas オブジェクトの個別の長方形範囲の全使用済み領域を構成する使用済み RangeAreas を返します。|
||[address](/javascript/api/excel/excel.rangeareas#address)|A1 スタイルで RageAreas 参照を返します。 address 値にはセルの各長方形ブロックのワークシート名が含まれます ("Sheet1!A1:B4, Sheet1!D1:D4" など)。 読み取り専用です。|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|ユーザー ロケールで RageAreas 参照を返します。 読み取り専用です。|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|この RangeAreas オブジェクトを構成する長方形範囲の数を返します。|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|この RangeAreas オブジェクトを構成する長方形範囲の集まりを返します。|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|RangeAreas オブジェクトのセル数を返します。すべての個別長方形範囲のセル数が合計されます。 セル数が 2^31-1 (2,147,483,647) を超える場合、-1 を返します。 読み取り専用です。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|この RangeAreas オブジェクトのセルと交差する ConditionalFormats の集まりを返します。 読み取り専用です。|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|RangeAreas の全範囲に対して dataValidation オブジェクトを返します。|
||[format](/javascript/api/excel/excel.rangeareas#format)|rangeFormat オブジェクトを返します。RangeAreas オブジェクトの全範囲を対象にフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化します。 読み取り専用です。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|この RangeAreas オブジェクトの全範囲が列全体を表すかどうかを示します ("A:C, Q:Z" など)。 読み取り専用です。|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|この RangeAreas オブジェクトの全範囲が行全体を表すかどうかを示します ("1:3, 5:7" など)。 読み取り専用です。|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|現在の RangeAreas のワークシートを返します。 読み取り専用です。|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|次の再計算が発生したときに再計算する RangeAreas を設定します。|
||[style](/javascript/api/excel/excel.rangeareas#style)|この RangeAreas オブジェクトの全範囲のスタイルを表します。|
||[track()](/javascript/api/excel/excel.rangeareas#track--)|ドキュメントの環境変更に基づいて自動的に調整する目的でオブジェクトを追跡します。 これは context.trackedObjects.add(thisObject) 呼び出しの省略形です。 ".sync" 呼び出し間で、かつ ".run" バッチの連続実行の外でこのオブジェクトを使用しているとき、オブジェクトであるプロパティを設定したか、あるメソッドを呼び出したときに "InvalidObjectPath" エラーが表示される場合、オブジェクトを最初に作成したときに、追跡対象オブジェクトの集まりにそのオブジェクトを追加しておく必要がありました。|
||[untrack()](/javascript/api/excel/excel.rangeareas#untrack--)|前に追跡されていた場合、このオブジェクトに関連付けられているメモリを解放します。 これは context.trackedObjects.remove(thisObject) 呼び出しの省略形です。 追跡対象オブジェクトが多いとホスト アプリケーションの動作が遅くなります。追加したオブジェクトが不要になったら、必ずそれを解放してください。 メモリ リリースを有効にするには、"context.sync()" を先に呼び出す必要があります。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|範囲の境界線の色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|範囲の境界線の色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|RangeCollection 内の範囲数を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|RangeCollection 内のその位置に基づいて範囲オブジェクトを返します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.rangecollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.rangecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|範囲のパターンを取得または設定します。 詳細については、Excel.FillPattern をご覧ください。 LinearGradient と RectangularGradient はサポートされていません。|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Range パターンの色を表す HTML カラー コードを設定します。形式は #RRGGBB (例: "FFA500") か名前付き HTML 色 (例: "オレンジ") です。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|範囲の塗りつぶしのパターン色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|範囲の塗りつぶしの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|フォントの取り消し線の状態を表します。 null 値は、範囲全体に同じ取り消し線設定がないことを示します。|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|フォントの下付きの状態を表します。|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|フォントの上付きの状態を表します。|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|範囲のフォントの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|テキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|インデント レベルを示す 0 から 250 までの整数。|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|範囲に適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|操作によって削除された重複行の数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|結果として生じた範囲に存在する残りの一意の行の数。|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|一致方法として完全一致か部分一致を指定します。 既定値は false (部分一致) です。|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|照合の際に大文字と小文字を区別するかどうかを指定します。 既定値は false (区別しない) です。|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)||
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)||
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)||
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|一致方法として完全一致か部分一致を指定します。 既定値は false (部分一致) です。|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|照合の際に大文字と小文字を区別するかどうかを指定します。 既定値は false (区別しない) です。|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|検索の方向を指定します。 既定値は前方向です。 Excel.SearchDirection をご覧ください。|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)||
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)||
||[style](/javascript/api/excel/excel.settablecellproperties#style)||
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)||
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)||
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Shape オブジェクトの代替説明テキストを取得または設定します。|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Shape オブジェクトの代替タイトル テキストを取得または設定します。|
||[delete()](/javascript/api/excel/excel.shape#delete--)|ワークシートから図形を削除します。|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|この幾何学的図形の種類を表します。 詳細については、Excel.GeometricShapeType をご覧ください。 図形の種類が "GeometricShape" ではない場合は、null を返します。|
||[getAsImage(format: "UNKNOWN" \| "BMP" \| "JPEG" \| "GIF" \| "PNG" \| "SVG")](/javascript/api/excel/excel.shape#getasimage-format-)|図形を画像に変換し、base 64 でエンコードされた文字列として画像を返します。 DPI は 96 です。 サポートされている形式は、`Excel.PictureFormat.BMP`、`Excel.PictureFormat.PNG`、`Excel.PictureFormat.JPEG`、`Excel.PictureFormat.GIF` だけです。|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|図形を画像に変換し、base 64 でエンコードされた文字列として画像を返します。 DPI は 96 です。 サポートされている形式は、`Excel.PictureFormat.BMP`、`Excel.PictureFormat.PNG`、`Excel.PictureFormat.JPEG`、`Excel.PictureFormat.GIF` だけです。|
||[height](/javascript/api/excel/excel.shape#height)|図形の高さをポイント数で表します。|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|指定したポイント数だけ水平方向に図形を移動します。|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|z 軸を中心に、指定された度数だけ、図形を時計回りに回転します。|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|指定したポイント数だけ垂直方向に図形を移動します。|
||[left](/javascript/api/excel/excel.shape#left)|図形の左側からワークシートの左側までの距離 (ポイント数) です。|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|この図形の縦横比をロックするかどうかを指定します。|
||[name](/javascript/api/excel/excel.shape#name)|図形の名前を表します。|
||[placement](/javascript/api/excel/excel.shape#placement)|オブジェクトがその下のセルに接続されている方法を表します。|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|この図形の結合点の数を返します。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.shape#fill)|この図形の塗りつぶしの書式設定を返します。 読み取り専用です。|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|図形に関連付けられた幾何学的図形を返します。 図形の種類が "GeometricShape" ではない場合は、エラーがスローされます。|
||[group](/javascript/api/excel/excel.shape#group)|図形に関連付けられた図形グループを返します。 図形の種類が "GroupShape" ではない場合は、エラーがスローされます。|
||[id](/javascript/api/excel/excel.shape#id)|図形 ID を表します。 読み取り専用です。|
||[image](/javascript/api/excel/excel.shape#image)|図形に関連付けられた画像を返します。 図形の種類が "Image" ではない場合は、エラーがスローされます。|
||[level](/javascript/api/excel/excel.shape#level)|指定された図形のレベルを表します。 たとえば、レベル 0 は図形がどのグループの一部でもないことを意味し、レベル 1 は図形が最上位グループの一部であることを意味し、レベル 2 は図形が最上位レベルのサブグループの一部であることを意味します。|
||[line](/javascript/api/excel/excel.shape#line)|図形に関連付けられた線を返します。 図形の種類が "Line" ではない場合は、エラーがスローされます。|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|この図形の線の書式設定を返します。 読み取り専用です。|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|図形がアクティブになったときに発生します。|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|図形が非アクティブになると発生します。|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|この図形の親グループを表します。|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|この図形のテキスト フレーム オブジェクトを返します。 読み取り専用です。|
||[type](/javascript/api/excel/excel.shape#type)|この図形の種類を返します。 詳細については、Excel.ShapeType をご覧ください。 読み取り専用です。|
||[zorderPosition](/javascript/api/excel/excel.shape#zorderposition)|指定された図形の z オーダーでの位置を返します。0 はオーダー スタックの一番下を表します。 読み取り専用です。|
||[rotation](/javascript/api/excel/excel.shape#rotation)|図形の回転を角度で表します。|
||[scaleHeight(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|指定した係数分だけ図形の高さを変更します。 画像の場合は、図形を元のサイズに対して拡大または縮小するのか、現在のサイズに対して拡大または縮小するのかを指定できます。 画像以外の図形の場合は、常に現在の高さに対して拡大または縮小されます。|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|指定した係数分だけ図形の高さを変更します。 画像の場合は、図形を元のサイズに対して拡大または縮小するのか、現在のサイズに対して拡大または縮小するのかを指定できます。 図以外の図形の場合は、常に現在の高さに対して拡大または縮小されます。|
||[scaleWidth(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|指定した係数分だけ図形の幅を変更します。 画像の場合は、図形を元のサイズに対して拡大または縮小するのか、現在のサイズに対して拡大または縮小するのかを指定できます。 図以外の図形の場合は、常に現在の幅に対して拡大または縮小されます。|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|指定した係数分だけ図形の幅を変更します。 画像の場合は、図形を元のサイズに対して拡大または縮小するのか、現在のサイズに対して拡大または縮小するのかを指定できます。 画像以外の図形の場合は、常に現在の幅に対して拡大または縮小されます。|
||[setZOrder(position: "BringToFront" \| "BringForward" \| "SendToBack" \| "SendBackward")](/javascript/api/excel/excel.shape#setzorder-position-)|指定された図形をコレクションの z オーダーで上または下に移動します。他の図形の手前または奥に移動します。|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|指定された図形をコレクションの z オーダーで上または下に移動します。他の図形の手前または奥に移動します。|
||[top](/javascript/api/excel/excel.shape#top)|図形の上端からワークシートの上までのポイント単位の距離です。|
||[visible](/javascript/api/excel/excel.shape#visible)|この図形の可視性を表します。|
||[width](/javascript/api/excel/excel.shape#width)|図形の幅 (ポイント数) を表します。|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|アクティブ化された図形の ID を取得します。|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|図形がアクティブにされたワークシートの ID を取得します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: "LineInverse" \| "Triangle" \| "RightTriangle" \| "Rectangle" \| "Diamond" \| "Parallelogram" \| "Trapezoid" \| "NonIsoscelesTrapezoid" \| "Pentagon" \| "Hexagon" \| "Heptagon" \| "Octagon" \| "Decagon" \| "Dodecagon" \| "Star4" \| "Star5" \| "Star6" \| "Star7" \| "Star8" \| "Star10" \| "Star12" \| "Star16" \| "Star24" \| "Star32" \| "RoundRectangle" \| "Round1Rectangle" \| "Round2SameRectangle" \| "Round2DiagonalRectangle" \| "SnipRoundRectangle" \| "Snip1Rectangle" \| "Snip2SameRectangle" \| "Snip2DiagonalRectangle" \| "Plaque" \| "Ellipse" \| "Teardrop" \| "HomePlate" \| "Chevron" \| "PieWedge" \| "Pie" \| "BlockArc" \| "Donut" \| "NoSmoking" \| "RightArrow" \| "LeftArrow" \| "UpArrow" \| "DownArrow" \| "StripedRightArrow" \| "NotchedRightArrow" \| "BentUpArrow" \| "LeftRightArrow" \| "UpDownArrow" \| "LeftUpArrow" \| "LeftRightUpArrow" \| "QuadArrow" \| "LeftArrowCallout" \| "RightArrowCallout" \| "UpArrowCallout" \| "DownArrowCallout" \| "LeftRightArrowCallout" \| "UpDownArrowCallout" \| "QuadArrowCallout" \| "BentArrow" \| "UturnArrow" \| "CircularArrow" \| "LeftCircularArrow" \| "LeftRightCircularArrow" \| "CurvedRightArrow" \| "CurvedLeftArrow" \| "CurvedUpArrow" \| "CurvedDownArrow" \| "SwooshArrow" \| "Cube" \| "Can" \| "LightningBolt" \| "Heart" \| "Sun" \| "Moon" \| "SmileyFace" \| "IrregularSeal1" \| "IrregularSeal2" \| "FoldedCorner" \| "Bevel" \| "Frame" \| "HalfFrame" \| "Corner" \| "DiagonalStripe" \| "Chord" \| "Arc" \| "LeftBracket" \| "RightBracket" \| "LeftBrace" \| "RightBrace" \| "BracketPair" \| "BracePair" \| "Callout1" \| "Callout2" \| "Callout3" \| "AccentCallout1" \| "AccentCallout2" \| "AccentCallout3" \| "BorderCallout1" \| "BorderCallout2" \| "BorderCallout3" \| "AccentBorderCallout1" \| "AccentBorderCallout2" \| "AccentBorderCallout3" \| "WedgeRectCallout" \| "WedgeRRectCallout" \| "WedgeEllipseCallout" \| "CloudCallout" \| "Cloud" \| "Ribbon" \| "Ribbon2" \| "EllipseRibbon" \| "EllipseRibbon2" \| "LeftRightRibbon" \| "VerticalScroll" \| "HorizontalScroll" \| "Wave" \| "DoubleWave" \| "Plus" \| "FlowChartProcess" \| "FlowChartDecision" \| "FlowChartInputOutput" \| "FlowChartPredefinedProcess" \| "FlowChartInternalStorage" \| "FlowChartDocument" \| "FlowChartMultidocument" \| "FlowChartTerminator" \| "FlowChartPreparation" \| "FlowChartManualInput" \| "FlowChartManualOperation" \| "FlowChartConnector" \| "FlowChartPunchedCard" \| "FlowChartPunchedTape" \| "FlowChartSummingJunction" \| "FlowChartOr" \| "FlowChartCollate" \| "FlowChartSort" \| "FlowChartExtract" \| "FlowChartMerge" \| "FlowChartOfflineStorage" \| "FlowChartOnlineStorage" \| "FlowChartMagneticTape" \| "FlowChartMagneticDisk" \| "FlowChartMagneticDrum" \| "FlowChartDisplay" \| "FlowChartDelay" \| "FlowChartAlternateProcess" \| "FlowChartOffpageConnector" \| "ActionButtonBlank" \| "ActionButtonHome" \| "ActionButtonHelp" \| "ActionButtonInformation" \| "ActionButtonForwardNext" \| "ActionButtonBackPrevious" \| "ActionButtonEnd" \| "ActionButtonBeginning" \| "ActionButtonReturn" \| "ActionButtonDocument" \| "ActionButtonSound" \| "ActionButtonMovie" \| "Gear6" \| "Gear9" \| "Funnel" \| "MathPlus" \| "MathMinus" \| "MathMultiply" \| "MathDivide" \| "MathEqual" \| "MathNotEqual" \| "CornerTabs" \| "SquareTabs" \| "PlaqueTabs" \| "ChartX" \| "ChartStar" \| "ChartPlus")](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|幾何学的図形をワークシートに追加します。 新しい図形を表す Shape オブジェクトを返します。|
||[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|幾何学的図形をワークシートに追加します。 新しい図形を表す Shape オブジェクトを返します。|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|このコレクションのワークシート内の図形のサブセットをグループ化します。 図形の新しいグループを表す Shape オブジェクトを返します。|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|base64 エンコード文字列から画像を作成し、それをワークシートに追加します。 新しい画像を表す Shape オブジェクトを返します。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: "Straight" \| "Elbow" \| "Curve")](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|ワークシートに行を追加します。 新しい行を表す Shape オブジェクトを返します。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|ワークシートに行を追加します。 新しい行を表す Shape オブジェクトを返します。|
||[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。 新しい画像を表す Shape オブジェクトを返します。|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|指定されたテキストを含むテキスト ボックスをワークシートに追加します。 新しいテキスト ボックスを表す Shape オブジェクトを返します。|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|ワークシートの図形数を返します。 読み取り専用です。|
||[getItem(name: string)](/javascript/api/excel/excel.shapecollection#getitem-name-)|名前を使用して、その図形を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|コレクション内の位置を使用して図形を取得します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.shapecollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.shapecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|非アクティブにされた図形の ID を取得します。|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|図形が非アクティブにされたワークシートの ID を取得します。|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|この図形の塗りつぶしの書式設定をクリアします。|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|図形塗りつぶしの前面色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[type](/javascript/api/excel/excel.shapefill#type)|図形の塗りつぶしの種類を返します。 読み取り専用です。 詳細については、Excel.ShapeFillType をご覧ください。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|図形の塗りつぶしの書式設定を均一な色に設定します。 これにより、塗りつぶしの種類が "Solid" に変更されます。|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|塗りつぶしの透明度の割合を示す 0.0 (不透明) から 1.0 (透明) までの値を取得または設定します。 図形の種類が透明度をサポートしていない場合、または図形の塗りつぶしが一定ではない場合 (グラデーション塗りつぶしの種類など) は、null を返します。|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|フォントの太字の状態を表します。 TextRange に太字テキストと太字ではないテキストの両方が含まれている場合、null を返します。|
||[color](/javascript/api/excel/excel.shapefont#color)|テキストの色を表す HTML カラー コードです (例: "#FF0000" は赤を表します)。 TextRange に色の異なるテキストが含まれている場合、null を返します。|
||[italic](/javascript/api/excel/excel.shapefont#italic)|フォントの斜体の状態を表します。 TextRange に斜体テキストと斜体ではないテキストの両方が含まれている場合、null を返します。|
||[name](/javascript/api/excel/excel.shapefont#name)|フォント名 (例: "Calibri") を表します。 テキストが複雑なスクリプトか東アジアの言語の場合、それに対応するフォント名です。それ以外の場合は、ラテン フォントの名前です。|
||[size](/javascript/api/excel/excel.shapefont#size)|フォント サイズをポイント単位で表します (11 など)。 TextRange にサイズの異なるテキストが含まれている場合、null を返します。|
||[underline](/javascript/api/excel/excel.shapefont#underline)|フォントに適用する下線の種類。 TextRange に下線スタイルの異なるテキストが含まれている場合、null を返します。 詳細については、Excel.ShapeFontUnderlineStyle をご覧ください。|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|図形 ID を表します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|グループに関連付けられた Shape オブジェクトを返します。 読み取り専用です。|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Shape オブジェクトのコレクションを返します。 読み取り専用です。|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|指定した図形グループに含まれるグループ化された図形のグループを解除します。|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|線の色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|図形の線スタイルを表します。 線が非表示の場合、または破線のスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[style](/javascript/api/excel/excel.shapelineformat#style)|図形の線スタイルを表します。 線が非表示の場合、またはスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|指定された線の透明度を示す 0.0 (不透明) から 1.0 (透明) までの値を表します。 図形の透明度が一定ではない場合は、null を返します。|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|図形要素の線の書式設定が表示されるかどうかを表します。 図形の可視性が一定ではない場合は、null を返します。|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|線の太さ (ポイント数) を表します。 線が非表示の場合、または線の太さが一定ではない場合は、null を返します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|スライサーのキャプションを表します。|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|現在スライサーに適用されているすべてのフィルターを消去します。|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|スライサーを削除します。|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|選択されたアイテムのキーの配列を返します。 読み取り専用です。|
||[height](/javascript/api/excel/excel.slicer#height)|スライサーの高さ (ポイント数) を表します。|
||[left](/javascript/api/excel/excel.slicer#left)|スライサーの左側からワークシートの左までの距離を表します (ポイント数)。|
||[name](/javascript/api/excel/excel.slicer#name)|スライサーの名前を表します。|
||[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|数式で使用する名前を表します。|
||[id](/javascript/api/excel/excel.slicer#id)|スライサーの一意の ID を表します。 読み取り専用です。|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|スライサーに現在適用されているフィルターがすべて消去されている場合、true となります。|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|スライサーに含まれる SlicerItems のコレクションを表します。 読み取り専用です。|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|スライサーを含んでいるワークシートを表します。 読み取り専用です。|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|キーに基づいてスライサー アイテムを選択します。 前の選択は消去されます。|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|スライサーに含まれるアイテムの並べ替え順序を表します。 指定可能な値は DataSourceOrder、Ascending、Descending です。|
||[style](/javascript/api/excel/excel.slicer#style)|スライサー スタイルを表す定数値。 使用可能な値は、SlicerStyleLight1 から SlicerStyleLight6 まで、TableStyleOther1 から TableStyleOther2 まで、SlicerStyleDark1 から SlicerStyleDark6 までです。 ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
||[top](/javascript/api/excel/excel.slicer#top)|スライサーの上端からワークシートの上端までの距離を表します (ポイント数)。|
||[width](/javascript/api/excel/excel.slicer#width)|スライサーの幅 (ポイント数) を表します。|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|ブックに新しいスライサーを追加します。|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|コレクションに含まれるスライサーの数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|名前または ID を使用してスライサー オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|コレクション内の位置に基づいてスライサーを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|名前または ID に基づいてスライサーを取得します。スライサーが存在しない場合は null オブジェクトを返します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.slicercollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.slicercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|スライサー アイテムが選択されている場合、true となります。|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|スライサー アイテムにデータが含まれている場合、true となります。|
||[key](/javascript/api/excel/excel.sliceritem#key)|スライサー アイテムを表す一意の値を表します。|
||[name](/javascript/api/excel/excel.sliceritem#name)|UI に表示される値を表します。|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|スライサーのスライサー アイテム数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|そのキーまたは名前を利用してスライサー アイテム オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|コレクション内の位置に基づいてスライサー アイテムを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|そのキーまたは名前を使用してスライサー アイテムを取得します。 スライサー アイテムが存在しない場合は null オブジェクトを返します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.sliceritemcollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|SlicerStyle を削除します。|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|すべてのスタイル要素のコピーでこの SlicerStyle の複製を作成します。|
||[name](/javascript/api/excel/excel.slicerstyle#name)|SlicerStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|True は、この SlicerStyle オブジェクトが読み取り専用であることを意味します。 読み取り専用です。|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|指定された名前で空の SlicerStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|コレクション内のスライサー スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|親オブジェクトのスコープに対する既定の SlicerStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|名前で SlicerStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|名前で SlicerStyle を取得します。 SlicerStyle が存在しない場合は、null オブジェクトを返します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.slicerstylecollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の SlicerStyle を設定します。|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|並べ替え基準となるリッチな値のターゲット プロパティ名である下位フィールドを表します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|コレクション内のスタイルの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|コレクション内の位置に基づいてスタイルを取得します。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[autoFilter](/javascript/api/excel/excel.table#autofilter)|テーブルの AutoFilter オブジェクトを表します。 読み取り専用です。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|フィルターが特定のテーブルに適用されたときに発生します。|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|追加されたテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|テーブルが追加されたワークシートの ID を取得します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|変更の詳細に関する情報を表します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|新しいテーブルがブックに追加されたときに発生します。|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|指定されたテーブルがブックで削除されたときに発生します。|
||[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|ブックまたはワークシートのテーブルにフィルターが適用されたときに発生します。|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|イベントのソースを指定します。 詳細については、Excel.EventSource をご覧ください。|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|削除されたテーブルの ID を指定します。|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|削除されたテーブルの名前を指定します。|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|イベントの種類を指定します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|テーブルが削除されたワークシートの ID を指定します。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|フィルターが適用されたテーブルの ID を表します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|イベントの種類を表します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|テーブルが含まれるワークシートの ID を表します。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|コレクションに含まれるテーブルの数を取得します。|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|コレクション内の最初のテーブルを取得します。 一番上の左のテーブルがコレクション内で最初のテーブルになるように、コレクションのテーブルが上から下へ、左から右への順で並べ替えられます。|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|名前または ID を使用してテーブルを取得します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.tablescopedcollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|TableStyle を削除します。|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|すべてのスタイル要素のコピーでこの TableStyle の複製を作成します。|
||[name](/javascript/api/excel/excel.tablestyle#name)|TableStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|true は、この TableStyle オブジェクトが読み取り専用であることを意味します。 読み取り専用です。|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|指定された名前で空の TableStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|コレクションに含まれるテーブル スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|親オブジェクトのスコープに対する既定の TableStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|名前で TableStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|名前で TableStyle を取得します。 TableStyle が存在しない場合は、null オブジェクトを返します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.tablestylecollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の TableStyle を設定します。|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|テキスト フレームの自動サイズ変更設定を取得または設定します。 テキストをテキスト フレームに自動的に合わせる、テキスト フレームをテキストに自動的に合わせる、自動サイズ変更を行わない、のいずれかにテキスト フレームを設定できます。|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|テキスト フレームの下余白を表します (ポイント数)。|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|テキスト フレーム内のテキストをすべて削除します。|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|テキスト フレームの水平方向の配置を表します。 詳細については、Excel.ShapeTextHorizontalAlignment を参照してください。|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|テキスト フレームの水平方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextHorizontalOverflow を参照してください。|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|テキスト フレームの左余白を表します (ポイント数)。|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|テキスト フレームのテキストの向きを表します。 詳細については、Excel.ShapeTextOrientation を参照してください。|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|テキスト フレームの読む方向を表します (左から右または右から左)。 詳細については、Excel.ShapeTextReadingOrder を参照してください。|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|テキスト フレームにテキストが含まれるかどうかを指定します。|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|テキスト フレーム内の図形にアタッチされているテキスト、およびテキストを操作するためのプロパティとメソッドを表します。 詳細については、Excel.TextRange を参照してください。|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|テキスト フレームの右余白を表します (ポイント数)。|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|テキスト フレームの上余白を表します (ポイント数)。|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|テキスト フレームの垂直方向の配置を表します。 詳細については、Excel.ShapeTextVerticalAlignment を参照してください。|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|テキスト フレームの垂直方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextVerticalOverflow を参照してください。|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|指定された範囲の部分文字列に対する TextRange オブジェクトを返します。|
||[font](/javascript/api/excel/excel.textrange#font)|テキスト範囲のフォント属性を表す ShapeFont オブジェクトを返します。 読み取り専用です。|
||[text](/javascript/api/excel/excel.textrange#text)|テキスト範囲のプレーン テキスト コンテンツを表します。|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|TableStyle を削除します。|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|すべてのスタイル要素のコピーでこの TimelineStyle の複製を作成します。|
||[name](/javascript/api/excel/excel.timelinestyle#name)|TimelineStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|true は、この TimelineStyle オブジェクトが読み取り専用であることを意味します。 読み取り専用です。|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|指定された名前で空の TimelineStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|コレクションに含まれるタイムライン スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|親オブジェクトのスコープに対する既定の TimelineStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|名前で TimelineStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|名前で TimelineStyle を取得します。 TimelineStyle が存在しない場合は、null オブジェクトを返します。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.timelinestylecollection#load-option-)|オブジェクトの指定のプロパティを読み込むコマンドを待ち行列に入れます。 プロパティを読み取るには先に "context.sync()" を呼び出す必要があります。|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の TimelineStyle を設定します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|関連付けられている実際のデータ ポイントをブックの全グラフが追跡している場合、true となります。|
||[close(closeBehavior?: "Save" \| "SkipSave")](/javascript/api/excel/excel.workbook#close-closebehavior-)|現在のブックを閉じます。|
||[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|現在のブックを閉じます。|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|ブックで現在アクティブになっているグラフを取得します。 アクティブになっているグラフがない場合、このステートメントを呼び出すと、例外がスローされます。|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|ブックで現在アクティブになっているグラフを取得します。 アクティブになっているグラフがない場合、null オブジェクトを返します。|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|ブックで現在アクティブになっているスライサーを取得します。 アクティブになっているスライサーがない場合、このステートメントを呼び出すと、例外がスローされます。|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|ブックで現在アクティブになっているスライサーを取得します。 アクティブになっているスライサーがない場合、null オブジェクトを返します。|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|ブックが複数のユーザーによって編集されている場合 (共同編集)、true となります。|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|ブックから現在選択されている 1 つまたは複数の範囲を取得します。 getSelectedRange() の場合と同様に、このメソッドは、選択されているすべての範囲を表す RangeAreas オブジェクトを返します。|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|ブックが最後に保存された後に変更が行われたかどうかを指定します。|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|ブックが自動保存モードかどうかを指定します。 読み取り専用です。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Excel 計算エンジンのバージョンとして数字を返します。 読み取り専用です。|
||[comments](/javascript/api/excel/excel.workbook#comments)|ブックに関連付けられているコメントの集まりを表します。 読み取り専用です。|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|ブックで autoSave の設定が変更されると発生します。|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|ブックに関連付けられている PivotTableStyle のコレクションを表します。 読み取り専用です。|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|ブックがローカル環境またはオンライン環境に保存されたかどうかを指定します。 読み取り専用です。|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|ブックに関連付けられている SlicerStyle のコレクションを表します。 読み取り専用です。|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|ブックに関連付けられているスライサーの集まりを表します。 読み取り専用です。|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|ブックに関連付けられている TableStyle のコレクションを表します。 読み取り専用です。|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|ブックに関連付けられている TimelineStyle のコレクションを表します。 読み取り専用です。|
||[save(saveBehavior?: "Save" \| "Prompt")](/javascript/api/excel/excel.workbook#save-savebehavior-)|現在のブックを保存します。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|現在のブックを保存します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|ブックを表示桁数でのみ計算する場合、true となります。|
|[WorkbookAutoSaveSetting[...]](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|イベントの種類を表します。 詳細については、Excel.EventType をご覧ください。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|ワークシートの enableCalculation プロパティを取得または設定します。|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|指定された条件に基づいて指定された文字列の発生箇所をすべて見つけ、1 つまたは複数の長方形範囲を構成する RangeAreas オブジェクトとして返します。|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|指定された条件に基づいて指定された文字列の発生箇所をすべて見つけ、1 つまたは複数の長方形範囲を構成する RangeAreas オブジェクトとして返します。|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|アドレスまたは名前で指定され、1 つまたは複数の長方形範囲ブロックを表す RangeAreas オブジェクトを取得します。|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|ワークシートの AutoFilter オブジェクトを表します。 読み取り専用です。|
||[comments](/javascript/api/excel/excel.worksheet#comments)|ワークシート上のすべての Comments オブジェクトの集まりを返します。 読み取り専用です。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|ワークシートの水平改ページをまとめて取得します。 このコレクションには、手動の改ページのみが含まれます。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|フィルターが特定のワークシートに適用されたときに発生します。|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|フォーマットが特定のワークシートで変更されたときに発生します。|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|ワークシートの PageLayout オブジェクトを取得します。|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|ワークシート上のすべての Shape オブジェクトをまとめて返します。 読み取り専用です。|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|ワークシートに含まれるスライサーをまとめて返します。 読み取り専用です。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|ワークシートの垂直改ページをまとめて取得します。 このコレクションには、手動の改ページのみが含まれます。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|現在のワークシート内で、指定された条件に基づき、指定された文字列を検索し、置換します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|変更の詳細に関する情報を表します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: "None" \| "Before" \| "After" \| "Beginning" \| "End", relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|ブックのワークシートが変更されたときに発生します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|ブック内のワークシートの書式が変更されたときに発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|ワークシートで選択範囲を変更したときに発生します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を表します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されたワークシートの ID を表します。|
|[WorksheetFormatChanged[...]](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。 null オブジェクトを返すこともあります。|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|一致方法として完全一致か部分一致を指定します。 既定値は false (部分一致) です。|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|照合の際に大文字と小文字を区別するかどうかを指定します。 既定値は false (区別しない) です。|

## <a name="whats-new-in-excel-javascript-api-18"></a>Excel JavaScript API 1.8 の新機能

Excel JavaScript API 要件セット 1.8 の機能には、ピボットテーブル、データの入力規則、グラフ、グラフのイベント、パフォーマンス オプション、ブック作成に対応する API が含まれます。

### <a name="pivottable"></a>ピボットテーブル

ピボットテーブル API の Wave 2 では、アドインでピボットテーブルの階層を設定できます。 データとデータの集計方法を制御できるようになりました。 新しいピボットテーブルの機能について詳しくは、[ピボットテーブルの記事](/office/dev/add-ins/excel/excel-add-ins-pivottables)を参照してください。

### <a name="data-validation"></a>データの入力規則

データの入力規則により、ユーザーがワークシートに入力する内容を制御できます。 定義済みの回答セットにセルを制限したり、望ましくない入力に関する警告をポップアップ表示したりできます。 詳細については、[データの入力規則を範囲に追加する方法](/office/dev/add-ins/excel/excel-add-ins-data-validation)を参照してください。

### <a name="charts"></a>グラフ

グラフ要素をより詳細にプログラムで制御できる、一連のグラフ API がさらに追加されました。 凡例、軸、近似曲線、プロット エリアがより使いやすくなっています。

### <a name="events"></a>イベント

グラフの[イベント](/office/dev/add-ins/excel/excel-add-ins-events)がさらに追加されました。 グラフを操作するユーザーに対し、アドインで対応できます。 ブック全体にわたり、起動する[イベントの切り替え](/office/dev/add-ins/excel/performance#enable-and-disable-events)もできます。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[アプリケーション](/javascript/api/excel/excel.application)|_メソッド_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|base64 でエンコードされたオプションの .xlsx ファイルを使用して、非表示のブックを新規に作成します。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_プロパティ_ > formula1|Formula1 (最小値) または演算子に応じた値を取得あるいは設定します。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_プロパティ_ > formula2|Formula2 (最大値) または演算子に応じた値を取得あるいは設定します。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_リレーションシップ_ > operator|データの検証に使用する演算子。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_プロパティ_ > categoryLabelLevel|項目ラベルのソースとなるレベルを参照している ChartCategoryLabelLevel 列挙体定数を返すか設定します。 読み取り/書き込み可能。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_プロパティ_ > plotVisibleOnly|true の場合、可視セルだけがプロットされます。 false の場合、可視セルと非表示セルの両方がプロットされます。 読み取り/書き込み可能。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_プロパティ_ > seriesNameLevel|系列名のソースとなるレベルを参照している ChartSeriesNameLevel 列挙体定数を返すか設定します。 読み取り/書き込み可能。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_プロパティ_ > showDataLabelsOverMaximum|値が数値軸の最大値より大きい場合にデータ ラベルを表示するかどうかを表します。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_プロパティ_ > style|グラフのグラフ スタイルを返すか設定します。 読み取り/書き込み可能。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_リレーションシップ_ > displayBlanksAs|空白のセルがグラフでプロットされる方法を返すか設定します。 読み取り/書き込み可能。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_リレーションシップ_ > plotArea|グラフのプロット エリアを表します。 読み取り専用です。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_リレーションシップ_ > plotBy|グラフ上で列または行がデータ系列として使用される方法を返すか設定します。 読み取り/書き込み可能。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_プロパティ_ > chartId|アクティブにされたグラフの ID を取得します。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_プロパティ_ > type|イベントの種類を取得します。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_プロパティ_ > worksheetId|グラフがアクティブにされたワークシートの ID を取得します。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_プロパティ_ > chartId|ワークシートに追加されたグラフの ID を取得します。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_プロパティ_ > type|イベントの種類を取得します。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_プロパティ_ > worksheetId|グラフが追加されたワークシートの ID を取得します。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_リレーションシップ_ > source|イベントのソースを取得します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > isBetweenCategories|項目の境界で数値軸が項目軸と交差するかどうかを表します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > multiLevel|軸がマルチレベルかどうかを表します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > numberFormat|軸の目盛ラベルの書式コードを表します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > offset|ラベルのレベル間の距離、および先頭レベルと軸線との距離を表します。 値は 0 から 1000 の範囲内でなければなりません。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > positionAt|指定した軸と他の軸との交差位置を表します。 このプロパティを設定するには、SetPositionAt(double) メソッドを使用する必要があります。 読み取り専用です。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > textOrientation|軸の目盛ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_リレーションシップ_ > alignment|指定した軸の目盛ラベルの配置を表します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_リレーションシップ_ > position|指定した軸と他の軸との交差位置を表します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_メソッド_ > [setPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|指定した軸と他の軸との交差位置を設定します。|1.8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_リレーションシップ_ > fill|グラフの塗りつぶしの書式設定を表します。 読み取り専用です。|1.8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_メソッド_ > [setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|A1 スタイルの表記法を使用するグラフの軸タイトルの数式を表す文字列値。|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_リレーションシップ_ > border|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_リレーションシップ_ > fill|グラフの塗りつぶしの書式設定を表します。 読み取り専用です。|1.8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_メソッド_ > [clear()](/javascript/api/excel/excel.chartborder)|グラフ要素の罫線の書式設定をクリアします。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > autoText|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > formula|A1 スタイルの表記法を使用するグラフのデータ ラベルの数式を表す文字列値。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > height|グラフのデータ ラベルの高さ (ポイント数) を返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null 値となります。 読み取り専用です。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > left|グラフのデータ ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > numberFormat|データ ラベルの書式コードを表す文字列値。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > text|グラフのデータ ラベルのテキストを表す文字列。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > textOrientation|グラフのデータ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > top|グラフのデータ ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > width|グラフのデータ ラベルの幅 (ポイント数) を返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null 値となります。 読み取り専用です。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_リレーションシップ_ > format|グラフのデータ ラベルの書式設定を表します。 読み取り専用です。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_リレーションシップ_ > horizontalAlignment|グラフのデータ ラベルの水平方向の配置を表します。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_リレーションシップ_ > verticalAlignment|グラフのデータ ラベルの垂直方向の配置を表します。|1.8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_リレーションシップ_ > border|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_プロパティ_ > autoText|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表します。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_プロパティ_ > numberFormat|データ ラベルの書式コードを表します。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_プロパティ_ > textOrientation|データ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数、縦書きテキストの場合は 0 から 180 の範囲内の整数でなければなりません。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_リレーションシップ_ > horizontalAlignment|グラフのデータ ラベルの水平方向の配置を表します。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_リレーションシップ_ > verticalAlignment|グラフのデータ ラベルの垂直方向の配置を表します。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_プロパティ_ > chartId|非アクティブにされたグラフの ID を取得します。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_プロパティ_ > type|イベントの種類を取得します。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_プロパティ_ > worksheetId|グラフが非アクティブにされたワークシートの ID を取得します。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_プロパティ_ > chartId|ワークシートから削除されたグラフの ID を取得します。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_プロパティ_ > type|イベントの種類を取得します。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_プロパティ_ > worksheetId|グラフが削除されたワークシートの ID を取得します。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_リレーションシップ_ > source|イベントのソースを取得します。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_ > height|グラフの凡例に表示される凡例エントリの高さを表します。 読み取り専用です。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_ > index|グラフの凡例に含まれる凡例エントリのインデックスを表します。 読み取り専用です。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_ > left|グラフの凡例エントリの左を表します。 読み取り専用です。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_ > top|グラフの凡例エントリの上を表します。 読み取り専用です。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_ > width|グラフの凡例に表示される凡例エントリの幅を表します。 読み取り専用です。|1.8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_リレーションシップ_ > border|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > height|プロット エリアの height 値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > insideHeight|プロット エリアの insideHeight 値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > insideLeft|プロット エリアの insideLeft 値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > insideTop|プロット エリアの insideTop 値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > insideWidth|プロット エリアの insideWidth 値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > left|プロット エリアの left 値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > top|プロット エリアの top 値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > width|プロット エリアの width 値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_リレーションシップ_ > format|グラフ プロット エリアの書式設定を表します。 読み取り専用です。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_リレーションシップ_ > position|プロット エリアの位置を表します。|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_リレーションシップ_ > border|グラフ プロット エリアの罫線の属性を表します。 読み取り専用です。|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_リレーションシップ_ > fill|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。読み取り専用。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > explosion|円グラフまたはドーナツ グラフのスライス切り出し表示の値を返すか設定します。 切り出し表示は行われず、スライスの先端が円の中心と一致する場合、0 を返します。 読み取り/書き込み可能。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > firstSliceAngle|円グラフまたはドーナツ グラフの最初のスライスの角度 (縦の中心から時計回りでの度数) を返すか設定します。 円グラフ、3-D 円グラフ、およびドーナツ グラフにのみ適用されます。 0 から 360 の範囲内で値を指定できます。 読み取り/書き込み可能。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > invertIfNegative|true の場合、Microsoft Excel により、負の数値に対応する項目でパターンが反転されます。 読み取り/書き込み可能。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > overlap|横棒と縦棒の配置方法を指定します。 -100 から 100 の範囲内で値を指定できます。 2-D 横棒グラフと 2-D 縦棒グラフにのみ適用されます。 読み取り/書き込み可能。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > secondPlotSize|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフのセカンダリ セクションのサイズを、プライマリ セクションのサイズのパーセンテージとして返すか設定します。 5 から 200 の範囲内で値を指定できます。 読み取り/書き込み可能。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > varyByCategories|true の場合、Microsoft Excel により、データ マーカーごとに異なる色またはパターンが割り当てられます。 グラフに含まれるデータ系列は 1 つだけでなければなりません。 読み取り/書き込み可能。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_リレーションシップ_ > axisGroup|指定した系列のグループを取得または設定します。値の取得および設定が可能です。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_リレーションシップ_ > dataLabels|系列に含まれるすべてのデータ ラベルのコレクションを表します。 読み取り専用です。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_リレーションシップ_ > splitType|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフを 2 つの部分に分割する方法を返すか設定します。 読み取り/書き込み可能。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > backwardPeriod|近似曲線を後方へ拡張するときの区間数を表します。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > forwardPeriod|近似曲線を前方へ拡張するときの区間数を表します。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > showEquation|true の場合、グラフに近似曲線の数式が表示されます。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > showRSquared|true の場合、グラフに近似曲線の R-2 乗値が表示されます。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_リレーションシップ_ > label|グラフの近似曲線のラベルを表します。 読み取り専用です。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > autoText|近似曲線ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > formula|A1 スタイルの表記法を使用するグラフの近似曲線ラベルの数式を表す文字列値。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > height|グラフの近似曲線ラベルの高さ (ポイント数) を返します。 読み取り専用です。 グラフの近似曲線ラベルが表示されない場合は null 値となります。 読み取り専用です。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > left|グラフの近似曲線ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > numberFormat|近似曲線ラベルの書式コードを表す文字列値。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > text|グラフの近似曲線ラベルのテキストを表す文字列。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > textOrientation|グラフの近似曲線ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > top|グラフの近似曲線ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > width|グラフの近似曲線ラベルの幅 (ポイント数) を返します。 読み取り専用です。 グラフの近似曲線ラベルが表示されない場合は null 値となります。 読み取り専用です。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_リレーションシップ_ > format|グラフの近似曲線ラベルの書式設定を表します。 読み取り専用です。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_リレーションシップ_ > horizontalAlignment|グラフの近似曲線ラベルの水平方向の配置を表します。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_リレーションシップ_ > verticalAlignment|グラフの近似曲線ラベルの垂直方向の配置を表します。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_リレーションシップ_ > border|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_リレーションシップ_ > fill|現在のグラフの近似曲線ラベルの塗りつぶしの書式設定を表します。 読み取り専用です。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_リレーションシップ_ > font|グラフの近似曲線ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|1.8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_プロパティ_ > formula| ユーザーの入力規則のカスタム数式。 セルの範囲での重複の防止や合計の制限など、特殊な入力規則を作成します。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_プロパティ_ > id|DataPivotHierarchy の ID。 読み取り専用です。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_プロパティ_ > name|DataPivotHierarchy の名前。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_プロパティ_ > numberFormat|DataPivotHierarchy の数値形式。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_プロパティ_ > position|DataPivotHierarchy の位置。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_リレーションシップ_ > field|DataPivotHierarchy に関連付けられているピボット フィールドを返します。 読み取り専用です。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_リレーションシップ_ > showAs|データを特定の集計計算として表示するかどうかどうかを指定します。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_リレーションシップ_ > summarizeBy|DataPivotHierarchy のすべての項目を表示するかどうかを指定します。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_メソッド_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|DataPivotHierarchy を既定値にリセットします。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_プロパティ_ > items|dataPivotHierarchy オブジェクトのコレクション。 読み取り専用です。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_メソッド_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|現在の軸にピボット階層を追加します。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|コレクションに含まれるピボット階層の数を取得します。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|名前または ID に基づいて DataPivotHierarchy を取得します。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|名前に基づいて DataPivotHierarchy を取得します。 DataPivotHierarchy が存在しない場合は null オブジェクトを返します。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_メソッド_ > [remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|現在の軸からピボット階層を削除します。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_プロパティ_ > ignoreBlanks|空白を無視します。つまり、空白のセルではデータの入力規則が検証されません。既定では true に設定されます。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_プロパティ_ > valid|すべてのセルの値がデータの入力規則に従っているかどうかを表します。 読み取り専用です。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_リレーションシップ_ > errorAlert|無効なデータが入力された場合のエラー警告。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_リレーションシップ_ > prompt|ユーザーにセルの選択を求めるときのプロンプト。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_リレーションシップ_ > rule|各種のデータ検証基準を含む、データの入力規則。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_リレーションシップ_ > type|データの入力規則の種類。詳細については、[Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype) を参照してください。 読み取り専用です。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_メソッド_ > [clear()](/javascript/api/excel/excel.datavalidation)|現在の範囲からデータの入力規則をクリアします。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_プロパティ_ > message|エラー警告メッセージを表します。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_プロパティ_ > showAlert|無効なデータが入力されたときにエラー警告ダイアログを表示するかどうかを指定します。 既定値は true です。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_プロパティ_ > title|エラー警告ダイアログのタイトルを表します。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_リレーションシップ_ > style|データの入力規則に対する警告の種類を表します。詳細については、[Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle) を参照してください。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_プロパティ_ > message|プロンプトのメッセージを表します。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_プロパティ_ > showPrompt|ユーザーがデータの入力規則が適用されているセルを選択したときに、プロンプトを表示するかどうかを指定します。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_プロパティ_ > title|プロンプトのタイトルを表します。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_リレーションシップ_ > custom|データ検証条件のカスタム数式。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_リレーションシップ_ > date|日付のデータ検証条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_リレーションシップ_ > decimal|10 進数のデータ検証条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_リレーションシップ_ > list|リストのデータ検証条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_リレーションシップ_ > textLength|テキスト長のデータ検証条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_リレーションシップ_ > time|時刻のデータ検証条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_リレーションシップ_ > wholeNumber|整数のデータ検証条件。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_プロパティ_ > formula1|Formula1 (最小値) または演算子に応じた値を取得あるいは設定します。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_プロパティ_ > formula2|Formula2 (最大値) または演算子に応じた値を取得あるいは設定します。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_リレーションシップ_ > operator|データの検証に使用する演算子。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_プロパティ_ > enableMultipleFilterItems|複数のフィルター項目を許可するかどうかを指定します。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_プロパティ_ > id|FilterPivotHierarchy の ID。 読み取り専用です。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_プロパティ_ > name|FilterPivotHierarchy の名前。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_プロパティ_ > position|FilterPivotHierarchy の位置。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_リレーションシップ_ > fields|FilterPivotHierarchy に関連付けられているピボット フィールドを返します。 読み取り専用です。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_メソッド_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|FilterPivotHierarchy を既定値にリセットします。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_プロパティ_ > items|filterPivotHierarchy オブジェクトのコレクション。 読み取り専用です。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_メソッド_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|現在の軸にピボット階層を追加します。 階層が行、列、またはフィルター軸の他の場所に存在する場合は、その場所から削除されます。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|コレクションに含まれるピボット階層の数を取得します。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|名前または ID に基づいて FilterPivotHierarchy を取得します。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|名前に基づいて FilterPivotHierarchy を取得します。 FilterPivotHierarchy が存在しない場合は null オブジェクトを返します。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_メソッド_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|現在の軸からピボット階層を削除します。|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_プロパティ_ > inCellDropDown|セルのドロップダウンにリストを表示するかどうかを指定します。既定では true に設定されます。|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_プロパティ_ > source|データの入力規則のリストのソース。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_プロパティ_ > id|PivotField の ID。 読み取り専用です。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_プロパティ_ > name|PivotField の名前。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_プロパティ_ > showAllItems|PivotField のすべての項目を表示するかどうかを指定します。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_リレーションシップ_ > items|PivotField に関連付けられているピボット フィールドを返します。 読み取り専用です。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_リレーションシップ_ > subtotals|PivotField の小計。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_メソッド_ > [sortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|PivotField を並べ替えます。 DataPivotHierarchy を指定すると、そのピボット階層に基づいて並べ替えが適用されます。指定しない場合、ピボット フィールド自体が並べ替えの基準になります。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_プロパティ_ > items|pivotField オブジェクトのコレクション。 読み取り専用です。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|コレクションに含まれるピボット階層の数を取得します。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|名前または ID に基づいて PivotHierarchy を取得します。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|名前に基づいて PivotHierarchy を取得します。 PivotHierarchy が存在しない場合は null オブジェクトを返します。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_プロパティ_ > id|PivotHierarchy の ID。 読み取り専用です。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_プロパティ_ > name|PivotHierarchy の名前。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_リレーションシップ_ > fields|PivotHierarchy に関連付けられているピボット フィールドを返します。 読み取り専用です。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_プロパティ_ > items|pivotHierarchy オブジェクトのコレクション。 読み取り専用です。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|コレクションに含まれるピボット階層の数を取得します。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|名前または ID に基づいて PivotHierarchy を取得します。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|名前に基づいて PivotHierarchy を取得します。 PivotHierarchy が存在しない場合は null オブジェクトを返します。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_プロパティ_ > id|PivotItem の ID。 読み取り専用です。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_プロパティ_ > isExpanded|項目を展開して子項目を表示するか、または項目を折りたたんで子項目を非表示にするかを指定します。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_プロパティ_ > name|PivotItem の名前。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_プロパティ_ > visible|PivotItem を表示するかどうかを指定します。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_プロパティ_ > items|pivotItem オブジェクトのコレクション。 読み取り専用です。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.pivotitemcollection)|コレクションに含まれるピボット階層の数を取得します。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection)|名前または ID に基づいて PivotHierarchy を取得します。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection)|名前に基づいて PivotHierarchy を取得します。 PivotHierarchy が存在しない場合は null オブジェクトを返します。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_プロパティ_ > showColumnGrandTotals|true の場合、ピボットテーブル レポートに列の総計が表示されます。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_プロパティ_ > showRowGrandTotals|true の場合、ピボットテーブル レポートに行の総計が表示されます。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_プロパティ_ > subtotalLocation|このプロパティは、ピボットテーブルのすべてのフィールドの SubtotalLocationType を示します。 フィールドによって状態が異なる場合は null 値になります。 有効な値は、AtTop、AtBottom です。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_リレーションシップ_ > layoutType|このプロパティは、ピボットテーブルのすべてのフィールドの PivotLayoutType を示します。 フィールドによって状態が異なる場合は null 値になります。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_メソッド_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|ピボットテーブルの列ラベルが存在する範囲を返します。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_メソッド_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|ピボットテーブルのデータ値が存在する範囲を返します。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_メソッド_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|ピボットテーブルのフィルター エリアの範囲を返します。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_メソッド_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|フィルター エリアを除く、ピボットテーブルが存在する範囲を返します。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_メソッド_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|ピボットテーブルの行ラベルが存在する範囲を返します。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_リレーションシップ_ > columnHierarchies|ピボットテーブルの列ピボット階層。 読み取り専用です。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_リレーションシップ_ > dataHierarchies|ピボットテーブルのデータ ピボット階層。 読み取り専用です。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_リレーションシップ_ > filterHierarchies|ピボットテーブルのフィルター ピボット階層。 読み取り専用です。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_リレーションシップ_ > hierarchies|ピボットテーブルのピボット階層。 読み取り専用です。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_リレーションシップ_ > layout|ピボットテーブルのレイアウトとビジュアル構造を記述する PivotLayout。 読み取り専用です。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_リレーションシップ_ > rowHierarchies|ピボットテーブルの行ピボット階層。 読み取り専用です。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_メソッド_ > [delete()](/javascript/api/excel/excel.pivottable)|ピボットテーブルを削除します。|1.8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_メソッド_ > [add(name: string, source: object, destination: object)](/javascript/api/excel/excel.pivottablecollection)|指定したソース データに基づくピボットテーブルを追加し、コピー先範囲の左上のセルに挿入します。|1.8|
|[range](/javascript/api/excel/excel.range)|_リレーションシップ_ > dataValidation|dataValidation オブジェクトを返します。 読み取り専用です。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_プロパティ_ > id|RowColumnPivotHierarchy の ID。 読み取り専用です。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_プロパティ_ > name|RowColumnPivotHierarchy の名前。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_プロパティ_ > position|RowColumnPivotHierarchy の位置。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_リレーションシップ_ > fields|RowColumnPivotHierarchy に関連付けられているピボット フィールドを返します。 読み取り専用です。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_メソッド_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|RowColumnPivotHierarchy を既定値にリセットします。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_プロパティ_ > items|RowColumnPivotHierarchy オブジェクトのコレクション。 読み取り専用です。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_メソッド_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|現在の軸にピボット階層を追加します。 階層が行、列、またはフィルター軸の他の場所に存在する場合は、その場所から削除されます。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|コレクションに含まれるピボット階層の数を取得します。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|名前または ID に基づいて RowColumnPivotHierarchy を取得します。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|名前に基づいて RowColumnPivotHierarchy を取得します。 RowColumnPivotHierarchy が存在しない場合は null オブジェクトを返します。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_メソッド_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|現在の軸からピボット階層を削除します。|1.8|
|[runtime](/javascript/api/excel/excel.runtime)|_プロパティ_ > enableEvents|現在の作業ウィンドウまたはコンテンツ アドインで JavaScript イベントを切り替えます。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_リレーションシップ_ > baseField|ShowAsCalculation 型に基づき、該当する場合は ShowAs 計算の基準となるベース ピボット フィールド。それ以外の場合は null 値です。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_リレーションシップ_ > baseItem|ShowAsCalculation 型に基づき、該当する場合は ShowAs 計算の基準となるベース項目。それ以外の場合は null 値です。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_リレーションシップ_ > calculation|データ ピボット フィールドに使用する ShowAs 計算。|1.8|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > autoIndent|セル内のテキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|1.8|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > textOrientation|スタイルで適用されるテキストの向き。|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > automatic|automatic が true に設定されている場合、小計を設定する際に、他の値はすべて無視されます。|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > average| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > count| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > countNumbers| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > max| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > min| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > product| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > standardDeviation| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > standardDeviationP| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > sum| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > variance| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_プロパティ_ > varianceP| |1.8|
|[table](/javascript/api/excel/excel.table)|_プロパティ_ > legacyId|数値の ID を返します。読み取り専用。|1.8|
|[workbook](/javascript/api/excel/excel.workbook)|_プロパティ_ > readOnly|true の場合、ブックが読み取り専用モードで開かれます。 読み取り専用です。|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_プロパティ_ > id|WorkbookCreated オブジェクトを一意に識別する値を返します。 読み取り専用です。|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_メソッド_ > [open()](/javascript/api/excel/excel.workbookcreated)|ブックを開きます。|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_プロパティ_> > showGridlines|ワークシートの gridlines フラグを取得または設定します。|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_プロパティ_ > showHeadings|ワークシートの headings フラグを取得または設定します。|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_プロパティ_ > type|イベントの種類を取得します。|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_プロパティ_ > worksheetId|計算対象のワークシートの ID を取得します。|1.8|

## <a name="whats-new-in-excel-javascript-api-17"></a>Excel JavaScript API 1.7 の新機能

Excel JavaScript API 要件セット 1.7 の機能には、グラフ、イベント、ワークシート、範囲、ドキュメント プロパティ、名前付きアイテム、保護のオプションとスタイルに対応する API が含まれます。

### <a name="customize-charts"></a>グラフのカスタマイズ

新しいグラフ API を使用して実行できる操作には、他の種類のグラフの作成、グラフへのデータ系列の追加、グラフ タイトルの設定、軸タイトルの追加、表示単位の追加、移動平均を使用した近似曲線の追加、線形近似曲線への変更などがあります。 以下はその例です。

* グラフの軸: グラフ内の軸の単位、ラベル、タイトルを取得、設定、書式設定する。
* グラフのデータ系列: グラフ内のデータ系列を追加、設定、削除する。  データ系列マーカー、プロット順序、サイジングを変更する。
* グラフの近似曲線: グラフ内の近似曲線を追加、取得、書式設定する。
* グラフの凡例: グラフ内の凡例のフォントを書式設定する。
* グラフ データ ポイント要素: データ要素の色を設定する。
* グラフ タイトルのサブ文字列: グラフ タイトルのサブ文字列を取得、設定する。
* グラフの種類: 他の種類のグラフを作成するオプションを使用する。

### <a name="events"></a>イベント

Excel イベント API には各種のイベント ハンドラーが用意されています。これらのハンドラーを使用することで、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できます。 実行する関数は、目的のシナリオに必要な処理を行うように設計できます。 現在利用可能なイベントのリストについては、[Excel JavaScript API を使用してイベントを操作する](/office/dev/add-ins/excel/excel-add-ins-events)を参照してください。

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>ワークシートと範囲の外観のカスタマイズ

新しい API を使用して、ワークシートの外観をさまざまな方法でカスタマイズできます。たとえば、次のようなカスタマイズが可能です。

* ワークシート内でスクロールするときに特定の行または列が常に表示されるよう、ウィンドウ枠を固定する。 たとえば、ワークシート内の最初の行にヘッダーが示される場合、その行にウィンドウ枠を固定すると、ワークシートをスクロールダウンしても列の見出しは表示されたままになります。
* ワークシートのタブの色を変更する。
* ワークシートの見出しを追加する。


範囲の外観をさまざまな方法でカスタマイズできます。たとえば、次のようなカスタマイズが可能です。

* 特定の範囲に対してセルのスタイルを設定し、その範囲内のすべてのセルに一貫した書式設定が適用されるようにする。 セルのスタイルとは、フォント、フォントのサイズ、数値形式、セルの罫線、セルの網掛けなど、文字に定義された書式設定一式を指します。 Excel の組み込みのセル スタイルのいずれかを使用することも、独自のカスタム セル スタイルを作成することもできます。
* 範囲に適用するテキストの向きを設定する。
* 特定の範囲からブック内の別の場所または外部の場所にリンクするハイパーリンクを追加または変更する。

### <a name="manage-document-properties"></a>ドキュメント プロパティの管理

ドキュメント プロパティ API を使用して、組み込みのドキュメント プロパティにアクセスできます。また、ブックの状態を格納してワークフローやビジネス ロジックを操作するためのカスタム ドキュメント プロパティを作成、管理することもできます。

### <a name="copy-worksheets"></a>ワークシートのコピー

ワークシート コピー API を使用して、ワークシートのデータと書式設定を同じブック内の新しいワークシートにコピーできます。これにより、必要となるデータの転送量を削減することができます。

### <a name="handle-ranges-with-ease"></a>範囲の操作の簡易化

各種の範囲 API を使用して、周りの領域の取得や範囲のサイズ変更など、さまざまな操作を行うことができます。 これらの API により、範囲の操作やアドレス指定などのタスクが効率化されます。

さらに、次の機能も使用できます。

* ブックとワークシートの保護オプション: これらの API を使用して、ワークシートおよびブック構造内のデータを保護する。
* 名前付きアイテムの更新: この API を使用して、名前付きアイテムを更新する。
* アクティブ セルの取得: この API を使用して、ブックのアクティブ セルを取得する。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_プロパティ_ > chartType|グラフの種類を表します。 有効な値: ColumnClustered、ColumnStacked、ColumnStacked100、BarClustered、BarStacked、BarStacked100、LineStacked、LineStacked100、LineMarkers、LineMarkersStacked、LineMarkersStacked100、PieOfPie など。|1.7|
|[chart](/javascript/api/excel/excel.chart)|_プロパティ_ > id|グラフの一意の ID。 読み取り専用です。|1.7|
|[chart](/javascript/api/excel/excel.chart)|_プロパティ_ > showAllFieldButtons|ピボットグラフにすべてのフィールド ボタンを表示するかどうかを示します。|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_リレーションシップ_ > border|グラフ エリアの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_メソッド_ > getItem(type: string, group: string)|種類とグループで識別された特定の軸を返します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > axisBetweenCategories|項目の境界で数値軸が項目軸と交差するかどうかを表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > axisGroup|指定した軸に対応するグループを表します。 読み取り専用です。 有効な値は、Primary、Secondary です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > categoryType|項目軸の種類を返すか設定します。 有効な値は、Automatic、TextAxis、DateAxis です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > crosses|他の軸と交差する、指定の軸を表します。 有効な値は、Automatic、Maximum、Minimum、Custom です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > crossesAt|指定した軸と他の軸との交差位置を表します。 読み取り専用です。 このプロパティを設定するには、SetCrossesAt(double) メソッドを使用する必要があります。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > customDisplayUnit|カスタム軸の表示単位の値を表します。 読み取り専用です。 このプロパティを設定するには、SetCustomDisplayUnit(double) メソッドを使用してください。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > displayUnit|軸の表示単位を表します。 有効な値は、None、Hundreds、Thousands、TenThousands、HundredThousands、Millions、TenMillions、HundredMillions、Billions、Trillions、Custom です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > height|グラフ軸の高さ (ポイント数) を表します。 軸が表示されない場合は null 値です。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > left|軸の左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 軸が表示されない場合は null 値です。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > logBase|対数目盛りを使用する場合の対数の底を表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > reversePlotOrder|Microsoft Excel でデータ ポイントを最後から最初への順でプロットするかどうかを表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > scaleType|数値軸のスケールの種類を表します。 有効な値は、Linear、Logarithmic です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > showDisplayUnitLabel|軸の表示単位のラベルを表示するかどうかを表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > tickLabelSpacing|目盛ラベル間の項目または系列の数を表します。 1 から 31999 の範囲内で値を設定できます。自動的に設定する場合は、空の文字列にします。 戻り値は常に数値です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > tickMarkSpacing|目盛間の項目または系列の数を表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > top|軸の上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 軸が表示されない場合は null 値です。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > type|軸の種類を表します。 読み取り専用です。 有効な値は、Invalid、Category、Value、Series です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > visible|軸を表示するかどうかを表すブール値。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > width|グラフ軸の幅 (ポイント数) を表します。 軸が表示されない場合は null 値です。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_リレーションシップ_ > baseTimeUnit|指定された項目軸の基本単位を返すか設定します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_リレーションシップ_ > majorTickMark|指定した軸の目盛の種類を表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_リレーションシップ_ > majorTimeUnitScale|CategoryType プロパティが lTimeScale に設定されている場合、項目軸の目盛のスケール値を返すか設定します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_リレーションシップ_ > minorTickMark|指定した軸の補助目盛の種類を表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_リレーションシップ_ > minorTimeUnitScale|CategoryType プロパティが lTimeScale に設定されている場合、項目軸の補助目盛のスケール値を返すか設定します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_リレーションシップ_ > tickLabelPosition|指定した軸での目盛ラベルの位置を示します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_メソッド_ > setCategoryNames(sourceData: Range)|指定した軸のすべてのカテゴリ名を設定します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_メソッド_ > setCrossesAt(value: double)|指定した軸と他の軸との交差位置を設定します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_メソッド_ > setCustomDisplayUnit(value: double)|軸の表示単位をカスタム値に設定します。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_プロパティ_ > color|グラフの罫線の色を表す HTML カラー コード。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_プロパティ_ > weight|罫線の太さ (ポイント数) を表します。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_リレーションシップ_ > lineStyle|罫線のスタイルを表します。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > position|データ ラベルの位置を表すDataLabelPosition 値。使用可能な値は次のとおりです。None、Center、InsideEnd、InsideBase、OutsideEnd、Left、Right、Top、Bottom、BestFit、Callout。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > separator|グラフのデータ ラベルに使用される区切り文字を表す文字列。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showBubbleSize|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール値。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showCategoryName|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール値。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showLegendKey|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール値。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showPercentage|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール値。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showSeriesName|データ ラベルの系列名を表示するか非表示にするかを表すブール値。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showValue|データ ラベルの値を表示するか非表示にするかを表すブール値。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_プロパティ_ > height|グラフに表示される凡例の高さを表します。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_プロパティ_ > left|グラフ凡例の左を表します。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_プロパティ_ > showShadow|グラフで凡例を影付きにするかどうかを表します。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_プロパティ_ > top|グラフ凡例の上を表します。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_プロパティ_ > width|グラフに表示される凡例の幅を表します。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_リレーションシップ_ > legendEntries|凡例に含まれる凡例エントリのコレクションを表します。 読み取り専用です。|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_ > visible|グラフの凡例エントリを表示するかどうかを表します。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_プロパティ_ > items|chartLegendEntry オブジェクトのコレクションです。 読み取り専用です。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_メソッド_ > getCount()|コレクションに含まれる凡例エントリの数を返します。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_メソッド_ > getItemAt(index: number)|指定されたインデックスに位置する凡例エントリを返します。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_プロパティ_ > hasDataLabel|データ ポイントのデータ ラベルの有無を表します。 等高線グラフには適用されません。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_プロパティ_ > markerBackgroundColor|データ ポイントのマーカー背景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_プロパティ_ > markerForegroundColor|データ ポイントのマーカー前景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_プロパティ_ > markerSize|データ ポイントのマーカー サイズを表します。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_プロパティ_ > markerStyle|データ ポイントのマーカー スタイルを表します。 有効な値は、Invalid、Automatic、None、Square、Diamond、Triangle、X、Star、Dot、Dash、Circle、Plus、Picture です。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_リレーションシップ_ > dataLabel|グラフ データ ポイントのデータ ラベルを返します。 読み取り専用です。|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_リレーションシップ_ > border|グラフ データ ポイントの罫線の書式設定 (色、スタイル、線の太さなどの情報) を表します。 読み取り専用です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > chartType|グラフ系列の種類を表します。 有効な値: ColumnClustered、ColumnStacked、ColumnStacked100、BarClustered、BarStacked、BarStacked100、LineStacked、LineStacked100、LineMarkers、LineMarkersStacked、LineMarkersStacked100、PieOfPie など。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > doughnutHoleSize|グラフ系列のドーナツの穴の大きさを表します。  ドーナツ グラフと doughnutExploded グラフでのみ有効です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > filtered|データ系列がフィルター処理されるかどうかを表すブール値。 等高線グラフには適用されません。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > gapWidth|グラフ系列間に設けられる間隔を表します。  横棒グラフと縦棒グラフでのみ有効です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > hasDataLabels|系列のデータ ラベルの有無を表すブール値。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > markerBackgroundColor|グラフ系列のマーカー背景色を表します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > markerForegroundColor|グラフ系列のマーカー前景色を表します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > markerSize|グラフ系列のマーカー サイズを表します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > markerStyle|グラフ系列のマーカー スタイルを表します。 有効な値は、Invalid、Automatic、None、Square、Diamond、Triangle、X、Star、Dot、Dash、Circle、Plus、Picture です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > plotOrder|グラフ グループ内でのグラフ系列のプロット順序を表します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > showShadow|系列が影付きであるかどうかを表すブール値。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > smooth|系列が平滑化されるかどうかを表すブール値。 折れ線グラフと散布図にのみ適用されます。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_リレーションシップ_ > dataLabels|系列に含まれるすべてのデータ ラベルのコレクションを表します。 読み取り専用です。|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_リレーションシップ_ > trendlines|データ系列に含まれる近似曲線のコレクションを表します。 読み取り専用です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_メソッド_ > delete()|グラフ系列を削除します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_メソッド_ > setBubbleSizes(sourceData: Range)|グラフ系列のバブル サイズを設定します。 バブル チャートにのみ適用されます。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_メトリック_ > setValues(sourceData: Range)|グラフ系列の値を設定します。 散布図の場合、Y 軸の値を意味します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_メソッド_ > setXAxisValues(sourceData: Range)|グラフ系列の X 軸の値を設定します。 散布図にのみ適用されます。|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_メソッド_ > add(name: string, index: number)|コレクションに新しい系列を追加します。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_プロパティ_ > height|グラフ タイトルの高さ (ポイント数) を返します。 読み取り専用です。 グラフ タイトルが表示されない場合は null 値です。 読み取り専用です。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_プロパティ_ > horizontalAlignment|グラフ タイトルの水平方向の配置を表します。 有効な値は、Center、Left、Justify、Distributed、Right です。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_プロパティ_ > left|グラフ タイトルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフ タイトルが表示されない場合は null 値です。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_プロパティ_ > position|グラフ タイトルの位置を表します。 有効な値は、Top、Automatic、Bottom、Right、Left です。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_プロパティ_ > showShadow|グラフ タイトルが影付きにされるかどうかを指定するブール値を表します。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_プロパティ_ > textOrientation|グラフ タイトルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_プロパティ_ > top|グラフ タイトルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフ タイトルが表示されない場合は null 値です。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_プロパティ_ > verticalAlignment|グラフ タイトルの垂直方向の配置を表します。 有効な値は、Center、Bottom、Top、Justify、Distributed です。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_プロパティ_ > width|グラフ タイトルの幅 (ポイント数) を返します。 読み取り専用です。 グラフ タイトルが表示されない場合は null 値です。 読み取り専用です。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_メソッド_ > setFormula(formula: string)|A1 スタイルの表記法を使用するグラフ タイトルの数式を表す文字列値を設定します。|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_リレーションシップ_ > border|グラフ タイトルの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > backward|近似曲線を後方へ拡張するときの区間数を表します。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > displayEquation|true の場合、グラフに近似曲線の数式が表示されます。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > displayRSquared|true の場合、グラフに近似曲線の R-2 乗値が表示されます。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > forward|近似曲線を前方へ拡張するときの区間数を表します。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > intercept|近似曲線の切片の値を表します。 数値または空の文字列を設定できます (値を自動的に設定する場合)。 戻り値は常に数値です。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > movingAveragePeriod|グラフの近似曲線の期間を表します。MovingAverage 型の近似曲線にのみ適用されます。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > name|近似曲線の名前を表します。 文字列値または null 値 (値を自動的に設定する場合) に設定できます。 戻り値は常に文字列です。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > polynomialOrder|グラフの近似曲線の順序を表します。Polynomial 型の近似曲線にのみ適用されます。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > type|グラフの近似曲線の種類を表します。 有効な値は、Linear、Exponential、Logarithmic、MovingAverage、Polynomial、Power です。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_リレーションシップ_ > format|グラフの近似曲線の書式設定を表します。 読み取り専用です。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_メソッド_ > delete()|trendline オブジェクトを削除します。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_プロパティ_ > items|chartTrendline オブジェクトのコレクション。 読み取り専用です。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_メソッド_ > add(type: string)|近似曲線のコレクションに新しい近似曲線を追加します。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_メソッド_ > getCount()|コレクションに含まれる近似曲線の数を返します。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_メソッド_ > getItem(index: number)|インデックス (項目配列内の挿入順序) に基づいて trendline オブジェクトを取得します。|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_リレーションシップ_ > line|グラフの線の書式設定を表します。読み取り専用。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_プロパティ_ > key|カスタム プロパティのキーを取得します。読み取り専用。読み取り専用。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_プロパティ_ > type|カスタム プロパティの値の型を取得します。 読み取り専用。 読み取り専用です。 有効な値は、Number、Boolean、Date、String、Float です。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_プロパティ_ > value|カスタム プロパティの値を取得または設定します。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_メソッド_ > delete()|カスタム プロパティを削除します。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_プロパティ_ > items|customProperty オブジェクトのコレクション。読み取り専用。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_メソッド_ > add(key: string, value: object)|新しいカスタム プロパティを作成、または既存のカスタム プロパティを設定します。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_メソッド_ > deleteAll()|このコレクション内のすべてのカスタム プロパティを削除します。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_メソッド_ > getCount()|カスタム プロパティの数を取得します。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_メソッド_ > getItem(key: string)|キーに基づいてカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。カスタム プロパティが存在しない場合はエラーをスローします。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_メソッド_ > getItemOrNullObject(key: string)|キーに基づいてカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。カスタム プロパティが存在しない場合は null オブジェクトを返します。|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_プロパティ_ > items|dataConnection オブジェクトのコレクション。 読み取り専用です。|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_メソッド_ > refreshAll()|コレクションに含まれるすべてのデータ接続を更新します。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > author|ブックの作成者を取得または設定します。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > category|ブックのカテゴリを取得または設定します。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > comments|ブックのコメントを取得または設定します。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > company|ブックの会社を取得または設定します。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > keywords|ブックのキーワードを取得または設定します。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > lastAuthor|ブックの最後の作成者を取得します。 読み取り専用です。 読み取り専用です。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > manager|ブックのマネージャーを取得または設定します。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > revisionNumber|ブックのリビジョン番号を取得します。 読み取り専用です。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > subject|ブックの件名を取得または設定します。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > title|ブックのタイトルを取得または設定します。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_リレーションシップ_ > creationDate|ブックの作成日を取得します。 読み取り専用です。 読み取り専用です。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_リレーションシップ_ > custom|ブックのカスタム プロパティのコレクションを取得します。 読み取り専用です。 読み取り専用です。|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_プロパティ_ > formula|名前付きのアイテムの数式を取得または設定します。  数式は常に '=' 記号で始まります。|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_リレーションシップ_ > arrayValues|名前付きアイテムの値と型を含むオブジェクトを返します。 読み取り専用です。|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_プロパティ_ > types|名前付きアイテムの配列に含まれる各アイテムの型を表します。読み取り専用。 有効な値は、Unknown、Empty、String、Integer、Double、Boolean、Error です。|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_プロパティ_ > values|名前付きアイテムの配列に含まれる各アイテムの値を表します。読み取り専用。 読み取り専用です。|1.7|
|[range](/javascript/api/excel/excel.range)|_プロパティ_ > isEntireColumn|現在の範囲が列全体であるかどうかを表します。 読み取り専用です。|1.7|
|[range](/javascript/api/excel/excel.range)|_プロパティ_ > isEntireRow|現在の範囲が行全体であるかどうかを表します。 読み取り専用です。|1.7|
|[range](/javascript/api/excel/excel.range)|_プロパティ_ > numberFormatLocal|ユーザーの言語で文字列として指定された範囲に対応する、Excel の数値形式コードを表します。|1.7|
|[range](/javascript/api/excel/excel.range)|_プロパティ_ > style|現在の範囲のスタイルを表します。 null 値または文字列が返されます。|1.7|
|[range](/javascript/api/excel/excel.range)|_メソッド_ > getAbsoluteResizedRange(numRows: number, numColumns: number)|現在の Range オブジェクトと左上のセルが同じで、指定した数の行と列を含む Range オブジェクトを取得します。|1.7|
|[range](/javascript/api/excel/excel.range)|_メソッド_ > getImage()|base64 でエンコードされた画像として範囲をレンダリングします。|1.7|
|[range](/javascript/api/excel/excel.range)|_メソッド_ > getSurroundingRegion()|指定された範囲の左上のセルを囲む領域を表す Range オブジェクトを返します。 周囲の領域は、この範囲に相対の空白の行と空白の列の任意の組み合わせで囲まれた範囲です。|1.7|
|[range](/javascript/api/excel/excel.range)|_メソッド_ > showCard()|アクティブ セルに多数の値が含まれる場合、そのセルのカードを表示します。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_プロパティ_ > textOrientation|該当する範囲内のすべてのセルのテキストの向きを設定します。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_プロパティ_ > useStandardHeight|Range オブジェクトの行の高さを、シートの標準の高さと等しくするかどうかを指定します。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_プロパティ_ > useStandardWidth|Range オブジェクトの列の幅を、シートの標準の幅と等しくするかどうかを指定します。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_プロパティ_ > address|ハイパーリンクの URL ターゲットを表します。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_プロパティ_ > document|ハイパーリンクのドキュメント  ターゲットを表します。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_プロパティ_ > screenTip|ハイパーリンクの上にカーソルを合わせると表示される文字列を表します。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_プロパティ_ > textToDisplay|該当する範囲内の左上端のセルに表示される文字列を表します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > addIndent|セル内のテキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > autoIndent|セル内のテキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > builtIn|スタイルが組み込みのスタイルであるかどうかを示します。 読み取り専用です。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > formulaHidden|ワークシートが保護されている場合、数式を非表示にするかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > horizontalAlignment|スタイルでの水平方向の配置を表します。 有効な値は、General、Left、Center、Right、Fill、Justify、CenterAcrossSelection、Distributed です。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includeAlignment|スタイルに配置のプロパティ (AddIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel、および TextOrientation) が含まれるかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includeBorder|スタイルに罫線のプロパティ (Color、ColorIndex、LineStyle、Weight) が含まれているかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includeFont|スタイルにフォントのプロパティ (Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscript、Superscript、Underline) が含まれているかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includeNumber|スタイルに NumberFormat プロパティが含まれているかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includePatterns|スタイルに塗りつぶしのプロパティ (Color、ColorIndex、InvertIfNegative、Pattern、PatternColor、PatternColorIndex) が含まれているかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includeProtection|スタイルに保護のプロパティ (FormulaHidden および Locked) が含まれているかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > indentLevel|スタイルのインデント レベルを示す 0 から 250 の範囲内の整数。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > locked|ワークシートが保護されている場合、オブジェクトがロックされるかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > name|スタイルの名前。 読み取り専用です。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > numberFormat|スタイルで適用される数値形式の表示形式コード。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > numberFormatLocal|スタイルで適用される数値形式のローカライズされた表示形式コード。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > orientation|スタイルで適用されるテキストの向き。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > readingOrder|スタイルで適用される読み上げ順序。 有効な値は、Context、LeftToRight、RightToLeft です。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > shrinkToFit|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > textOrientation|スタイルで適用されるテキストの向き。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > verticalAlignment|スタイルで適用される垂直方向の配置を表します。 有効な値は、Top、Center、Bottom、Justify、Distributed です。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > wrapText|Microsoft Excel でオブジェクト内のテキストをラップするかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_リレーションシップ_ > borders|4 つの辺の罫線のスタイルを表す、4 つの Border オブジェクトのコレクション。 読み取り専用です。|1.7|
|[style](/javascript/api/excel/excel.style)|_リレーションシップ_ > fill|スタイルの塗りつぶし。 読み取り専用です。|1.7|
|[style](/javascript/api/excel/excel.style)|_リレーションシップ_ > font|スタイルのフォントを表す Font オブジェクト。 読み取り専用です。|1.7|
|[style](/javascript/api/excel/excel.style)|_メソッド_ > delete()|このスタイルを削除します。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_プロパティ_ > items|style オブジェクトのコレクション。 読み取り専用です。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_メソッド_ > add(name: string)]|コレクションに新しいスタイルを追加します。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_メソッド_ > getItem(name: string)|名前に基づいてスタイルを取得します。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_ > address|特定のワークシート上のテーブル内で変更されたエリアを表すアドレスを取得します。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_ > changeType|Changed イベントがトリガーされる方法を表す変更の種類を取得します。 有効な値は、Others、RangeEdited、RowInserted、RowDeleted、ColumnInserted、ColumnDeleted、CellInserted、CellDeleted です。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_ > source|イベントのソースを取得します。 有効な値は、Local、Remote です。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_ > tableId|データが変更されたテーブルの ID を取得します。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_ > type|イベントの種類を取得します。 有効な値は、WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted です。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_ > worksheetId|データが変更されたワークシートの ID を取得します。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_プロパティ_ > address|特定のワークシート上のテーブル内で選択されたエリアを表す範囲のアドレスを取得します。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_プロパティ_ > isInsideTable|選択範囲がテーブル内に収まっているかどうかを示します。IsInsideTable が false の場合、アドレスは無効です。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_プロパティ_ > tableId|選択範囲が変更されたテーブルの ID を取得します。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_プロパティ_ > type|イベントの種類を取得します。 有効な値は、WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted です。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_プロパティ_ > worksheetId|選択範囲が変更されたワークシートの ID を取得します。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_プロパティ_ > name|ブックの名前を取得します。 読み取り専用です。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > dataConnections|ブックに含まれるすべてのデータ接続を更新します。 読み取り専用です。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > properties|ブックのプロパティを取得します。 読み取り専用です。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > protection|ブックの workbookProtection オブジェクトを返します。 読み取り専用です。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > styles|ブックに関連付けられているスタイルのコレクションを表します。 読み取り専用です。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_メソッド_ > getActiveCell()|ブックで現在アクティブなセルを取得します。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_プロパティ_ > protected|ブックが保護されているかどうかを示します。 読み取り専用です。 読み取り専用です。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_メソッド_ > protect(password: string)|ブックを保護します。 ブックが保護されている場合は失敗します。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_メソッド_ > unprotect(password: string)|ブックの保護を解除します。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_プロパティ_ > gridlines|ワークシートの gridlines フラグを取得または設定します。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_プロパティ_ > headings|ワークシートの headings フラグを取得または設定します。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_プロパティ_ > showHeadings|ワークシートの headings フラグを取得または設定します。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_プロパティ_ > standardHeight|ワークシート内のすべての行の標準 (既定) の高さ (ポイント数) を返します。 読み取り専用です。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_プロパティ_ > standardWidth|ワークシートのすべての列の標準 (既定) の幅を返すか設定します。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_プロパティ_ > tabColor|ワークシートのタブの色を取得または設定します。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_リレーションシップ_ > freezePanes|ワークシートで固定されたウィンドウを操作するために使用できるオブジェクトを取得します。読み取り専用。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_メソッド_ > copy(positionType: WorksheetPositionType, relativeTo: Worksheet)|ワークシートをコピーして、指定した位置に配置します。 コピーしたワークシートを返します。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_メソッド_ > getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)|特定の行インデックスと列インデックスから開始し、一定数の行と列にわたる、Range オブジェクトを取得します。|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_プロパティ_ > type|イベントの種類を取得します。 有効な値は、WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted です。|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_プロパティ_ > worksheetId|アクティブにされたワークシートの ID を取得します。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_プロパティ_ > source|イベントのソースを取得します。 有効な値は、Local、Remote です。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_プロパティ_ > type|イベントの種類を取得します。 有効な値は、WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted です。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_プロパティ_ > worksheetId|ブックに追加されたワークシートの ID を取得します。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_プロパティ_ > address|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_プロパティ_ > changeType|Changed イベントがトリガーされる方法を表す変更の種類を取得します。 有効な値は、Others、RangeEdited、RowInserted、RowDeleted、ColumnInserted、ColumnDeleted、CellInserted、CellDeleted です。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_プロパティ_ > source|イベントのソースを取得します。 有効な値は、Local、Remote です。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_プロパティ_ > type|イベントの種類を取得します。 有効な値は、WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted です。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_プロパティ_ > worksheetId|データが変更されたワークシートの ID を取得します。|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_プロパティ_ > type|イベントの種類を取得します。 有効な値は、WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted です。|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_プロパティ_ > worksheetId|非アクティブにされたワークシートの ID を取得します。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_プロパティ_ > source|イベントのソースを取得します。 有効な値は、Local、Remote です。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_プロパティ_ > type|イベントの種類を取得します。 有効な値は、WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted です。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_プロパティ_ > worksheetId|ブックから削除されたワークシートの ID を取得します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > freezeAt(frozenRange: Range or string)|アクティブなワークシート ビューに固定セルを設定します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > freezeColumns(count: number)|ワークシートの最初の列 (複数可) を所定の場所に固定します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > freezeRows(count: number)|ワークシートの最初の行 (複数可) を所定の場所に固定します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > getLocation()|アクティブなワークシート ビュー内の固定セルを記述する範囲を取得します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > getLocationOrNullObject()|アクティブなワークシート ビュー内の固定セルを記述する範囲を取得します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > unfreeze()|ワークシートからすべての固定ウィンドウを削除します。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowEditObjects|オブジェクトの編集を可能にするワークシート保護オプションを表します。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowEditScenarios|シナリオの編集を可能にするワークシート保護オプションを表します。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_リレーションシップ_ > selectionMode|選択モードのワークシート保護オプションを表します。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_プロパティ_ > address|特定のワークシートで選択されたエリアを表す範囲のアドレスを取得します。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_プロパティ_ > type|イベントの種類を取得します。 有効な値は、WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted です。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_プロパティ_ > worksheetId|選択範囲が変更されたワークシートの ID を取得します。|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>Excel JavaScript API 1.6 の新機能 

### <a name="conditional-formatting"></a>条件付き書式

範囲の条件付き書式が導入されています。 次の種類の条件付き書式を使用できます。

* カラー スケール
* データ バー
* アイコン セット
* カスタム

さらに、次の機能も使用できます。

* 条件付き書式が適用された範囲を返す。 
* 条件付き書式を削除する。 
* 優先順位機能と stopifTrue 機能を提供する。 
* 指定した範囲のすべての条件付き書式のコレクションを取得する。 
* 現在指定している範囲でアクティブなすべての条件付き書式をクリアする。 

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[アプリケーション](/javascript/api/excel/excel.application)|_メソッド_ > suspendApiCalculationUntilNextSync()|次の "context.sync()" が呼び出されるまで、計算を中断します。設定されると、依存関係が確実に伝達されるようにブックを再計算するのは開発者の責任です。|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_リレーションシップ_ > format|Format オブジェクト (条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用。|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_リレーションシップ_ > rule|この条件付き書式の Rule オブジェクトを表します。|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_プロパティ_ > threeColorScale|true の場合、カラー スケールのポイントは 3 つ (最小、中間値、最大) になり、それ以外の場合は 2 つ (最小、最大) になります。読み取り専用。|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_リレーションシップ_ > criteria|カラー スケールの条件。2 ポイントのカラー スケールを使用する場合、中間値はオプションです。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_プロパティ_ > formula1|条件付き書式ルールを評価するために必要な場合、数式。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_プロパティ_ > formula2|条件付き書式ルールを評価するために必要な場合、数式。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_プロパティ_ > operator|テキストの条件付き書式の演算子。有効な値は、Invalid、Between、NotBetween、EqualTo、NotEqualTo、GreaterThan、LessThan、GreaterThanOrEqual、LessThanOrEqual です。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_リレーションシップ_ > maximum|最大ポイントのカラー スケール条件。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_リレーションシップ_ > midpoint|カラー スケールが 3 色スケールの場合のカラー スケール条件の中間値。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_リレーションシップ_ > minimum|最小ポイントのカラー スケール条件。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_プロパティ_ > color|カラー スケールの色を表す HTML カラー コード。たとえば、#FF0000 は赤を表します。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_プロパティ_ > formula|数値、数式、(型が LowestValue の場合は) null。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_プロパティ_ > type|アイコンの条件式を適用する基準。有効な値は、Invalid、LowestValue、HighestValue、Number、Percent、Formula、Percentile です。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_プロパティ_ > borderColor|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_プロパティ_ > fillColor|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_プロパティ_ > matchPositiveBorderColor|負の DataBar に正の DataBar と同じ枠線の色があるかどうかを表すブール値。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_プロパティ_ > matchPositiveFillColor|負の DataBar に正の DataBar と同じ塗りつぶしの色があるかどうかを表すブール値。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_プロパティ_ > borderColor|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_プロパティ_ > fillColor|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_プロパティ_ > gradientFill|DataBar のグラデーションの有無を表すブール値。|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_プロパティ_ > formula|databar のルールを評価するために必要な場合、数式。|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_プロパティ_ > type|databar のルールの種類。有効な値は、LowestValue、HighestValue、Number、Percent、Formula、Percentile、Automatic です。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_プロパティ_ > id|現在の ConditionalFormatCollection 内での条件付き書式の優先順位。 読み取り専用です。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_プロパティ_ > priority|この条件付き書式が現在存在する、条件付き書式のコレクション内の優先順位 (またはインデックス)。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_プロパティ_ > stopIfTrue|この条件付き書式の条件が満たされた場合、優先順位の低い書式はそのセルに影響を及ぼしません。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_プロパティ_ > type|条件付き書式の種類。一度に 1 つのみ設定できます。読み取り専用。有効な値は、Custom、DataBar、ColorScale、IconSet です。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > cellValue|現在の条件付き書式が CellValue 型の場合、セル値の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > cellValueOrNullObject|現在の条件付き書式が CellValue 型の場合、セル値の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > colorScale|現在の条件付き書式が ColorScale 型の場合、ColorScale の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > colorScaleOrNullObject|現在の条件付き書式が ColorScale 型の場合、ColorScale の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > custom|現在の条件付き書式がカスタム型の場合、カスタムの条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > customOrNullObject|現在の条件付き書式がカスタム型の場合、カスタムの条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > dataBar|現在の条件付き書式がデータ バーの場合、データ バーのプロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > dataBarOrNullObject|現在の条件付き書式がデータ バーの場合、データ バーのプロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > iconSet|現在の条件付き書式が IconSet 型の場合、IconSet の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > iconSetOrNullObject|現在の条件付き書式が IconSet 型の場合、IconSet の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > preset|above averagebelow averageunique valuescontains blanknonblankerrornoerror properties などの事前に設定された条件付き書式を返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > presetOrNullObject|above averagebelow averageunique valuescontains blanknonblankerrornoerror properties などの事前に設定された条件付き書式を返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > textComparison|現在の条件付き書式がテキスト型の場合、特定のテキストの条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > textComparisonOrNullObject|現在の条件付き書式がテキスト型の場合、特定のテキストの条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > topBottom|現在の条件付き書式が TopBottom 型の場合、TopBottom の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_リレーションシップ_ > topBottomOrNullObject|現在の条件付き書式が TopBottom 型の場合、TopBottom の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_メソッド_ > delete()|この条件付き書式を削除します。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_メソッド_ > getRange()|条件付き書式が適用される範囲を返します。範囲が連続していない場合は null オブジェクトを返します。読み取り専用。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_メソッド_ > getRangeOrNullObject()|条件付き書式が適用される範囲を返します。範囲が連続していない場合は null オブジェクトを返します。読み取り専用。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_プロパティ_ > items|conditionalFormat オブジェクトのコレクション。読み取り専用。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_メソッド_ > add(type: string)|最高度の優先順位でコレクションに新しい条件付き書式を追加します。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_メソッド_ > clearAll()|現在指定している範囲でアクティブなすべての条件付き書式をクリアする。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_メソッド_ > getCount()|ブックに含まれる条件付き書式の数を返します。読み取り専用。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_メソッド_ > getItem(id: string)|指定された ID に対応する条件付き書式を返します。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_メソッド_ > getItemAt(index: number)|指定されたインデックスに条件付き書式を返します。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_プロパティ_ > formula|条件付き書式ルールを評価するために必要な場合、数式。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_プロパティ_ > formulaLocal|ユーザーの言語で条件付き書式ルールを評価するために必要な場合、数式。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_プロパティ_ > formulaR1C1|R1C1 形式の表記法で条件付き書式ルールを評価するために必要な場合、数式。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_プロパティ_ > formula|種類によっては数値または数式。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_プロパティ_ > operator|アイコン条件付き書式のそれぞれのルールの種類に適用する演算子。有効な値は、Invalid、GreaterThan、GreaterThanOrEqual です。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_リレーションシップ_ > customIcon|既定の IconSet と異なる場合は現在の条件のカスタム アイコン、そうでない場合は null が返されます。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_リレーションシップ_ > type|アイコンの条件式は次のものに基づいています。|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_プロパティ_ > criterion|条件付き書式の条件です。使用可能な値は次のとおりです。Invalid、Blanks、NonBlanks、Errors、NonErrors、Yesterday、Today、Tomorrow、LastSevenDays、LastWeek、ThisWeek、NextWeek、LastMonth、ThisMonth、NextMonth、AboveAverage、BelowAverage、EqualOrAboveAverage、EqualOrBelowAverage、OneStdDevAboveAverage、OneStdDevBelowAverage、TwoStdDevAboveAverage、TwoStdDevBelowAverage、ThreeStdDevAboveAverage、ThreeStdDevBelowAverage、UniqueValues、DuplicateValues。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_プロパティ_ > color|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_プロパティ_ > id|罫線の識別子を表します。読み取り専用。有効な値は、EdgeTop、EdgeBottom、EdgeLeft、EdgeRight です。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_プロパティ_ > sideIndex|罫線の特定の辺を表す定数値。読み取り専用。有効な値は、EdgeTop、EdgeBottom、EdgeLeft、EdgeRight です。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_プロパティ_ > style|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。使用可能な値は次のとおりです。None、Continuous、Dash、DashDot、DashDotDot、Dot、Double、SlantDashDot。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_プロパティ_ > count|コレクションに含まれる border オブジェクトの数。読み取り専用。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_プロパティ_ > items|conditionalRangeBorder オブジェクトのコレクション。読み取り専用。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_リレーションシップ_ > bottom|上罫線を取得します。読み取り専用です。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_リレーションシップ_ > left|上罫線を取得します。読み取り専用です。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_リレーションシップ_ > right|上罫線を取得します。読み取り専用です。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_リレーションシップ_ > top|上罫線を取得します。読み取り専用です。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_メソッド_ > getItem(index: string)|名前に基づいて border オブジェクトを取得します。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_メソッド_ > getItemAt(index: number)|インデックスに基づいて border オブジェクトを取得します。|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_プロパティ_ > color|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_メソッド_ > clear()|塗りつぶしをリセットします。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_プロパティ_ > bold|フォントの太字の状態を表します。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_プロパティ_ > color|テキストの色を表す HTML カラー コード。たとえば、#FF0000 は赤を表します。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_プロパティ_ > italic|フォントの斜体の状態を表します。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_プロパティ_ > strikethrough|フォントの取り消し線の状態を表します。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_プロパティ_ > underline|フォントに適用する下線の種類。有効な値は、None、Single、Double です。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_メソッド_ > clear()|フォントの書式設定をリセットします。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_プロパティ_ > numberFormat|指定した範囲の Excel の数値書式コードを表します。null 値が渡された場合はクリアされます。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_リレーションシップ_ > borders|条件付き書式範囲全体に適用される border オブジェクトのコレクション。読み取り専用。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_リレーションシップ_ > fill|条件付き書式範囲全体に定義された fill オブジェクトを返します。読み取り専用。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_リレーションシップ_ > font|条件付き書式範囲全体に定義された font オブジェクトを返します。読み取り専用。|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_プロパティ_ > operator|テキストの条件付き書式の演算子。有効な値は、Invalid、Contains、NotContains、BeginsWith、EndsWith です。|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_プロパティ_ > text|条件付き書式のテキスト値。|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_プロパティ_ > rank|数値のランクに対する 1 から 1000、またはパーセントのランクに対する 1 から 100 のランク。|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_プロパティ_ > type|上位または下位のランクに基づく値を書式設定します。有効な値は、Invalid、TopItems、TopPercent、BottomItems、BottomPercent です。|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_リレーションシップ_ > format|Format オブジェクト (条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用。|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_リレーションシップ_ > rule|この条件付き書式の Rule オブジェクトを表します。読み取り専用。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_プロパティ_ > axisColor|軸の線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_プロパティ_ > axisFormat|Excel のデータ バーの軸を決定する方法を表します。有効な値は、Automatic、None、CellMidPoint です。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_プロパティ_ > barDirection|データ バーの図の基準となる方向を表します。有効な値は、Context、LeftToRight、RightToLeft です。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_プロパティ_ > showDataBarOnly|true の場合、データ バーが適用されているセルの値を非表示にします。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_リレーションシップ_ > lowerBoundRule|データ バーの下限値 (および該当する場合はその計算方法) を構成するルール。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_リレーションシップ_ > negativeFormat|Excel データ バーの軸の左側のすべての値を表します。読み取り専用。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_リレーションシップ_ > positiveFormat|Excel データ バーの軸の右側のすべての値を表します。読み取り専用。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_リレーションシップ_ > upperBoundRule|データ バーの上限値 (および該当する場合はその計算方法) を構成するルール。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_プロパティ_ > reverseIconOrder|true の場合、IconSet のアイコンの順序を反転します。カスタム アイコンを使用する場合、この設定は適用できません。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_プロパティ_ > showIconOnly|true の場合、値は非表示にされて、アイコンのみが表示されます。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_プロパティ_ > style|設定した場合、条件付き書式の IconSet オプションを表示します。使用可能な値は次のとおりです。Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_リレーションシップ_ > criteria|条件付きアイコンの規則と潜在的なカスタム アイコンの、抽出条件と IconSets の配列。最初の条件として、カスタム アイコンのみを変更することができますが、設定された場合に、型、数式、演算子は無視されます。|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_リレーションシップ_ > format|Format オブジェクト (条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用。|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_リレーションシップ_ > rule|条件付き書式のルール。|1.6|
|[range](/javascript/api/excel/excel.range)|_リレーションシップ_ > conditionalFormats|範囲を交差する ConditionalFormats のコレクション。読み取り専用。|1.6|
|[range](/javascript/api/excel/excel.range)|_メソッド_ > calculate()|ワークシート上のセルの範囲を計算します。|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_リレーションシップ_ > format|Format オブジェクト (条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用。|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_リレーションシップ_ > rule|条件付き書式のルール。|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_リレーションシップ_ > format|Format オブジェクト (条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用。|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_リレーションシップ_ > rule|TopBottom の条件付き書式の条件。|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > internalTest|内部使用のみ。読み取り専用。|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_メソッド_ > calculate(markAllDirty: bool)|ワークシート上のすべてのセルを計算します。|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>Excel JavaScript API 1.5 の新機能

### <a name="custom-xml-part"></a>カスタム XML パーツ

* ブック オブジェクトにカスタム XML パーツのコレクションを追加します。
* ID を使用したカスタム XML パーツの取得
* 名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。
* パーツに関連付けられている XML 文字列を取得します。
* パーツの ID と名前空間を指定します。
* ブックに新しいカスタム XML 部分を追加します。
* XML パーツ全体を設定します。
* カスタム XML パーツを削除します。
* xpath で識別される要素から、指定された名前を持つ属性を削除します。
* xpath で XML の内容を照会します。
* 属性を挿入、更新、および削除します。

**参照実装:** アドインでカスタム XML パーツを使用する方法に関する参照実装については、[こちら](https://github.com/mandren/Excel-CustomXMLPart-Demo)を参照してください。

### <a name="others"></a>その他
* `range.getSurroundingRegion()` は、この範囲の周囲の領域を表す Range オブジェクトを返します。周囲の領域は、この範囲に相対の空白の行と空白の列の任意の組み合わせで囲まれた範囲です。
* テーブル列に対して `getNextColumn()` および `getPreviousColumn()`、`getLast() を指定します。
* ブックに対して `getActiveWorksheet()` を指定します。
* ブックから `getRange(address: string)` を指定解除します。
* `getBoundingRange(ranges: )` は、指定した範囲を包含する、最小の範囲オブジェクトを取得します。たとえば、"B2:C5" から "D10:E15" までの隣接する範囲は、"B2:E15" になります。
* コレクション内の項目の数を取得する、名前付きの項目、ワークシート、テーブルなどのさまざまなコレクションに対して `getCount()` を指定します。 `workbook.worksheets.getCount()`
* ワークシート、テーブル列、グラフのポイント、範囲ビュー コレクションなどのさまざまなコレクションに対して `getFirst()`、`getLast()`、および get last を指定します。
* ワークシート、テーブル列の各コレクションに対して `getNext()` および `getPrevious()` を指定します。
* `getRangeR1C1()` は、特定の行インデックスと列インデックスから開始し、一定数の行と列にわたる、range オブジェクトを取得します。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_プロパティ_ > id|カスタム XML パーツの ID。読み取り専用。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_プロパティ_ > namespaceUri|カスタム XML パーツの名前空間 URI。読み取り専用。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_メソッド_ > delete()|カスタム XML パーツを削除します。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_メソッド_ > getXml()|カスタム XML パーツのすべての XML コンテンツを取得します。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_メソッド_ > setXml(xml: string)|カスタム XML パーツのすべての XML コンテンツを設定します。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_プロパティ_ > items|customXmlPart オブジェクトのコレクション。読み取り専用。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_メソッド_ > add(xml: string)|ブックに新しいカスタム XML パーツを追加します。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_メソッド_ > getByNamespace(namespaceUri: string)|名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_メソッド_ > getCount()|コレクションに含まれる CustomXml パーツの数を取得します。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_メソッド_ > getItem(id: string)|ID に基づいて、カスタム XML パーツを取得します。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_メソッド_ > getItemOrNullObject(id: string)|ID に基づいて、カスタム XML パーツを取得します。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_プロパティ_ > items|customXmlPartScoped オブジェクトのコレクション。読み取り専用。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_メソッド_ > getCount()|コレクションに含まれる CustomXML パーツの数を取得します。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_メソッド_ > getItem(id: string)|ID に基づいて、カスタム XML パーツを取得します。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_メソッド_ > getItemOrNullObject(id: string)|ID に基づいて、カスタム XML パーツを取得します。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_メソッド_ > getOnlyItem()|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_メソッド_ > getOnlyItemOrNullObject()|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|1.5|
|[workbook](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > customXmlParts|このブックに含まれる、カスタム XML パーツのコレクションを表します。読み取り専用。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_メソッド_ > getNext(visibleOnly: bool)|このワークシートの後に続くワークシートを取得します。後続のワークシートがない場合、このメソッドによってエラーがスローされます。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_メソッド_ > getNextOrNullObject(visibleOnly: bool)|このワークシートの後に続くワークシートを取得します。後続のワークシートがない場合、このメソッドによって null オブジェクトがスローされます。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_メソッド_ > getPrevious(visibleOnly: bool)|このワークシートに先行するワークシートを取得します。先行するワークシートがない場合、このメソッドによってエラーがスローされます。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_メソッド_ > getPreviousOrNullObject(visibleOnly: bool)|このワークシートに先行するワークシートを取得します。先行するワークシートがない場合、このメソッドによって null オブジェクトがスローされます。|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_メソッド_ > getFirst(visibleOnly: bool)|コレクション内の最初のワークシートを取得します。|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_メソッド_ > getLast(visibleOnly: bool)|コレクション内の最後のワークシートを取得します。|1.5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Excel JavaScript API 1.4 の新機能
要件セット 1.4 の Excel JavaScript API に新たに追加された機能は次のとおりです。

### <a name="named-item-add-and-new-properties"></a>名前付きアイテムの追加と新しいプロパティ

新しいプロパティ:

* `comment`
* `scope` ワークシートまたはブックの対象になるアイテム
* `worksheet` 名前付きアイテムの対象になるワークシートを返します。

新しいメソッド:

* `add(name: string, reference: Range or string, comment: string)` は、新しい名前を指定したスコープのコレクションに追加します。
* `addFormulaLocal(name: string, formula: string, comment: string)` は、ユーザーのロケールを数式に使用して、新しい名前を指定したスコープのコレクションに追加します。

### <a name="settings-api-in-the-excel-namespace"></a>Excel 名前空間内の Setting API

[Setting](/javascript/api/excel/excel.setting) オブジェクトは、ドキュメントに永続的に適用される設定のキーと値のペアを表します。 `Excel.Setting` の機能は `Office.Settings` と同等ですが、共通 API のコールバック モデルではなくバッチ API 構文を使用します。

API には、キーを使用して設定エントリを取得するための `getItem()` と、指定したキーと値の設定のペアをワークブックに追加するための `add()` が含まれています。

### <a name="others"></a>その他

* テーブルの列名を設定します (以前のバージョンでは読み取りのみ可能)。
* テーブルの列をテーブルの末尾に追加します (以前のバージョンでは末尾以外の任意の場所のみ可能)。
* 一度に複数の行をテーブルに追加します (以前のバージョンでは一度に 1 行のみ可能)。
* `range.getColumnsAfter(count: number)` および `range.getColumnsBefore(count: number)` を使用して、現在の Range オブジェクトの左右にある特定の数の列を取得します。
* アイテムまたは null オブジェクト関数。この機能により、キーを使用してオブジェクトを取得できます。オブジェクトが存在しない場合、返されたオブジェクトの isNullObject プロパティは true になります。これにより、開発者は例外処理を通じてオブジェクトを処理する必要なしに、オブジェクトが存在するかどうかを確認することができます。ワークシート、名前付きアイテム、バインド、グラフの系列などで使用できます。

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_ > getCount()|コレクションに含まれるバインドの数を取得します。|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_ > getItemOrNullObject(id: string)|ID に基づいて binding オブジェクトを取得します。binding オブジェクトが存在しない場合は null オブジェクトを返します。|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_メソッド_ > getCount()|ワークシート上のグラフの数を返します。|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_メソッド_ > getItemOrNullObject(name: string)|名前に基づいてグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_メソッド_ > getCount()|系列に含まれるグラフのポイントの数を返します。|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_メソッド_ > getCount()|コレクションに含まれるデータ系列の数を返します。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_プロパティ_ > comment|この名前に関連付けられているコメントを表します。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_プロパティ_ > scope|名前がブックを対象にしているのか、特定のワークシートを対象にしているのかを示します。読み取り専用です。使用可能な値は次のとおりです。Equal、Greater、GreaterEqual、Less、LessEqual、NotEqual。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_リレーションシップ_ > worksheet|名前付きのアイテムの対象になるワークシートを返します。アイテムがブックを対象にしている場合は、エラーをスローします。読み取り専用です。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_リレーションシップ_ > worksheetOrNullObject|名前付きのアイテムの対象になるワークシートを返します。アイテムがブックを対象にしている場合は、null オブジェクトを返します。読み取り専用です。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_メソッド_ > delete()|指定された名前を削除します。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_メソッド_ > getRangeOrNullObject()|名前に関連付けられている range オブジェクトを返します。名前付きアイテムの型が範囲でない場合は、null オブジェクトを返します。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_メソッド_ > add(name: string, reference: Range or string, comment: string)|指定のスコープのコレクションに新しい名前を追加します。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_メソッド_ > addFormulaLocal(name: string, formula: string, comment: string)|ユーザーのロケールを数式に使用して、指定のスコープのコレクションに新しい名前を追加します。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_メソッド_ > getCount()|コレクションに含まれる名前付きアイテムの数を取得します。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_メソッド_ > getItemOrNullObject(name: string)|名前に基づいて nameditem オブジェクトを取得します。nameditem オブジェクトが存在しない場合は null オブジェクトを返します。|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_メソッド_ > getCount()|コレクションに含まれるピボット テーブルの数を取得します。|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_メソッド_ > getItemOrNullObject(name: string)|名前に基づいてピボットテーブルを取得します。ピボットテーブルが存在しない場合は null オブジェクトを返します。|1.4|
|[range](/javascript/api/excel/excel.range)|_メソッド_ > getIntersectionOrNullObject(anotherRange: Range or string)|指定した範囲の長方形の交差を表す範囲オブジェクトを取得します。 交差部分が見つからない場合は、null オブジェクトを返します。|1.4|
|[range](/javascript/api/excel/excel.range)|_メソッド_ > getUsedRangeOrNullObject(valuesOnly: bool)|指定した範囲オブジェクトのうち使用されている範囲を返します。範囲内に使用済みのセルがない場合、この関数は null オブジェクトを返します。|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_メソッド_ > getCount()|コレクションに含まれる RangeView オブジェクトの数を取得します。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_プロパティ_ > key|Setting の ID を表すキーを返します。読み取り専用。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_プロパティ_ > value|この設定に格納されている値を表します。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_メソッド_ > delete()|設定を削除します。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_プロパティ_ > items|setting オブジェクトのコレクション。読み取り専用。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_ > add(key: string, value: (any))|指定した設定をブックに設定または追加します。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_ > getCount()|コレクションに含まれる設定の数を取得します。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_ > getItem(key: string)|キーに基づいて設定エントリを取得します。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_ > getItemOrNullObject(key: string)|キーに基づいて設定エントリを取得します。設定が存在しない場合は null オブジェクトを返します。|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_リレーションシップ_ > settings|SettingsChanged イベントが発生したバインドを表す Setting オブジェクトを取得します。|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_メソッド_ > getCount()]|コレクションに含まれるテーブルの数を取得します。|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_メソッド_ > getItemOrNullObject(key: number or string)|名前または ID に基づいてテーブルを取得します。テーブルが存在しない場合は null オブジェクトを返します。|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_メソッド_ > getCount()|表の列数を取得します。|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_メソッド_ > getItemOrNullObject(key: number or string)|名前または ID に基づいて column オブジェクトを取得します。列が存在しない場合は null オブジェクトを返します。|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_メソッド_ > getCount()|表の行数を取得します。|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > settings|ブックに関連付けられている Setting のコレクションを表します。読み取り専用。|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_リレーションシップ_ > names|現在のワークシートにスコープされている名前のコレクション。読み取り専用です。|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_メソッド_ > getUsedRangeOrNullObject(valuesOnly: bool)|使用範囲とは、値または書式設定が割り当たっているすべてのセルを包含する最小の範囲です。ワークシート全体が空白の場合、この関数は null オブジェクトを返します。|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_メソッド_ > getCount(visibleOnly: bool)|コレクションに含まれるワークシートの数を取得します。|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_メソッド_ > getItemOrNullObject(key: string)|名前または ID に基づいて worksheet オブジェクトを取得します。ワークシートが存在しない場合は null オブジェクトを返します。|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 の新機能

要件セット 1.3 の Excel JavaScript API に新しく追加された点は次のとおりです。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_メソッド_ > delete()|バインドを削除します。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_ > add(range: Range or string, bindingType: string, id: string)|特定の範囲に新しいバインドを追加します。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_ > addFromNamedItem(name: string, bindingType: string, id: string)|ブック内の名前付きアイテムに基づいて新しいバインドを追加します。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_ > addFromSelection(bindingType: string, id: string)|現在の選択範囲に基づいて新しいバインドを追加します。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_ > getItemOrNull(id: string)|ID に基づいて binding オブジェクトを取得します。binding オブジェクトが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_メソッド_ > getItemOrNull(name: string)|名前に基づいてグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_メソッド_ > getItemOrNull(name: string)|名前に基づいて nameditem オブジェクトを取得します。nameditem オブジェクトが存在しない場合、返されるオブジェクトの isNull プロパティは true になります。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_プロパティ_ > name|ピボットテーブルの名前。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_リレーションシップ_ > worksheet|現在のピボットテーブルを含んでいるワークシート。読み取り専用。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_メソッド_ > refresh()|ピボットテーブルを更新します。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_プロパティ_ > items|ピボットテーブル オブジェクトのコレクション。読み取り専用。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_メソッド_ > getItem(name: string)|名前に基づいてピボットテーブルを取得します。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_メソッド_ > getItemOrNull(name: string)|名前に基づいてピボットテーブルを取得します。ピボットテーブルが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|1.3|
|[range](/javascript/api/excel/excel.range)|_メソッド_ > getIntersectionOrNull(anotherRange: Range or string)|指定した範囲の長方形の交差部分を表す Range オブジェクトを取得します。交差部分が見つからない場合は、null オブジェクトを返します。|1.3|
|[range](/javascript/api/excel/excel.range)|_メソッド_ > getVisibleView()|現在の範囲で表示される行を表します。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_プロパティ_ > cellAddresses|RangeView のセル アドレスを表します。読み取り専用。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_プロパティ_ > columnCount|表示されている列の数を返します。読み取り専用。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_プロパティ_ > formulas|A1 スタイル表記の数式を表します。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_プロパティ_ > formulasLocal|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, introduced in 1.5" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_プロパティ_ > formulasR1C1|R1C1 スタイル表記の数式を表します。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_プロパティ_ > index|RangeView のインデックスを表す値を返します。読み取り専用。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_プロパティ_ > numberFormat|指定したセルの Excel の数値書式コードを表します。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_プロパティ_ > rowCount|表示されている行の数を返します。読み取り専用。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_プロパティ_ > text|指定した範囲のテキスト値。テキスト値は、セルの幅には依存しません。Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。読み取り専用です。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_プロパティ_ > valueTypes|各セルのデータの種類を表します。読み取り専用です。使用可能な値は次のとおりです。Unknown、Empty、String、Integer、Double、Boolean、Error。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_プロパティ_ > values|指定した範囲ビューの Raw 値を表します。返されるデータの型は、文字列、数値、ブール値のいずれかになります。エラーが含まれているセルは、エラー文字列を返します。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_リレーションシップ_ > rows|範囲に関連付けられている範囲ビューのコレクションを表します。読み取り専用。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_メソッド_ > getRange()|現在の RangeView に関連付けられている親の範囲を取得します。|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_プロパティ_ > items|rangeView オブジェクトのコレクション。読み取り専用。|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_メソッド_ > getItemAt(index: number)|インデックスに基づいて範囲ビューの行番号を取得します。0 を起点とする番号になります。|1.3|
|[setting](/javascript/api/excel/excel.setting)|_プロパティ_ > key|Setting の ID を表すキーを返します。読み取り専用。|1.3|
|[setting](/javascript/api/excel/excel.setting)|_メソッド_ > delete()|設定を削除します。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_プロパティ_ > items|setting オブジェクトのコレクション。読み取り専用。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_ > getItem(key: string)|キーに基づいて設定エントリを取得します。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_ > getItemOrNull(key: string)|キーに基づいて設定エントリを取得します。設定が存在しない場合、返されるオブジェクトの isNull プロパティは true になります。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_ > set(key: string, value: string)|指定した設定をブックに設定または追加します。|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_リレーションシップ_ > settingCollection|SettingsChanged イベントが発生したバインドを表す Setting オブジェクトを取得します。|1.3|
|[table](/javascript/api/excel/excel.table)|_プロパティ_ > highlightFirstColumn|最初の列に特別な書式設定が含まれているかどうかを示します。|1.3|
|[table](/javascript/api/excel/excel.table)|_プロパティ_ > highlightLastColumn|最後の列に特別な書式設定が含まれているかどうかを示します。|1.3|
|[table](/javascript/api/excel/excel.table)|_プロパティ_ > showBandedColumns|テーブルを見やすくするため、奇数列を偶数列とは異なる方法で強調表示する書式設定にして、列を縞模様で表示するかどうかを示します。|1.3|
|[table](/javascript/api/excel/excel.table)|_プロパティ_ > showBandedRows|テーブルを見やすくするため、奇数行を偶数行とは異なる方法で強調表示する書式設定にして、行を縞模様で表示するかどうかを示します。|1.3|
|[table](/javascript/api/excel/excel.table)|_プロパティ_ > showFilterButton|フィルター ボタンを各列のヘッダーの上部に表示するかどうかを示します。これは、テーブルにヘッダー行が含まれている場合のみ設定できます。|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_メソッド_ > getItemOrNull(key: number or string)|名前または ID に基づいてテーブルを取得します。テーブルが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_メソッド_ > getItemOrNull(key: number or string)|名前または ID に基づいて column オブジェクトを取得します。列が存在しない場合、返されるオブジェクトの isNull プロパティは true になります。|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > pivotTables|ブックに関連付けられているピボットテーブルのコレクションを表します。読み取り専用。|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > settings|ブックに関連付けられている Setting のコレクションを表します。読み取り専用。|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_リレーションシップ_ > pivotTables|ワークシートの一部になっているピボットテーブルのコレクション。読み取り専用。|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Excel JavaScript API 1.2 の新機能

要件セット 1.2 の Excel JavaScript API に新たに追加された点は次のとおりです。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_プロパティ_ > id|コレクション内での位置を基にグラフを取得します。読み取り専用です。|1.2|
|[chart](/javascript/api/excel/excel.chart)|_リレーションシップ_ > worksheet|現在のグラフを含んでいるワークシート。読み取り専用。|1.2|
|[chart](/javascript/api/excel/excel.chart)|_メソッド_ > getImage(height: number, width: number, fittingMode: string)|指定したサイズに合わせてグラフを拡大、縮小することで、グラフを Base64 でエンコードされた画像としてレンダリングします。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_リレーションシップ_ > criteria|指定した列に現在適用されているフィルターです。読み取り専用です。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > apply(criteria: FilterCriteria)|指定の列に、指定したフィルター条件を適用します。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > applyBottomItemsFilter(count: number)|指定した数の要素の列に "下位アイテム" フィルターを適用します。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > applyBottomPercentFilter(percent: number)]|指定した割合の要素の列に "下位パーセント" フィルターを適用します。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > applyCellColorFilter(color: string)|指定した色の列に "セルの色" フィルターを適用します。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > applyCustomFilter(criteria1: string, criteria2: string, oper: string)|指定した条件の文字列の列に "アイコン" フィルターを適用します。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > applyDynamicFilter(criteria: string)|列に "動的" フィルターを適用します。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > applyFontColorFilter(color: string)|指定した色の列に "フォントの色" フィルターを適用します。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > applyIconFilter(icon: Icon)|指定したアイコンの列に "アイコン" フィルターを適用します。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > applyTopItemsFilter(count: number)|指定した数の要素の列に "上位アイテム" フィルターを適用します。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > applyTopPercentFilter(percent: number)|指定した割合の要素の列に "上位パーセント" フィルターを適用します。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > applyValuesFilter(values: ())|指定した値の列に "値" フィルターを適用します。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_メソッド_ > clear()|指定した列に適用されているフィルターをクリアします。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > color|セルをフィルター処理するために使用する HTML カラー文字列。「CellColor」フィルターおよび「fontColor」フィルターと併用します。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > criterion1|データをフィルター処理するために使用する最初の条件。「カスタム」フィルター処理の場合には、演算子として使用されます。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > criterion2|データをフィルター処理するために使用する 2 番目の条件。「カスタム」フィルター処理の場合には、演算子としてのみ使用されます。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ >dynamicCriteria|この列に適用する Excel.DynamicFilterCriteria の動的条件。「動的」フィルター処理で使用します。使用可能な値は次のいずれかです。Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > filterOn|値を表示したままにするかどうかを判別するために、フィルターで使用するプロパティ。使用可能な値は次のとおりです。BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > operator|"カスタム" フィルター処理を使用するときに、条件 1 と条件 2 と結合との使用する演算子。使用可能な値は次のとおりです。And、Or。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > values|"値" フィルター処理の一部として使用する値のセット。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_リレーションシップ_ > icon|セルをフィルター処理するために使用するアイコン。「アイコン」フィルター処理で使用します。|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_プロパティ_ > date|データのフィルター処理に使用する ISO8601 形式の日付です。|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_プロパティ_ > specificity|データを保持するのに、日付をどの程度詳細に使用するか。たとえば、date が 2005-04-02 で "month" に設定した場合、フィルター操作では 2005 年 4 月の日付データを含むすべての行が保持されます。使用可能な値は次のとおりです。Year、Month、Day、Hour、Minute、Second。|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_プロパティ_ > formulaHidden|Excel が範囲内のセルの数式を非表示にするかどうかを示します。null 値は、範囲全体に一様な数式非表示設定がないことを表します。|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_プロパティ_ > locked|Excel がオブジェクト内のセルをロックするかどうかを示します。null 値は、範囲全体に一様なロック設定がないことを表します。|1.2|
|[icon](/javascript/api/excel/excel.icon)|_プロパティ_ > index|指定したセット内のアイコンのインデックスを表します。|1.2|
|[icon](/javascript/api/excel/excel.icon)|_プロパティ_ > set|アイコンがその一部であるセットを表します。使用可能な値は次のとおりです。Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.2|
|[range](/javascript/api/excel/excel.range)|_プロパティ_ > columnHidden|現在の範囲のすべての列が非表示になっているかどうかを表します。|1.2|
|[range](/javascript/api/excel/excel.range)|_プロパティ_ > formulasR1C1|R1C1 スタイル表記の数式を表します。|1.2|
|[range](/javascript/api/excel/excel.range)|_プロパティ_ > hidden|現在の範囲のすべてのセルが非表示になっているかどうかを表します。読み取り専用です。|1.2|
|[range](/javascript/api/excel/excel.range)|_プロパティ_ > rowHidden|現在の範囲のすべての行が非表示になっているかどうかを表します。|1.2|
|[range](/javascript/api/excel/excel.range)|_リレーションシップ_ > sort|現在の範囲について、範囲の並べ替えを表します。読み取り専用。|1.2|
|[range](/javascript/api/excel/excel.range)|_メソッド_ > merge(across: bool)|範囲内のセルをワークシートの 1 つの領域に結合します。|1.2|
|[range](/javascript/api/excel/excel.range)|_メソッド_ > unmerge()|範囲内のセルを結合解除して別々のセルにします。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_プロパティ_ > columnWidth|範囲内のすべての列の幅を取得または設定します。列の幅が均一でない場合は、null が返されます。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_プロパティ_ > rowHeight|範囲内のすべての行の高さを取得または設定します。行の高さが均一でない場合は、null が返されます。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_リレーションシップ_ > protection|範囲に対する書式保護オブジェクトを返します。読み取り専用です。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_メソッド_ > autofitColumns()|現在の列のデータに基づいて、現在の範囲の列の幅を最適な幅に変更します。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_メソッド_ > autofitRows()|現在の行のデータに基づいて、現在の範囲の行の高さを最適な高さに変更します。|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_プロパティ_ > address|現在の範囲の表示されている行を表します。|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_メソッド_ > apply(fields: SortField, matchCase: bool, hasHeaders: bool, orientation: string, method: string)|並べ替え操作を実行します。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_プロパティ_ > ascending|昇順で並べ替えるかどうかを表します。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_プロパティ_ > color|並べ替えがフォントまたはセルの色で行われる場合に、条件の対象となる色を表します。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_プロパティ_ > dataOption|このフィールドのその他の並べ替えオプションを表します。使用可能な値は次のとおりです。Normal、TextAsNumber。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_プロパティ_ > key|条件の対象とする列 (または行。並べ替えの方向によって異なります) を表します。最初の列 (または行) からのオフセットとして表します。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_プロパティ_ > sortOn|この条件の並べ替えの種類を表します。使用可能な値は次のとおりです。Value、CellColor、FontColor、Icon。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_リレーションシップ_ > icon|並べ替えがセルのアイコンで行われる場合に、条件の対象となるアイコンを表します。|1.2|
|[table](/javascript/api/excel/excel.table)|_リレーションシップ_ > sort|テーブル内の並べ替えを表します。読み取り専用。|1.2|
|[table](/javascript/api/excel/excel.table)|_リレーションシップ_ > worksheet|現在のテーブルを含んでいるワークシート。読み取り専用です。|1.2|
|[table](/javascript/api/excel/excel.table)|_メソッド_ > clearFilters()|現在テーブルに適用されているすべてのフィルターをクリアします。|1.2|
|[table](/javascript/api/excel/excel.table)|_メソッド_ > convertToRange()|テーブルを通常の範囲のセルに変換します。すべてのデータが保持されます。|1.2|
|[table](/javascript/api/excel/excel.table)|_メソッド_ > reapplyFilters()|テーブルに現在設定されているすべてのフィルターを再適用します。|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_リレーションシップ_ > filter|列に適用されるフィルターを取得します。読み取り専用です。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_プロパティ_ > matchCase|大文字小文字の区別が、テーブルの最後の並べ替え操作に影響を与えたかどうかを表します。読み取り専用です。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_プロパティ_ > method|テーブルの並べ替えで最後に使用した中国語文字の順序付け方法を表します。読み取り専用です。使用可能な値は次のとおりです。PinYin、StrokeCount。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_リレーションシップ_ > fields|テーブルの最後の並べ替えに使用する現在の条件を表します。読み取り専用です。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_メソッド_ > apply(fields: SortField, matchCase: bool, method: string)|並べ替え操作を実行します。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_メソッド_ > clear()|テーブルに現在設定されている並べ替えをクリアします。これにより表の順序が変更されることはありませんが、ヘッダーのボタンの状態がクリアされます。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_メソッド_ > reapply()|テーブルに、現在の並べ替えパラメーターを再適用します。|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > functions|このブックを含む Excel アプリケーションのインスタンスを表します。読み取り専用。|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_リレーションシップ_ > protection|ワークシートのシート保護オブジェクトを返します。読み取り専用です。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_プロパティ_ > protected|ワークシートが保護されているかどうかを示します。読み取り専用。読み取り専用。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_リレーションシップ_ > options|シートの保護のオプション。読み取り専用。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_メソッド_ > protect(options: WorksheetProtectionOptions)|ワークシートを保護します。ワークシートがすでに保護されている場合は失敗します。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_メソッド_ > unprotect()|ワークシートの保護を解除します。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowAutoFilter|自動フィルター機能の使用を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowDeleteColumns|列の削除を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowDeleteRows|行の削除を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowFormatCells|セルの書式設定を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowFormatColumns|列の書式設定を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowFormatRows|行の書式設定を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowInsertColumns|列の挿入を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowInsertHyperlinks|ハイパーリンクの挿入を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowInsertRows|行の挿入を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowPivotTables|ピボットテーブル機能の使用を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowSort|並ベ替え機能の使用を可能にするワークシート保護オプションを表します。|1.2|

## <a name="excel-javascript-api-11"></a>Excel JavaScript API 1.1

Excel JavaScript API 1.1 は、API の最初のバージョンです。API について詳しくは、[Excel JavaScript API](/javascript/api/excel) リファレンスのトピックをご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
