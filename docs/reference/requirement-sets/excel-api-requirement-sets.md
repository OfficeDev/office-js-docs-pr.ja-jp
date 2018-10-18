# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインでは、マニフェストで指定されている要件セットを使用して、またはランタイム チェックを使用して、Office ホストがアドインを必要とする API をサポートするかどうかを決定します。 詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。

、および Office オンラインの複数のバージョンの Office 2016 を含め、Office や Windows、iPad の Office、Office for Mac の後で Excel のアドインを実行します。 Excel の要件の設定、要件セットごとに、およびビルドのバージョンをサポートしたり、それらのアプリケーションの Office ホスト アプリケーションを次の表に一覧します。

> [!NOTE]
>  **ベータ版** としてマークされているすべての API は、エンド ・ ユーザーの生産の準備が完了ではありません。 ここで使用できるように開発者は、テストおよび開発環境で試してみるのです。 運用とビジネスの重要なドキュメントに対して使用するのにはないです。
> 
>  **ベータ版**としてマークされている要件のセットを指定された (またはそれ以降) バージョンの Office ソフトウェアを使用し、CDN のベータ版のライブラリを使用して: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。  **ベータ版** としてマークされていないエントリは、一般的に使用できると、CDN の生産のライブラリを使用することができます: https://appsforoffice.microsoft.com/lib/1/hosted/office.js。

|  要件セット  |  Office for Windows\*  |  IPad の office 365  |  Office for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| ベータ版  |  [Excel の JavaScript API のオープンな仕様のページをご覧](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)ください。 |
| ExcelApi1.8  | バージョン 1808 (ビルド 10730.20102) | 2.17 以降 | 16.17 またはそれ以降 | 2018 年 9 月 | 近日公開 |
| ExcelApi1.7  | 1801 (ビルド 9001.2171) のバージョンまたはそれ以降   | 2.9 以降 | 16.9 またはそれ以降 | 2018 年 4 月 | 近日公開 |
| ExcelApi1.6  | バージョン 1704 (ビルド 8201.2001) 以降   | 2.2 以降 |15.36 以降| 2017 年 4 月 | 近日公開|
| ExcelApi1.5  | バージョン 1703 (ビルド 8067.2070) 以降   | 2.2 以降 |15.36 以降| 2017 年 3 月 | 近日公開|
| ExcelApi1.4  | バージョン 1701 (ビルド 7870.2024) 以降   | 2.2 以降 |15.36 以降| 2017 年 1 月 | 近日公開|
| ExcelApi1.3  | バージョン 1608 (ビルド 7369.2055) 以降 | 1.27 以降 |  15.27 以降| 2016 年 9 月 | バージョン 1608 (ビルド 7601.6800) 以降|
| ExcelApi1.2  | バージョン 1601 (ビルド 6741.2088) 以降 | 1.21 以降 | 15.22 以降| 2016 年 1 月 ||
| ExcelApi1.1  | バージョン 1509 (ビルド 4266.1001) 以降 | 1.19 以降 | 15.20 以降| 2016 年 1 月 ||

> [!NOTE]
> 注: MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。 のみ、このバージョンには、ExcelApi 1.1 要件のセットが含まれています。

バージョン、ビルド番号、および Office のオンライン サーバーの詳細についてを参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="whats-new-in-excel-javascript-api-18"></a>Excel JavaScript API 1.8 の新機能

Excel の JavaScript API の要件のセット 1.8 機能には、ピボット テーブル、データの入力規則、グラフ、グラフ、パフォーマンス オプション、およびブックの作成のイベントの Api が含まれます。

### <a name="pivottable"></a>ピボットテーブル

ピボット テーブル Api の Wave 2 では、ピボット テーブルの階層を設定するアドインを使用できます。 これで、データおよび集計方法を制御できます。 詳細については、新規のピボット テーブル機能は、 [ピボット テーブルの資料](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables) があります。

### <a name="data-validation"></a>データ確認

データの検証により、ユーザーの入力、ワークシートを制御します。 定義済みの応答を設定するセルを制限したり、不適切な入力のポップアップ警告を与えます。 詳細 [範囲のデータの入力規則を追加する](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation) 今日。

### <a name="charts"></a>グラフ

グラフ Api の別のラウンドでは、グラフの要素をさらに大きなプログラムに制御が表示されます。 凡例、軸、近似曲線、およびプロット エリアへのアクセス拡大があるようになりました。

### <a name="events"></a>予定

グラフの他の [イベント](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events) が追加されました。 グラフと対話するユーザーを追加で反応があります。  [イベントの表示/非表示](https://docs.microsoft.com/office/dev/add-ins/excel/performance#enable-and-disable-events) のブックを全体にわたって発生することもできます。


|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[アプリケーション](/javascript/api/excel/excel.application)|_メソッド_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|省略可能な base64 エンコードされた .xlsx ファイルを使用して、新しい非表示のブックを作成します。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_プロパティ_ > formula1|取得または設定の Formula1、つまり最小値または演算子の値が異なります。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_プロパティ_ > formula2|取得または Formula2、つまり最大値または演算子の値が異なりますを設定します。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_関係_ > 演算子|データの検証に使用する演算子です。|1.8|
|[グラフ](/javascript/api/excel/excel.chart)|_プロパティ_ > categoryLabelLevel|項目軸ラベルが供給される場所のレベルを参照する ChartCategoryLabelLevel 列挙の定数を設定または返します。 読み取り/書き込み。|1.8|
|[グラフ](/javascript/api/excel/excel.chart)|_プロパティ_ > plotVisibleOnly|可視セルのみの場合は true を指定します。 False の場合は、両方表示と非表示のセルがプロットされます。 読み取り/書き込み。|1.8|
|[グラフ](/javascript/api/excel/excel.chart)|_プロパティ_ > seriesNameLevel|シリーズ名が供給される場所のレベルを参照する ChartSeriesNameLevel 列挙の定数を設定または返します。 読み取り/書き込み。|1.8|
|[グラフ](/javascript/api/excel/excel.chart)|_プロパティ_ > showDataLabelsOverMaximum|値が数値軸の最大値より大きい場合にデータ ラベルを表示するかどうかを表します。|1.8|
|[グラフ](/javascript/api/excel/excel.chart)|_プロパティ_ > style|グラフのグラフのスタイルを設定または返します。 読み取り/書き込み。|1.8|
|[グラフ](/javascript/api/excel/excel.chart)|_関係_ > displayBlanksAs|グラフへの空白セルのプロット方法を設定します。値の取得および設定が可能です。XlDisplayBlanksAs クラスの定数を使用します。 読み取り/書き込み。|1.8|
|[グラフ](/javascript/api/excel/excel.chart)|_関係_ > plotArea|グラフの模様を表します。 読み取り専用です。|1.8|
|[グラフ](/javascript/api/excel/excel.chart)|_関係_ > plotBy|行または列は、グラフのデータ系列として使用する方法を設定または返します。 読み取り/書き込み。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_プロパティ_ > chartId|アクティブにするグラフの id を取得します。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_プロパティ_ > タイプ|イベントの種類を取得します。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_プロパティ_ > worksheetId|グラフがアクティブであるワークシートの id を取得します。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_プロパティ_> chartId|グラフをワークシートに追加の id を取得します。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_プロパティ_> worksheetId|グラフを追加するワークシートの id を取得します。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_関係_ > ソース|イベントのソースを取得します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > isBetweenCategories|数値軸の項目間の項目軸を交差するかどうかを表します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > マルチレベル|軸がマルチレベルかどうかを表します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > numberFormat|軸の目盛ラベルの表示形式のコードを表します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > オフセット|ラベルのレベル、および先頭レベルと軸線間の距離の間の距離を表します。 値は、0 ~ 1000 の整数にする必要があります。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > positionAt|他の軸と交差する、指定した軸の位置を表します。 このプロパティを設定するのには、SetPositionAt(double) メソッドを使用する必要があります。 読み取り専用です。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > textOrientation|軸の目盛ラベルのテキストの向きを表します。 値は整数である必要があります 90、またはテキストの垂直方向に 180 を-90 度からです。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_関係_ > 配置|指定した軸の目盛ラベルの配置を表します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_関係_ > 位置|指定した軸の位置を他の軸と交差する位置を表します。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_メソッド_ > [setPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|他の軸と交差する位置指定した軸の位置を設定します。|1.8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_関係_ > 塗りつぶし|グラフの塗りつぶしの書式を表します。 読み取り専用です。|1.8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_メソッド_ > [setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|A1 スタイルの表記を使用してグラフの軸ラベルの数式を表す文字列値です。|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_関係_ > 境界線|色、linestyle、および重量が含まれていますの枠線の書式設定を表します。 読み取り専用です。|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_関係_> 塗りつぶし|グラフの塗りつぶしの書式を表します。 読み取り専用です。|1.8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_メソッド_ > [clear()](/javascript/api/excel/excel.chartborder)|グラフ要素の線の書式をクリアします。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > 定型句|ブール値を表すしたデータが自動的にラベルを付ける場合は、新たなコンテキストに基づいて適切なテキストを生成します。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > 数式|A1 スタイルの表記を使用してグラフのデータ ラベルの数式を表す文字列値です。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > height|ポイント、グラフのデータ ラベルの高さを返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null です。 読み取り専用です。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > 左|グラフ エリアの左端にグラフのデータ ラベルの左端からの距離をポイント単位でを表します。 グラフのデータ ラベルが表示されない場合は null です。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > numberFormat|データ ラベルの表示形式コードを表す文字列値。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > text|グラフのデータ ラベルのテキストを表す文字列。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_> textOrientation|グラフのデータ ラベルのテキストの向きを表します。 値は整数である必要があります 90、またはテキストの垂直方向に 180 を-90 度からです。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > トップ|グラフ エリアの最上部にデータ ラベルをグラフの上端からの距離をポイント単位でを表します。 グラフのデータ ラベルが表示されない場合は null です。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > width|ポイント、グラフのデータ ラベルの幅を返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null です。 読み取り専用です。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_関係_ > の形式|グラフのデータ ラベルの書式を表します。 読み取り専用です。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_リレーションシップ_ > horizontalAlignment|グラフのデータ ラベルの水平方向の配置を表します。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_リレーションシップ_ > verticalAlignment|グラフのデータ ラベルの垂直方向の配置を表します。|1.8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_関係_> 境界線|色、linestyle、および重量が含まれていますの枠線の書式設定を表します。 読み取り専用です。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_プロパティ_> 定型句|データ ラベルが自動的にコンテキストに基づいて適切なテキストを生成するかどうかを表します。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_プロパティ_ > numberFormat|データ ラベルの表示形式コードを表します。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_プロパティ_> textOrientation|データ ラベルのテキストの向きを表します。 値 90 度、-90 度からか、0 ~ 180 テキストの垂直方向の整数である必要があります。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_リレーションシップ_ > horizontalAlignment|グラフのデータ ラベルの水平方向の配置を表します。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_リレーションシップ_ > verticalAlignment|グラフのデータ ラベルの垂直方向の配置を表します。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_プロパティ_> chartId|非アクティブ化するグラフの id を取得します。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_プロパティ_> worksheetId|グラフが非アクティブ化するワークシートの id を取得します。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_プロパティ_> chartId|ワークシートから削除するグラフの id を取得します。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_プロパティ_> worksheetId|グラフを削除するワークシートの id を取得します。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_関係_> ソース|イベントのソースを取得します。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_ > height|グラフの凡例の凡例の高さを表します。 読み取り専用です。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_ > index|グラフの凡例の凡例のインデックスを表します。 読み取り専用です。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_> 左|グラフの凡例の左側を表します。 読み取り専用です。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_> トップ|グラフの凡例の一番上を表します。 読み取り専用です。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_ > width|グラフの凡例の凡例の幅を表します。 読み取り専用です。|1.8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_関係_> 境界線|色、linestyle、および重量が含まれていますの枠線の書式設定を表します。 読み取り専用です。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > height|返すプロパティの高さの値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > insideHeight|返すプロパティの insideHeight 値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > insideLeft|InsideLeft 返すプロパティ値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > insideTop|InsideTop 返すプロパティ値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > insideWidth|PlotArea insideWidth 値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_> 左|返すプロパティの左側の値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_> トップ|返すプロパティの一番上の値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_プロパティ_ > width|返すプロパティの幅の値を表します。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_関係_> の形式|グラフを返すプロパティの書式を表します。 読み取り専用です。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_関係_> 位置|返すプロパティの位置を表します。|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_関係_> 境界線|グラフを返すプロパティの境界線の属性を表します。 読み取り専用です。|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_関係_> 塗りつぶし|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。値の取得のみ可能です。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > 爆発|返すまたは、円グラフまたはドーナツ グラフのスライスの切り出し値を設定します。 切り出し表示を行わず、扇形の中心が円の中心と一致する場合、このプロパティは 0 を返します。 読み取り/書き込み。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > firstSliceAngle|時計回りに (垂直) の最初の円グラフまたはドーナツ グラフのスライスの角度を設定または返します。 このプロパティの対象は、円グラフ、3-D 円グラフ、およびドーナツ グラフだけです。 値の範囲は 0 ～ 360 です。 読み取り/書き込み|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > invertIfNegative|True の場合、負の数であるデータ要素を反転します。値の取得および設定が可能です。バリアント型 ( Variant ) の値を使用します。 読み取り/書き込み。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > 重複します。|バーと列を配置する方法を指定します。 -100 から 100 の間で値を指定できます。 このプロパティの対象は、2-D 横棒グラフと 2-D 縦棒グラフだけです。 読み取り/書き込み。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > secondPlotSize|主な円グラフのサイズに対する割合で、補助円グラフ付き円グラフまたは円グラフでは、いずれかのセカンダリのセクションのサイズを設定または返します。 5 から 200 までの値をすることができます。 読み取り/書き込み。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > varyByCategories|True の 場合、Microsoft Word の各データ マーカーにそれぞれ異なる色またはパターンを割り当てます。 データ系列が 1 つしか含まれないグラフを対象とします。 読み取り/書き込み。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_関係_ > axisGroup|指定した系列のグループを取得または設定します。値の取得および設定が可能です。 読み取り/書き込み|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_関係_ > データ ラベル|系列内のすべてのデータ ラベルのコレクションを表します。 読み取り専用です。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_関係_ > splitType|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフを 2 つの部分に分割する方法を設定します。値の取得および設定が可能です。XlChartSplitType クラスの定数を使用します。 読み取り/書き込み。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > backwardPeriod|近似曲線を後方に延長された期間の数を表します。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > forwardPeriod|近似曲線を前方に拡張する期間の数を表します。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > showEquation|True を設定すると、近似曲線の方程式をグラフに表示されます。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > showRSquared|True を設定すると、グラフの近似曲線の R 平方が表示されます。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_関係_ > ラベル|グラフの近似曲線のラベルを表します。 読み取り専用です。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_> 定型句|ブール値を表す場合、近似曲線ラベルは、コンテキストに基づいて適切なテキストを自動的に生成されます。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_> 数式|A1 スタイルの表記を使用してグラフの近似曲線ラベルの数式を表す文字列値。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > height|ポイント、グラフの近似曲線ラベルの高さを返します。 読み取り専用です。 グラフの近似曲線のラベルが表示されない場合は null です。 読み取り専用です。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_> 左|グラフ エリアの左端にグラフの近似曲線ラベルの左端からの距離をポイント単位でを表します。 グラフの近似曲線のラベルが表示されない場合は null です。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > numberFormat|近似曲線ラベルの書式コードを表す文字列値。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > text|グラフの近似曲線ラベルのテキストを表す文字列。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_> textOrientation|グラフの近似曲線ラベルのテキストの向きを表します。 値は整数である必要があります 90、またはテキストの垂直方向に 180 を-90 度からです。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_> トップ|グラフ エリアの最上部にグラフの近似曲線ラベルの上端からの距離をポイント単位でを表します。 グラフの近似曲線のラベルが表示されない場合は null です。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_プロパティ_ > width|ポイント、グラフの近似曲線ラベルの幅を返します。 読み取り専用です。 グラフの近似曲線のラベルが表示されない場合は null です。 読み取り専用です。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_関係_> の形式|グラフの近似曲線ラベルの書式を表します。 読み取り専用です。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_リレーションシップ_ > horizontalAlignment|グラフの近似曲線ラベルの水平方向の配置を表します。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_リレーションシップ_ > verticalAlignment|グラフの近似曲線ラベルの垂直方向の配置を表します。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_関係_> 境界線|色、linestyle、および重量が含まれていますの枠線の書式設定を表します。 読み取り専用です。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_関係_> 塗りつぶし|現在のグラフの近似曲線ラベルの塗りつぶしの書式を表します。 読み取り専用です。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_リレーションシップ_ > font|グラフのデータ ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。値の取得のみ可能です。 読み取り専用です。|1.8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_プロパティ_ > fakeFileId|クライアント側では、TableSelectionChangedEvent の worksheetId などの追加データを送信します。|1.8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_プロパティ_ > fileBase64|クライアント側では、TableSelectionChangedEvent の worksheetId などの追加データを送信します。|1.8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_関係_ > ファイアウォール|クライアント側では、TableSelectionChangedEvent の worksheetId などの追加データを送信します。|1.8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_プロパティ_> 数式| カスタムのデータ入力規則の数式です。 これは、重複を防止するか、セル範囲の合計を制限するなどの特殊な入力規則を作成します。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_プロパティ_ > id|DataPivotHierarchy の id です。 読み取り専用です。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_プロパティ_ > name|DataPivotHierarchy の名前です。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_プロパティ_ > numberFormat|DataPivotHierarchy の表示形式です。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_プロパティ_ > 位置|DataPivotHierarchy の位置です。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_関係_ > フィールド|DataPivotHierarchy に関連付けられているピボット フィールドを返します。 読み取り専用です。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_関係_ > showAs|特定の計算結果としてデータを表示するかどうかどうかを決定します。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_関係_ > summarizeBy|DataPivotHierarchy のすべての項目を表示するかどうかを決定します。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_メソッド_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault)|DataPivotHierarchy をその既定値にリセットします。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_プロパティ_ > items|DataPivotHierarchy オブジェクトのコレクションです。 読み取り専用です。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_メソッド_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|現在の軸に、PivotHierarchy を追加します。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|コレクション内には、ピボットの階層の数を取得します。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|DataPivotHierarchy の名前または id を取得します。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|名前で、DataPivotHierarchy を取得します。 DataPivotHierarchy が存在しない場合は、null オブジェクトを返します。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_メソッド_ > [remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|現在の軸から、PivotHierarchy を削除します。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_プロパティ_ > ignoreBlanks|空白を無視する: 空白のセルにデータ検証は実行されません、既定で true に設定します。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_プロパティ_ > 有効な|すべてのセル値は、データの入力規則に従った有効な場合を表します。 読み取り専用です。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_関係_ > errorAlert|無効なデータが入力されたときに警告をエラーします。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_関係_ > プロンプト|ユーザーがセルを選択したときにプロンプトします。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_関係_ > ルール|さまざまな種類データ入力規則の条件にはが含まれているデータの入力規則です。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_リレーションシップ_ > type|データの入力規則の種類、詳細については [Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype) を参照してください。 読み取り専用です。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_メソッド_ > [clear()](/javascript/api/excel/excel.datavalidation)|現在の範囲からのデータの入力規則を削除します。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_プロパティ_ > メッセージ|警告のエラー メッセージを表します。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_プロパティ_ > showAlert|またはないユーザーが無効なデータを入力とエラー警告のダイアログ ボックスを表示するかどうかを決定します。 既定値は True です。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_プロパティ_ > title|エラー通知ダイアログのタイトルを表します。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_関係_ > スタイル|データの入力規則を表す警告の種類、詳細については [Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle) を参照してください。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_プロパティ_> メッセージ|確認のメッセージを表します。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_プロパティ_ > showPrompt|ユーザーがデータの入力規則のセルを選択したときに、プロンプトを表示するかどうかを決定します。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_プロパティ_ > title|プロンプトのタイトルを表します。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_関係_ > カスタム|カスタムのデータ入力規則の条件です。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_関係_ > 日付|日付データの入力規則の条件です。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_関係_ > 10 進数|10 進数のデータの入力規則の条件です。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_リレーションシップ_ > list|データ入力規則の条件を一覧表示します。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_関係_ > textLength|TextLength データの入力規則の条件です。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_関係_ > 時間|時間データの入力規則の条件です。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_関係_ > wholeNumber|データ入力規則の条件を WholeNumber。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_プロパティ_> formula1|Formula1、つまり最小値または演算子は、値を設定を取得または取得します。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_プロパティ_> formula2|取得または Formula2、つまり最大値または演算子は、値を設定します。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_関係_> 演算子|データの検証に使用する演算子です。|1.8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_プロパティ_ > isEnableEvents {|クライアント側では、TableSelectionChangedEvent の worksheetId などの追加データを送信します。|1.8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_関係_> ファイアウォール|クライアント側では、TableSelectionChangedEvent の worksheetId などの追加データを送信します。|1.8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_関係_ > controlId|クライアント側では、TableSelectionChangedEvent の worksheetId などの追加データを送信します。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_プロパティ_ > enableMultipleFilterItems|複数のアイテムのフィルターを許可するかどうかを決定します。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_プロパティ_ > id|FilterPivotHierarchy の id です。 読み取り専用です。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_プロパティ_ > name|FilterPivotHierarchy の名前です。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_プロパティ_> 位置|FilterPivotHierarchy の位置です。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_リレーションシップ_ > fields|FilterPivotHierarchy に関連付けられているピボット フィールドを返します。 読み取り専用です。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_メソッド_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|FilterPivotHierarchy をその既定値にリセットします。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_プロパティ_ > items|FilterPivotHierarchy オブジェクトのコレクションです。 読み取り専用です。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_メソッド_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|現在の軸に、PivotHierarchy を追加します。 階層が行、列、またはフィルター軸の別の場所である場合は、その場所から削除されます。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|コレクション内には、ピボットの階層の数を取得します。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|FilterPivotHierarchy の名前または id を取得します。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|名前で、FilterPivotHierarchy を取得します。 FilterPivotHierarchy が存在しない場合は、null オブジェクトを返します。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_メソッド_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|現在の軸から、PivotHierarchy を削除します。|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_プロパティ_ > inCellDropDown|か、セルの一覧がドロップダウンを表示するには、デフォルトは true です。|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_プロパティ_ > ソース|データの入力規則のリストのソース|1.8|
|[openWorkbookPostProcessAction](/javascript/api/excel/excel.openworkbookpostprocessaction)|_プロパティ_> fakeFileId|クライアント側では、TableSelectionChangedEvent の worksheetId などの追加データを送信します。|1.8|
|[openWorkbookPostProcessAction](/javascript/api/excel/excel.openworkbookpostprocessaction)|_関係_> ファイアウォール|クライアント側では、TableSelectionChangedEvent の worksheetId などの追加データを送信します。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_プロパティ_ > id|ピボット フィールドの id です。 読み取り専用です。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_プロパティ_ > name|ピボット フィールドの名前です。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_プロパティ_ > showAllItems|ピボット フィールドのすべてのアイテムを表示するかどうかを決定します。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_関係_ > アイテム|ピボット フィールドに関連付けられているピボット フィールドを返します。 読み取り専用です。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_関係_ > 小計|ピボット フィールドの小計です。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_メソッド_ > [sortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|ピボット フィールドを並べ替えます。 DataPivotHierarchy を指定すると、し、並べ替えをに基づいて、適用するされていない場合の並べ替えは、ピボット フィールド自体に基づいて行われます。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_プロパティ_ > items|PivotField オブジェクトのコレクションです。 読み取り専用です。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|コレクション内には、ピボットの階層の数を取得します。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|PivotHierarchy の名前または id を取得します。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|名前で、PivotHierarchy を取得します。 PivotHierarchy が存在しない場合は、null オブジェクトを返します。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_プロパティ_ > id|PivotHierarchy の id です。 読み取り専用です。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_プロパティ_ > name|PivotHierarchy の名前です。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_リレーションシップ_ > fields|PivotHierarchy に関連付けられているピボット フィールドを返します。 読み取り専用です。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_プロパティ_ > items|PivotHierarchy オブジェクトのコレクションです。 読み取り専用です。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|コレクション内には、ピボットの階層の数を取得します。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|PivotHierarchy の名前または id を取得します。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|名前で、PivotHierarchy を取得します。 PivotHierarchy が存在しない場合は、null オブジェクトを返します。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_プロパティ_ > id|PivotItem の id です。 読み取り専用です。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_プロパティ_ > isExpanded|子アイテムを表示する項目が展開されているかどうか、または折りたたまれているかどうかと、子アイテムが非表示を決定します。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_プロパティ_ > name|またはの名前です。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_プロパティ_ > 表示|か、またはが表示されているかどうかを決定します。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_プロパティ_ > items|PivotItem オブジェクトのコレクションです。 読み取り専用です。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.pivotitemcollection)|コレクション内には、ピボットの階層の数を取得します。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection)|PivotHierarchy の名前または id を取得します。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection)|名前で、PivotHierarchy を取得します。 PivotHierarchy が存在しない場合は、null オブジェクトを返します。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_プロパティ_ > showColumnGrandTotals|True を設定すると列の合計をピボット テーブル レポートに総計が表示されます。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_プロパティ_ > showRowGrandTotals|True を設定すると行の合計をピボット テーブル レポートに総計が表示されます。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_プロパティ_ > subtotalLocation|このプロパティは、ピボット テーブルのすべてのフィールドの SubtotalLocationType を示します。 フィールドには、さまざまな状態がある、これは null になります。 使用可能な値: AtTop、AtBottom。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_関係_ > layoutType|このプロパティは、ピボット テーブルのすべてのフィールドの PivotLayoutType を示します。 フィールドには、さまざまな状態がある、これは null になります。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_メソッド_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|ピボット テーブルの列のラベルが存在する範囲を返します。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_メソッド_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|ピボット テーブルのデータ値が存在する範囲を返します。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout.md)|_メソッド_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|ピボット テーブルのフィルター領域の範囲を返します。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_メソッド_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|除外フィルター エリアに、ピボット テーブルが存在する範囲を返します。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_メソッド_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|ピボット テーブルの行のラベルが存在する範囲を返します。|1.8|
|[ピボットテーブル](/javascript/api/excel/excel.pivottable)|_関係_ > columnHierarchies|ピボット テーブルの列のピボットの階層です。 読み取り専用です。|1.8|
|[ピボットテーブル](/javascript/api/excel/excel.pivottable)|_関係_ > dataHierarchies|ピボット テーブルのデータのピボットの階層です。 読み取り専用です。|1.8|
|[ピボットテーブル](/javascript/api/excel/excel.pivottable)|_関係_ > filterHierarchies|ピボット テーブルのフィルターのピボットの階層です。 読み取り専用です。|1.8|
|[ピボットテーブル](/javascript/api/excel/excel.pivottable)|_関係_ > 階層|ピボット テーブルのピボットの階層です。 読み取り専用です。|1.8|
|[ピボットテーブル](/javascript/api/excel/excel.pivottable)|_関係_ > レイアウト|レイアウトと、ピボット テーブルの視覚的な構造を表す PivotLayout です。 読み取り専用です。|1.8|
|[ピボットテーブル](/javascript/api/excel/excel.pivottable)|_関係_ > rowHierarchies|ピボット テーブルの行のピボットの階層です。 読み取り専用です。|1.8|
|[ピボットテーブル](/javascript/api/excel/excel.pivottable)|_メソッド_ > [delete()](/javascript/api/excel/excel.pivottable)|ピボット テーブルを削除します。|1.8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_メソッド_ > [add(name: string, source: object, destination: object)](/javascript/api/excel/excel.pivottablecollection)|指定したソース データに基づくピボット テーブルを追加し、配置先範囲の左上のセルに挿入します。|1.8|
|[範囲](/javascript/api/excel/excel.range)|_関係_ > データの入力規則|データ検証オブジェクトを返します。 読み取り専用です。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_プロパティ_ > id|RowColumnPivotHierarchy の id です。 読み取り専用です。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_プロパティ_ > name|RowColumnPivotHierarchy の名前です。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_プロパティ_> 位置|RowColumnPivotHierarchy の位置です。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_リレーションシップ_ > fields|RowColumnPivotHierarchy に関連付けられているピボット フィールドを返します。 読み取り専用です。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_メソッド_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|RowColumnPivotHierarchy をその既定値にリセットします。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_プロパティ_ > items|RowColumnPivotHierarchy オブジェクトのコレクションです。 読み取り専用です。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_メソッド_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|現在の軸に、PivotHierarchy を追加します。 階層が行の他の場所に存在する場合] 列で、|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_メソッド_ > [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|コレクション内には、ピボットの階層の数を取得します。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_メソッド_ > [getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|RowColumnPivotHierarchy の名前または id を取得します。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_メソッド_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|名前で、RowColumnPivotHierarchy を取得します。 RowColumnPivotHierarchy が存在しない場合は、null オブジェクトを返します。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_メソッド_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|現在の軸から、PivotHierarchy を削除します。|1.8|
|[ランタイム](/javascript/api/excel/excel.runtime)|_プロパティ_ > enableEvents|現在の作業ウィンドウまたはコンテンツの追加の JavaScript イベントを切り替えます。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_関係_ > baseField|該当する場合、ShowAs 計算の基に、基本のピボット フィールドは、ShowAsCalculation 型、それ以外の場合に null に基づいています。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_関係_ > baseItem|ShowAs の計算の基に該当する場合に基本のアイテムは、ShowAsCalculation 型、それ以外の場合に null に基づいています。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_関係_ > 計算|データ ピボット フィールドに使用する ShowAs 計算します。|1.8|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > 自動インデント|セル内のテキストの配置が均等に設定されている場合かどうかはテキストを自動的にインデントを示します。|1.8|
|[style](/javascript/api/excel/excel.style)|_プロパティ_> textOrientation|スタイルのテキストの向きです。|1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > 自動|自動が true に設定し、他のすべての値に設定されている場合は、小計を設定するときに無視されます。|1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > 平均| |1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > カウント| |1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > countNumbers| |1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > 最大| |1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > 分| |1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > 製品| |1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > standardDeviation| |1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > standardDeviationP| |1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > 合計| |1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > 差異| |1.8|
|[小計](/javascript/api/excel/excel.subtotals)|_プロパティ_ > varianceP| |1.8|
|[テーブル](/javascript/api/excel/excel.table)|_プロパティ_ > legacyId|数値の id を返します。読み取り専用です。|1.8|
|[ブック](/javascript/api/excel/excel.workbook)|_プロパティ_ > 読み取り専用|ブックが読み取り専用モードで開かれている場合は true。 読み取り専用です。|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_プロパティ_ > id|WorkbookCreated オブジェクトを一意に識別する値を返します。 読み取り専用です。|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_メソッド_ > [open()](/javascript/api/excel/excel.workbookcreated)|ブックを開きます。|1.8|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_プロパティ_ > showGridlines|取得または、ワークシートの枠線] のフラグを設定します。|1.8|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_プロパティ_ > showHeadings|取得または、ワークシートの見出しのフラグを設定します。|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_プロパティ_> worksheetId|計算されるワークシートの id を取得します。|1.8|

## <a name="whats-new-in-excel-javascript-api-17"></a>Excel JavaScript API 1.7 の新機能

Excel の JavaScript API の要件のセット 1.7 機能には、グラフ、イベント、ワークシート、範囲、ドキュメントのプロパティ、項目、保護オプション、およびスタイルをという名前の Api が含まれます。

### <a name="customize-charts"></a>グラフをカスタマイズします。

グラフでは、新しい Api、その他のグラフの種類を作成、データ系列をグラフに追加する、グラフ タイトルの設定、軸のタイトルを追加、表示単位を追加、移動平均近似曲線を追加、できます、線形近似曲線を変更します。 以下はその例です。

* グラフの軸 - を取得、設定、書式設定およびグラフの軸単位、ラベル、およびタイトルを削除します。
* グラフの系列に追加するには、設定、およびグラフのデータ系列を削除します。  系列マーカー、プロットのオーダーとサイズを変更します。
* グラフの近似曲線を追加するには、取得、およびグラフの近似曲線の書式を設定します。
* グラフの凡例: グラフの凡例のフォントの書式を設定します。
* グラフのポイントのグラフのポイントの色を設定します。
* グラフのタイトルの部分文字列を取得し、グラフのタイトルの部分文字列を設定します。
* グラフの種類の複数の種類のグラフを作成するオプションです。

### <a name="events"></a>予定

Api が自動的に特定のイベントと、指定された関数を実行するように追加できるようにするイベント ハンドラーのさまざまなを提供する Excel のイベントが発生します。 その関数は、目的のシナリオに必要なアクションを実行するように設計できます。 現在利用可能なイベントのリストは、 [Excel の JavaScript API を使用してイベントでの作業](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events)を参照してください。

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>範囲とワークシートの外観をカスタマイズします。

新しい Api を使用すると、複数の方法でワークシートの外観をカスタマイズできます。

* 特定の行や列を表示したままワークシートをスクロールするときにウィンドウを固定します。 たとえば、ワークシートの最初の行にヘッダーが含まれている場合可能性があります固定行ようにすること列ヘッダーが表示されたままワークシートをスクロールするとします。
* ワークシートのタブの色を変更します。
* ワークシートの見出しを追加します。


複数の方法で範囲の外観をカスタマイズできます。

* 確認するセル範囲のセル スタイルを設定する一貫性のある書式を範囲内のすべてのセルがあることを確認します。 セルのスタイルは、フォントとフォント サイズ、表示形式、セルの罫線、セルの網かけなどの書式設定の定義済みセットです。 任意の Excel の組み込みのセル スタイルを使用して、または独自のカスタム セル スタイルを作成します。
* 範囲のテキストの向きを設定します。
* 追加またはブック内の別の場所にまたは外部の場所にリンクする範囲のハイパーリンクを変更します。

### <a name="manage-document-properties"></a>ドキュメントのプロパティを管理します。

Api のドキュメントのプロパティを使用すると、組み込みのドキュメント プロパティにアクセスしも作成して、ブック、およびドライブのワークフローとビジネス ロジックの状態を格納するカスタム ドキュメント プロパティを管理します。

### <a name="copy-worksheets"></a>ワークシートをコピーします。

ワークシートのコピー Api を使用して、同じブック内の新しいワークシートに 1 つのワークシートからのデータと書式をコピーし、必要なデータ転送量を削減できます。

### <a name="handle-ranges-with-ease"></a>範囲を容易に処理します。

さまざまな範囲の Api を使用すると、周囲の領域の取得などを行う、サイズが変更された範囲を取得します。 これらの Api 操作の範囲のアドレス指定などの作業は効率が向上する必要があります。

さらに、

* ブックとワークシートの保護オプションでは、ワークシートとブックの構造内のデータを保護するためにこれらの Api を使用します。
* 名前付きアイテムを更新する-名前付きの項目を更新するのにはこの API を使用します。
* アクティブなセルを取得する - この API を使用して、ブックのアクティブなセルを取得します。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[グラフ](/javascript/api/excel/excel.chart)|_プロパティ_ > グラフの種類|グラフの種類を表します。 使用可能な値: ColumnClustered、ColumnStacked、ColumnStacked100、BarClustered、BarStacked、BarStacked100、LineStacked、LineStacked100、LineMarkers、LineMarkersStacked、LineMarkersStacked100、補助縦棒グラフ、等..|1.7|
|[グラフ](/javascript/api/excel/excel.chart)|_プロパティ_ > id|グラフの一意の id です。 読み取り専用です。|1.7|
|[グラフ](/javascript/api/excel/excel.chart)|_プロパティ_ > showAllFieldButtons|ピボット グラフですべてのフィールド ボタンを表示するかどうかを表します。|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_関係_> 境界線|色、linestyle、および重量が含まれているグラフ エリアの枠線の書式を表します。 読み取り専用です。|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_メソッド_ > getItem (型: 文字列、グループ化: 文字列)|タイプとグループによって識別される特定の軸を返します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > axisBetweenCategories|数値軸の項目間の項目軸を交差するかどうかを表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > axisGroup|指定された軸のグループを表します。 読み取り専用です。 使用可能な値: プライマリ、セカンダリです。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > [社員コード]|項目軸の種類を設定または返します。 使用可能な値: 自動、TextAxis、DateAxis。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > との交点|他の軸と交差する位置の指定した軸を表します。 使用可能な値: 自動、最大値、最小値、ユーザー設定します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > crossesAt|指定された軸には、他の軸と交差する位置を表します。 読み取り専用です。 このプロパティに設定するには、SetCrossesAt(double) メソッドを使用する必要があります。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > customDisplayUnit|カスタムの軸の表示単位の値を表します。 読み取り専用です。 このプロパティを設定するには、SetCustomDisplayUnit(double) メソッドを使用してください。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > displayUnit|軸の表示単位を表します。 使用可能な値: なし、数百、数千、TenThousands、HundredThousands、数百万、TenMillions、HundredMillions、数十億、「数兆、カスタムです。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > height|グラフ オブジェクトの高さをポイント単位で表します。 軸の表示されていない場合は null です。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_> 左|グラフ エリアの左側に軸の左端からの距離をポイント単位でを表します。 軸の表示されていない場合は null です。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > logBase|対数目盛を使用する場合は、対数の底を表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > reversePlotOrder|Microsoft Excel に最初から最後のデータ ポイントをプロットするかどうかを表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > scaleType|数値軸の目盛の種類を表します。 値を指定できます。 線形近似、対数。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > showDisplayUnitLabel|軸の表示単位ラベルが表示されるかどうかを表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > tickLabelSpacing|項目または目盛ラベル間でデータ系列の数を表します。 1 から 31999 または自動設定に空の文字列から値を指定できます。 返された値は、常に番号です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > tickMarkSpacing|項目または系列の目盛りの間の数を表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_> トップ|グラフ エリアの最上部に軸の上端からの距離をポイント単位でを表します。 軸の表示されていない場合は null です。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_> タイプ|軸の種類を表します。 読み取り専用です。 使用可能な値: 無効なカテゴリ、値、系列です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_> 表示|ブール値は、軸の表示/非表示を表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_プロパティ_ > width|グラフ オブジェクトの幅をポイント単位で表します。 軸の表示されていない場合は null です。 読み取り専用です。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_関係_ > baseTimeUnit|指定された項目軸の基本単位を設定または返します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_関係_ > majorTickMark|指定した軸の目盛の種類を表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_関係_ > majorTimeUnitScale|[社員コード] プロパティがxlTimeScale に設定すると、項目軸の目盛間隔のスケールの値を設定を取得または取得します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_関係_ > minorTickMark|指定された軸の補助目盛の種類を表します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_関係_ > minorTimeUnitScale|[社員コード] プロパティが xlTimeScale に設定すると、項目軸の補助目盛間隔のスケールの値を設定を取得または取得します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_関係_ > tickLabelPosition|指定された軸の目盛ラベルの位置を指定します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_メソッド_ > setCategoryNames(sourceData: Range)|指定した軸のすべてのカテゴリ名を設定します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_メソッド_ > setCrossesAt(value: double)|指定された軸には、他の軸と交差する位置を設定します。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_メソッド_ > setCustomDisplayUnit(value: double)|カスタムの値軸表示単位を設定します。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_プロパティ_ > color|グラフの線の色を表す HTML カラー コード。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_プロパティ_ > 重量|ポイント単位で、罫線の太さを表します。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_関係_ > 縦|罫線の線のスタイルを表します。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_> 位置|データ ラベルの位置を表すDataLabelPosition 値。使用可能な値は次のとおりです。None、Center、InsideEnd、InsideBase、OutsideEnd、Left、Right、Top、Bottom、BestFit、Callout。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > 区切り記号|グラフのデータ ラベルに使用される区切り文字を表す文字列を設定します。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showBubbleSize|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール型の値。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showCategoryName|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール型の値。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showLegendKey|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール型の値。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showPercentage|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール型の値。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showSeriesName|データ ラベルの系列名を表示するか非表示にするかを表すブール型の値。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_プロパティ_ > showValue|データ ラベルの値を表示するか非表示にするかを表すブール型の値。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_プロパティ_ > height|グラフの凡例の高さを表します。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_プロパティ_> 左|グラフの凡例の左側を表します。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_プロパティ_ > showShadow|凡例がグラフに影を持つ場合を表します。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_プロパティ_> トップ|グラフの凡例の一番上を表します。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_プロパティ_ > width|グラフの凡例の幅を表します。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_関係_ > 返すメソッド|凡例に返すメソッドのコレクションを表します。 読み取り専用です。|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_プロパティ_> 表示|グラフの凡例の表示を表します。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_プロパティ_ > items|ChartLegendEntry オブジェクトのコレクションです。 読み取り専用です。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_メソッド_ > getCount()|コレクション内の凡例の数を返します。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_メソッド_ > getItemAt(index: number)|凡例を指定したインデックス位置を返します。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_プロパティ_ > hasDataLabel|データ要素にデータ ラベルがあるかどうかを表します。 等高線グラフには適用されません。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_プロパティ_ > markerBackgroundColor|HTML のデータ マーカーの背景色の色コードの表現は、次のポイントです。 例: #Ff0000 は、赤を表します。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_プロパティ_ > markerForegroundColor|HTML データのマーカーの前景色の色コードの表現は、次のポイントです。 例: #Ff0000 は、赤を表します。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_プロパティ_ > markerSize|データ ポイントのマーカーのサイズを表します。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_プロパティ_ > markerStyle|グラフのデータ ポイントのマーカーのスタイルを表します。 使用可能な値: 無効、自動、なし、正方形、ひし形、三角形、X、スター型、ドット (.)、ダッシュ、円、および、画像です。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_関係_ > dataLabel|グラフの要素のデータ ラベルを返します。 読み取り専用です。|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_関係_> 境界線|色、スタイル、重量の情報が含まれているグラフのデータ要素の枠線の書式を表します。 読み取り専用です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_> グラフの種類|系列のグラフの種類を表します。 使用可能な値: ColumnClustered、ColumnStacked、ColumnStacked100、BarClustered、BarStacked、BarStacked100、LineStacked、LineStacked100、LineMarkers、LineMarkersStacked、LineMarkersStacked100、補助縦棒グラフ、等..|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > doughnutHoleSize|グラフの系列のドーナツの穴の大きさを表します。  ドーナツ グラフおよび doughnutExploded のグラフでのみ有効です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > フィルター|ブール値を表すか、系列をフィルターの場合です。 等高線グラフには適用されません。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > gapWidth|グラフの系列の棒の間隔を表します。  だけが有効なのバーと縦棒グラフでは、だけでなく|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > hasDataLabels|系列ラベルがあるない場合のデータかを表すブール値です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_> markerBackgroundColor|グラフの系列のマーカーの背景色を表します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_> markerForegroundColor|グラフの系列のマーカーの前景色を表します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_> markerSize|グラフの系列のマーカーのサイズを表します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_> markerStyle|グラフの系列のデータ マーカーのスタイルを表します。 使用可能な値: 無効、自動、なし、正方形、ひし形、三角形、X、スター型、ドット (.)、ダッシュ、円、および、画像です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > plotOrder|グラフ種類グループのグラフの系列をプロットする順序を表します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_> showShadow|場合か、シャドウ データ系列を表すブール値です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_プロパティ_ > 滑らかな|系列が滑らかでない、または場合はブール値を表します。 折れ線グラフおよび散布図のグラフにのみ|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_関係_> データ ラベル|系列内のすべてのデータ ラベルのコレクションを表します。 読み取り専用です。|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_関係_ > 近似曲線|系列の近似曲線のコレクションを表します。 読み取り専用です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_メソッド_ > delete()|グラフの系列を削除します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_メソッド_ > setBubbleSizes(sourceData: Range)|グラフの系列のバブル サイズを設定します。 バブル チャートでのみ有効です。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_メソッド_ > setValues(sourceData: Range)|グラフの系列の値を設定します。 散布図のグラフの Y 軸の値を意味します。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_メソッド_ > setXAxisValues(sourceData: Range)|軸をグラフの系列の X の値を設定します。 散布図でのみ有効です。|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_メソッド_ > を追加 (名前: 文字列、インデックス: 数)|コレクションに新しいデータ系列を追加します。|1.7|
|[返すプロパティ](/javascript/api/excel/excel.charttitle)|_プロパティ_ > height|ポイント単位で、グラフ タイトルの高さを返します。 読み取り専用です。 グラフ タイトルの表示されていない場合は null です。 読み取り専用です。|1.7|
|[返すプロパティ](/javascript/api/excel/excel.charttitle)|_プロパティ_ > horizontalAlignment|グラフ タイトルの水平方向の配置を表します。 使用可能な値: 中央、左、両端揃え、均等割り付け、右です。|1.7|
|[返すプロパティ](/javascript/api/excel/excel.charttitle)|_プロパティ_> 左|グラフ エリアの左端をグラフ タイトルの左端からの距離をポイント単位でを表します。 グラフ タイトルの表示されていない場合は null です。|1.7|
|[返すプロパティ](/javascript/api/excel/excel.charttitle)|_プロパティ_> 位置|グラフ タイトルの位置を表します。 使用可能な値: 上、自動、下、右、左です。|1.7|
|[返すプロパティ](/javascript/api/excel/excel.charttitle)|_プロパティ_> showShadow|グラフのタイトルに影があるかどうかを決定するブール値を表します。|1.7|
|[返すプロパティ](/javascript/api/excel/excel.charttitle)|_プロパティ_> textOrientation|グラフ タイトルのテキストの向きを表します。 値は整数である必要があります 90、またはテキストの垂直方向に 180 を-90 度からです。|1.7|
|[返すプロパティ](/javascript/api/excel/excel.charttitle)|_プロパティ_> トップ|グラフ エリアの最上部にグラフ タイトルの上端からの距離をポイント単位でを表します。 グラフ タイトルの表示されていない場合は null です。|1.7|
|[返すプロパティ](/javascript/api/excel/excel.charttitle)|_プロパティ_ > verticalAlignment|グラフ タイトルの縦位置を表します。 使用可能な値: センター、下、上、両端揃え、均等割り付けします。|1.7|
|[返すプロパティ](/javascript/api/excel/excel.charttitle)|_プロパティ_ > width|グラフのタイトルをポイント単位での幅を返します。 読み取り専用です。 グラフ タイトルの表示されていない場合は null です。 読み取り専用です。|1.7|
|[返すプロパティ](/javascript/api/excel/excel.charttitle)|_メソッド_ > setFormula(formula: string)|A1 スタイルの表記を使用して、グラフ タイトルの数式を表す文字列値を設定します。|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_関係_> 境界線|色、linestyle、および重量が含まれているグラフ タイトルの境界線の形式を表します。 読み取り専用です。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > 下位|近似曲線を後方に延長された期間の数を表します。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > displayEquation|True を設定すると、近似曲線の方程式をグラフに表示されます。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > displayRSquared|True を設定すると、グラフの近似曲線の R 平方が表示されます。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > フォワード|近似曲線を前方に拡張する期間の数を表します。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > 切片|近似曲線の切片の値を表します。 数値または空の文字列 (の自動的な値) に設定できます。 返された値は、常に番号です。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > movingAveragePeriod|移動平均の種類の近似曲線は、グラフの近似曲線の期間を表します。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > name|近似曲線の名前を表します。 文字列値に設定することができます。 または null 値は自動的な値に設定することができます。 返された値が文字列では常に|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_ > polynomialOrder|近似曲線の多項式のタイプには、グラフの近似曲線の順序を表します。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_プロパティ_> タイプ|グラフの種類を表します。 値を指定できます。 線形近似、指数、対数、移動平均、多項式近似、電源です。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_関係_> の形式|グラフの近似曲線の書式を表します。 読み取り専用です。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_メソッド_> delete()|近似曲線オブジェクトを削除します。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_プロパティ_ > items|ChartTrendline オブジェクトのコレクションです。 読み取り専用です。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_メソッド_ > add(type: string)|新しい近似曲線を近似曲線のコレクションに追加します。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_メソッド_> getCount()|コレクション内には、近似曲線の数を返します。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_メソッド_ > getItem(index: number)|項目の配列の挿入順序のインデックスでは、近似曲線オブジェクトを取得します。|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_関係_ > 行|グラフの線の書式設定を表します。値の取得のみ可能です。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_プロパティ_ > key|カスタム プロパティのキーを取得します。読み取り専用。読み取り専用。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_プロパティ_> タイプ|カスタム プロパティの値を取得または設定します。 読み取り専用です。 読み取り専用です。 使用可能な値: 数値、ブール値、日付、文字列、浮動小数点数です。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_プロパティ_ > value|カスタム プロパティの値を取得または設定します。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_メソッド_> delete()|カスタム プロパティを削除します。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_プロパティ_ > items|customProperty オブジェクトのコレクション。読み取り専用。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_メソッド_ > を追加 (キー: 値の文字列: オブジェクト)|新しいカスタム プロパティを作成、または既存のカスタム プロパティを設定します。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_メソッド_ > deleteAll()|このコレクション内のすべてのカスタム プロパティを削除します。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_メソッド_> getCount()|カスタム プロパティの数を取得します。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_メソッド_ > getItem(key: string)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。カスタム プロパティが存在しない場合にスローします。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_メソッド_ > getItemOrNullObject(key: string)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。カスタム プロパティが存在しない場合は、null オブジェクトを返します。|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_プロパティ_ > items|DataConnection オブジェクトのコレクションです。 読み取り専用です。|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_メソッド_ > refreshAll()|コレクション内のすべてのデータ接続を更新します。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > author|取得またはブックの作成者を設定します。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > category|取得またはブックのカテゴリを設定します。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > comments|取得またはブックのコメントを設定します。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > company|取得またはブックの会社を設定します。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > keywords|取得またはブックのキーワードを設定します。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > lastAuthor|ブックの最後の作成者を取得します。 読み取り専用です。 読み取り専用です。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > manager|取得またはブックのマネージャーを設定します。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > revisionNumber|ブックのリビジョン番号を取得します。 読み取り専用です。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > subject|取得またはブックの件名を設定します。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_プロパティ_ > title|取得またはブックのタイトルを設定します。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_リレーションシップ_ > creationDate|ブックの作成日を取得します。 読み取り専用です。 読み取り専用です。|1.7|
|[オートメーション](/javascript/api/excel/excel.documentproperties)|_関係_> カスタム|ブックのカスタム プロパティのコレクションを取得します。 読み取り専用です。 読み取り専用です。|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_プロパティ_> 数式|取得または名前付きアイテムの数式を設定します。  数式は、常に '=' 記号で始まります。|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_関係_ > arrayValues|値および名前付きの項目の種類を格納するオブジェクトを返します。 読み取り専用です。|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_プロパティ_ > 種類|名前付きアイテムの配列内の各項目の種類を表す読み取り専用です。 指定できる値は、 "Unknown"、"Empty"、"String"、 "Integer"、"Double"、"Boolean"、"Error" です。|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_プロパティ_ > values|名前付きアイテムの配列内の各項目の値を表します。 読み取り専用です。|1.7|
|[範囲](/javascript/api/excel/excel.range)|_プロパティ_ > isEntireColumn|現在の範囲が、全体の列である場合を表します。 読み取り専用です。|1.7|
|[範囲](/javascript/api/excel/excel.range)|_プロパティ_ > isEntireRow|現在の範囲が行全体である場合を表します。 読み取り専用です。|1.7|
|[範囲](/javascript/api/excel/excel.range)|_プロパティ_ > numberFormatLocal|ユーザーの言語で文字列として Excel の特定の範囲の数値形式のコードを表します。|1.7|
|[範囲](/javascript/api/excel/excel.range)|_プロパティ_ > style|現在の範囲のスタイルを表します。 これは、null または文字列のいずれかを返します。|1.7|
|[範囲](/javascript/api/excel/excel.range)|_メソッド_ > getAbsoluteResizedRange (numRows: 番号、numColumns: 数)|行と列の指定した数値が、現在の範囲のオブジェクトと同じ左上のセルでは、Range オブジェクトを取得します。|1.7|
|[範囲](/javascript/api/excel/excel.range)|_メソッド_ > getImage()|範囲は、base64 でエンコードされたイメージとしてレンダリングします。|1.7|
|[範囲](/javascript/api/excel/excel.range)|_メソッド_ > getSurroundingRegion()|この範囲の左上のセルの周囲の領域を表す Range オブジェクトを返します。 周囲の領域は、空白行と、この範囲を基準にして空白の列の任意の組み合わせで囲まれた範囲です。|1.7|
|[範囲](/javascript/api/excel/excel.range)|_メソッド_ > showCard()|値の豊富なコンテンツがある場合は、カードが、アクティブ セルを表示します。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_プロパティ_> textOrientation|取得または範囲内のすべてのセルのテキストの向きを設定します。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_プロパティ_ > useStandardHeight|Range オブジェクトの行の高さがシートの標準の高さに等しいかどうかを決定します。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_プロパティ_ > useStandardWidth|Range オブジェクトの列幅がシートの標準の幅に等しいかどうかを決定します。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_プロパティ_ > address|ハイパーリンクの url ターゲットを表します。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_プロパティ_ > ドキュメント.|ドキュメントを表します。 ハイパーリンクのターゲットです。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_プロパティ_ > ヒント|ハイパーリンクでポイントされたときに表示される文字列を表します。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_プロパティ_ > 示します|範囲内のほとんどのセルの左上に表示される文字列を表します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > addIndent|セル内のテキストの配置が均等に設定されている場合かどうかはテキストを自動的にインデントを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_> 自動インデント|セル内のテキストの配置が均等に設定されている場合かどうかはテキストを自動的にインデントを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > 組み込み|スタイルが組み込みのスタイルを示します。 読み取り専用です。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > formulaHidden|ワークシートが保護されている場合、数式を非表示にするかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_> horizontalAlignment|スタイルの水平方向の配置を表します。 使用可能な値: 一般に、左、中央、右、塗り、両端揃え、CenterAcrossSelection、分散します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includeAlignment|スタイルに、自動インデント、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel、TextOrientation プロパティが含まれるかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includeBorder|スタイルに、カラー、ColorIndex、LineStyle、および太さの罫線のプロパティが含まれるかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includeFont|True の場合、選択されているセルまたはセル範囲のスタイルとして定義されている書式は、フォントの設定 ( Background 、 Bold 、 Color 、 ColorIndex 、 FontStyle 、 Italic 、 Name 、 Size 、 Strikethrough 、 Subscript 、 Superscript 、および Underline プロパティ) を含みます。値の取得および設定が可能です。ブール型 ( Boolean ) の値を使用します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includeNumber|スタイルに数字が含まれるかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includePatterns|スタイル、色、ColorIndex、InvertIfNegative、パターン、PatternColor、PatternColorIndex 内部プロパティが含まれるかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > includeProtection|スタイルになるとき、および [ロックの保護プロパティが含まれるかどうかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > indentLevel|スタイルのインデントのレベルを示す 0 から 250 の整数です。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > locked|オブジェクトがロックされているワークシートが保護されていることを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > name|スタイルの名前です。 読み取り専用です。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > numberFormat|スタイルの表示形式の書式コードです。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_> numberFormatLocal|スタイルの表示形式のローカライズされた書式コードです。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > 印刷の向き|スタイルのテキストの向きです。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > 読む|スタイルの読み取り順序です。 使用可能な値: コンテキスト、LeftToRight、RightToLeft です。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > shrinkToFit|使用可能な列幅に収まるように自動的に文字列を縮小する場合は  True を返します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_> textOrientation|スタイルのテキストの向きです。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_> verticalAlignment|スタイルの垂直方向の配置を表します。 使用可能な値: 上部、中央、下、両端揃え、分散します。|1.7|
|[style](/javascript/api/excel/excel.style)|_プロパティ_ > wrapText|オブジェクト内のテキストを折り返すかを示します。|1.7|
|[style](/javascript/api/excel/excel.style)|_関係_ > 罫線|4 つの枠線のスタイルを表す 4 つの Border オブジェクトの枠線のコレクションです。 読み取り専用です。|1.7|
|[style](/javascript/api/excel/excel.style)|_関係_> 塗りつぶし|スタイルの塗りつぶし。 読み取り専用です。|1.7|
|[style](/javascript/api/excel/excel.style)|_リレーションシップ_ > font|スタイルのフォントを表す Font オブジェクトを返します。 読み取り専用です。|1.7|
|[style](/javascript/api/excel/excel.style)|_メソッド_> delete()|このスタイルを削除します。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_プロパティ_ > items|スタイル オブジェクトのコレクションです。 読み取り専用です。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_メソッド_ > add(name: string)]|新しいスタイルをコレクションに追加します。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)||||UNTRANSLATED_CONTENT_START|||_Method_ > getItem(name: string)|||UNTRANSLATED_CONTENT_END||||名前でスタイルを取得します。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_ > address|特定のワークシート上のテーブルの変更された領域を表すアドレスを取得します。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_ > changeType|Changed イベントがトリガーされる方法を表す変更の種類を取得します。 使用可能な値: 他のユーザー、RangeEdited、RowInserted、RowDeleted、ColumnInserted、ColumnDeleted、CellInserted、CellDeleted。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_> ソース|イベントのソースを取得します。 値を指定できます: ローカル、リモートです。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_ > tableId|データが変更されるテーブルの id を取得します。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。 使用可能な値: WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_プロパティ_> worksheetId|ワークシートのデータが変更されたの id を取得します。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_プロパティ_ > address|特定のワークシート上のテーブルの選択した領域を表す範囲のアドレスを取得します。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_プロパティ_ > isInsideTable|アドレスは IsInsideTable が false の場合は、無効に選択範囲が表内にある場合を示します。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_プロパティ_> tableId|選択範囲が変更されるテーブルの id を取得します。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。 使用可能な値: WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_プロパティ_> worksheetId|ワークシートの選択範囲が変更されたの id を取得します。|1.7|
|[ブック](/javascript/api/excel/excel.workbook)|_プロパティ_ > name|ブック名を取得します。 読み取り専用です。|1.7|
|[ブック](/javascript/api/excel/excel.workbook)|_関係_ > dataConnections|ブック内のすべてのデータ接続を更新します。 読み取り専用です。|1.7|
|[ブック](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > properties|ブックのプロパティを取得します。 読み取り専用です。|1.7|
|[ブック](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > protection|ブックのブックの保護オブジェクトを返します。 読み取り専用です。|1.7|
|[ブック](/javascript/api/excel/excel.workbook)|_関係_ > スタイル|ブックに関連付けられているスタイルのコレクションを表します。 読み取り専用です。|1.7|
|[ブック](/javascript/api/excel/excel.workbook)|_メソッド_ > getActiveCell()|ブックからには、現在アクティブなセルを取得します。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_プロパティ_ > protected|ブックが保護されているかどうかを示します。 読み取り専用。 読み取り専用です。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_メソッド_ > protect(password: string)|ブックを保護します。 ブックが保護されている場合は失敗します。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_メソッド_ > unprotect(password: string)|ブックの保護を解除します。|1.7|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_プロパティ_ > グリッド線|取得または、ワークシートの枠線] のフラグを設定します。|1.7|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_プロパティ_ > 見出し|取得または、ワークシートの見出しのフラグを設定します。|1.7|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_プロパティ_> showHeadings|取得または、ワークシートの見出しのフラグを設定します。|1.7|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_プロパティ_ > standardHeight|ワークシート内のすべての行の標準 (既定) の高さをポイント単位で返します。値の取得のみ可能です。倍精度浮動小数点型 ( Double ) の値を使用します。 読み取り専用です。|1.7|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_プロパティ_ > standardWidth|ワークシート内のすべての列の標準 (既定) の幅を設定します。値の取得および設定が可能です。倍精度浮動小数点型 ( Double ) の値を使用します。|1.7|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_プロパティ_ > tabColor|取得または、ワークシートのタブの色を設定します。|1.7|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_関係_ > 場合|ワークシート上の固定されたウィンドウ枠を操作するために使用するオブジェクトを取得する読み取り専用です。|1.7|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_メソッド_ > コピー (positionType: WorksheetPositionType、relativeTo: ワークシート)|ワークシートをコピーして、指定した位置に配置します。 コピーするワークシートを返します。|1.7|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_メソッド_ > getRangeByIndexes (startRow: 番号、startColumn: 数、行数: 番号、列の数: 数値)|は、特定の行インデックスと列インデックスから開始し、一定数の行と列にわたる、range オブジェクトを取得します。|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。 使用可能な値: WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_プロパティ_> worksheetId|アクティブなワークシートの id を取得します。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_プロパティ_> ソース|イベントのソースを取得します。 値を指定できます: ローカル、リモートです。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。 使用可能な値: WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_プロパティ_> worksheetId|ワークシートをブックに追加の id を取得します。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_プロパティ_ > address|特定のワークシートの変更された領域を表す範囲のアドレスを取得します。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_プロパティ_> changeType|Changed イベントがトリガーされる方法を表す変更の種類を取得します。 使用可能な値: 他のユーザー、RangeEdited、RowInserted、RowDeleted、ColumnInserted、ColumnDeleted、CellInserted、CellDeleted。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_プロパティ_> ソース|イベントのソースを取得します。 値を指定できます: ローカル、リモートです。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。 使用可能な値: WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_プロパティ_> worksheetId|ワークシートのデータが変更されたの id を取得します。|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。 使用可能な値: WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_プロパティ_> worksheetId|非アクティブ化するワークシートの id を取得します。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_プロパティ_> ソース|イベントのソースを取得します。 値を指定できます: ローカル、リモートです。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。 使用可能な値: WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_プロパティ_> worksheetId|ブックから削除するワークシートの id を取得します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > freezeAt (frozenRange: 文字列または範囲)|作業中のワークシート ビューで保持されているセルを設定します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > freezeColumns(count: number)|埋め込みのワークシートの最初の列を固定します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > freezeRows(count: number)|上でワークシートの行を固定します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > getLocation()|作業中のワークシート ビューで固定されたセルを表す範囲を取得します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > getLocationOrNullObject()|作業中のワークシート ビューで固定されたセルを表す範囲を取得します。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_メソッド_ > unfreeze()|ワークシート内のすべての固定されたウィンドウ枠を削除します。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowEditObjects|オブジェクトを編集できるようにワークシートの保護オプションを表します。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_プロパティ_ > allowEditScenarios|シナリオを編集できないようにワークシートの保護オプションを表します。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_関係_ > selectionMode|選択モードのワークシートの保護オプションを表します。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_プロパティ_ > address|特定のワークシートの選択した領域を表す範囲のアドレスを取得します。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_プロパティ_> タイプ|イベントの種類を取得します。 使用可能な値: WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_プロパティ_> worksheetId|ワークシートの選択範囲が変更されたの id を取得します。|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>Excel JavaScript API 1.6 の新機能 

### <a name="conditional-formatting"></a>条件付き書式

範囲の条件付き書式設定をについて説明します。 次の種類の条件付き書式を設定することができます。

* カラー スケール
* データ バー
* アイコン セット
* カスタム

さらに、

* 条件付き書式が適用された範囲を返す。 
* 条件付き書式を削除する。 
* 優先順位機能と stopifTrue 機能を提供する。 
* 指定した範囲のすべての条件付き書式のコレクションを取得する。 
* 現在指定している範囲でアクティブなすべての条件付き書式をクリアする。 

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[アプリケーション](/javascript/api/excel/excel.application)|_メソッド_ > suspendApiCalculationUntilNextSync()|次の "context.sync()" が呼び出されるまで、計算を中断します。設定されると、依存関係が確実に伝達されるようにブックを再計算するのは開発者の責任です。|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_関係_> の形式|Format オブジェクト (条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用です。|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_関係_> ルール|この条件付き書式の Rule オブジェクトを表します。|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_プロパティ_ > threeColorScale|true の場合、カラー スケールのポイントは 3 つ (最小、中間値、最大) になり、それ以外の場合は 2 つ (最小、最大) になります。読み取り専用です。|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_リレーションシップ_ > criteria|カラー スケールの条件。2 ポイントのカラー スケールを使用する場合、中間値はオプションです。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_プロパティ_> formula1|条件付き書式ルールを評価するために必要な場合、数式。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_プロパティ_> formula2|条件付き書式ルールを評価するために必要な場合、数式。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_プロパティ_ > operator|テキストの条件付き書式の演算子です。使用可能な値は次のとおりです。Invalid、Between、NotBetween、EqualTo、NotEqualTo、GreaterThan、LessThan、GreaterThanOrEqual、LessThanOrEqual。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_関係_ > 最大|最大ポイントのカラー スケール条件。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_関係_ > 中間点|カラー スケールが 3 色スケールの場合のカラー スケール条件の中間値。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_関係_ > 最小|最小ポイントのカラー スケール条件。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_プロパティ_ > color|カラー スケールの色の HTML カラー コード表記。たとえば、#FF0000 は赤を表します。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_プロパティ_> 数式|数値、数式、(型が LowestValue の場合は) null。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_プロパティ_> タイプ|アイコンの条件式は次のものに基づいています。使用可能な値は次のとおりです。Invalid、LowestValue、HighestValue、Number、Percent、Formula、Percentile。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_プロパティ_ > 境界線色|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_プロパティ_ > れました。|塗りつぶしの色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_プロパティ_ > matchPositiveBorderColor|負の DataBar に正の DataBar と同じ枠線の色があるかどうかを表すブール値。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_プロパティ_ > matchPositiveFillColor|負の DataBar に正の DataBar と同じ塗りつぶしの色があるかどうかを表すブール値。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_プロパティ_> 境界線色|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_プロパティ_> れました。|塗りつぶしの色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_プロパティ_ > gradientFill|DataBar のグラデーションがあるかどうかのブール値の表記。|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_プロパティ_> 数式|databar のルールを評価するために必要な場合、数式。|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_プロパティ_> タイプ|databar のルールの種類。使用可能な値は次のとおりです。LowestValue、HighestValue、Number、Percent、Formula、Percentile、Automatic。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_プロパティ_ > id|現在の ConditionalFormatCollection 内で条件付き書式の優先度です。 読み取り専用です。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_プロパティ_ > 優先順位|この条件付き書式が現在存在する、条件付き書式のコレクション内の優先度 (またはインデックス)。これも変更します|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_プロパティ_ > stopIfTrue|この条件付き書式の条件が満たされた場合、優先順位の低い書式はそのセルに影響を及ぼしません。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_プロパティ_> タイプ|条件付き書式の種類。一度に 1 つのみ設定できます。読み取り専用。読み取り専用。使用可能な値は次のとおりです。Custom、DataBar、ColorScale、IconSet。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > cellValue|現在の条件付き書式が CellValue 型の場合、セル値の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > cellValueOrNullObject|現在の条件付き書式が CellValue 型の場合、セル値の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > colorScale|現在の条件付き書式が ColorScale 型の場合、ColorScale の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > colorScaleOrNullObject|現在の条件付き書式が ColorScale 型の場合、ColorScale の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_> カスタム|現在の条件付き書式がカスタム型の場合、カスタムの条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > customOrNullObject|現在の条件付き書式がカスタム型の場合、カスタムの条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > dataBar|現在の条件付き書式がデータ バーの場合、データ バーのプロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > dataBarOrNullObject|現在の条件付き書式がデータ バーの場合、データ バーのプロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > iconSet|現在の条件付き書式が IconSet 型の場合、IconSet の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > iconSetOrNullObject|現在の条件付き書式が IconSet 型の場合、IconSet の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > プリセット|above averagebelow averageunique valuescontains blanknonblankerrornoerror properties などの事前に設定された条件付き書式を返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > presetOrNullObject|above averagebelow averageunique valuescontains blanknonblankerrornoerror properties などの事前に設定された条件付き書式を返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > textComparison|現在の条件付き書式がテキスト型の場合、特定のテキストの条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > textComparisonOrNullObject|現在の条件付き書式がテキスト型の場合、特定のテキストの条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > topBottom|現在の条件付き書式が TopBottom 型の場合、TopBottom の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_関係_ > topBottomOrNullObject|現在の条件付き書式が TopBottom 型の場合、TopBottom の条件付き書式プロパティを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_メソッド_> delete()|この条件付き書式を削除します。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_メソッド_ > getRange()|条件付き書式が適用される範囲を返します。範囲が連続していない場合は null オブジェクトを返します。読み取り専用。|1.6|
|[条件](/javascript/api/excel/excel.conditionalformat)|_メソッド_ > getRangeOrNullObject()|条件付き書式が適用される範囲を返します。範囲が連続していない場合は null オブジェクトを返します。読み取り専用。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_プロパティ_ > items|conditionalFormat オブジェクトのコレクション。読み取り専用です。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_メソッド_> add(type: string)|最高度の優先順位でコレクションに新しい条件付き書式を追加します。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_メソッド_ > clearAll()|現在指定している範囲でアクティブなすべての条件付き書式をクリアする。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_メソッド_> getCount()|ブック内の条件付き書式の数を返します。読み取り専用です。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_メソッド_ > getItem(id: string)|指定された ID の条件付き書式を返します。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_メソッド_> getItemAt(index: number)|指定されたインデックスに条件付き書式を返します。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_プロパティ_> 数式|条件付き書式ルールを評価するために必要な場合、数式。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_プロパティ_ > formulaLocal|ユーザーの言語で条件付き書式ルールを評価するために必要な場合、数式。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_プロパティ_ > formulaR1C1|R1C1 形式の表記法で条件付き書式ルールを評価するために必要な場合、数式。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_プロパティ_> 数式|種類によっては数値または数式。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_プロパティ_ > operator|アイコン条件付き書式のそれぞれのルールの種類に対する GreaterThan または GreaterThanOrEqual。使用可能な値は次のとおりです。Invalid、GreaterThan、GreaterThanOrEqual。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_関係_ > customIcon|既定の IconSet と異なる場合は現在の条件のカスタム アイコン、そうでない場合は null が返されます。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_リレーションシップ_ > type|アイコンの条件式は次のものに基づいています。|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_プロパティ_ > 条件|条件付き書式の条件です。使用可能な値は次のとおりです。Invalid、Blanks、NonBlanks、Errors、NonErrors、Yesterday、Today、Tomorrow、LastSevenDays、LastWeek、ThisWeek、NextWeek、LastMonth、ThisMonth、NextMonth、AboveAverage、BelowAverage、EqualOrAboveAverage、EqualOrBelowAverage、OneStdDevAboveAverage、OneStdDevBelowAverage、TwoStdDevAboveAverage、TwoStdDevBelowAverage、ThreeStdDevAboveAverage、ThreeStdDevBelowAverage、UniqueValues、DuplicateValues。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_プロパティ_ > color|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_プロパティ_ > id|罫線の識別子を表します。読み取り専用です。使用可能な値は次のとおりです。EdgeTop、EdgeBottom、EdgeLeft、EdgeRight。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_プロパティ_ > sideIndex|罫線の特定の辺を表す定数値。読み取り専用です。使用可能な値は次のとおりです。EdgeTop、EdgeBottom、EdgeLeft、EdgeRight。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_プロパティ_ > style|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。使用可能な値は次のとおりです。None、Continuous、Dash、DashDot、DashDotDot、Dot、Double、SlantDashDot。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_プロパティ_> カウント|コレクションに含まれる境界線オブジェクトの数。読み取り専用です。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_プロパティ_ > items|conditionalRangeBorder オブジェクトのコレクション。読み取り専用です。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_関係_ > 下|上罫線を取得します。読み取り専用です。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_関係_ > 左|上罫線を取得します。読み取り専用です。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_関係_ > 右|上罫線を取得します。読み取り専用です。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_関係_ > トップ|上罫線を取得します。読み取り専用です。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_メソッド_ > getItem(index: string)|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_メソッド_> getItemAt(index: number)|オブジェクトのインデックスを使用して、境界線オブジェクトを取得します。|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_プロパティ_ > color|塗りつぶしの色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_メソッド_ > clear()|塗りつぶしをリセットします。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_プロパティ_ > 太字|フォントの太字の状態を表します。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_プロパティ_ > color|テキストの色の HTML カラー コード表記。たとえば、#FF0000 は赤を表します。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_プロパティ_ > 斜体|フォントの斜体の状態を表します。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_プロパティ_ > 取り消し線|フォントの取り消し線の状態を表します。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_プロパティ_ > 下線|フォントに適用する下線の種類。使用可能な値は次のとおりです。None、Single、Double。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_メソッド_> clear()|フォントの書式設定をリセットします。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_プロパティ_ > numberFormat|指定した範囲の Excel の数値書式コードを表します。null が渡された場合はクリアします。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_関係_> 罫線|条件付き書式範囲全体に適用する境界線オブジェクトのコレクション。読み取り専用です。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_関係_> 塗りつぶし|条件付き書式範囲全体に定義された塗りつぶしオブジェクトを返します。読み取り専用です。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_リレーションシップ_ > font|条件付き書式範囲全体に定義されたフォント オブジェクトを返します。読み取り専用です。|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_プロパティ_ > operator|テキストの条件付き書式の演算子です。使用可能な値は次のとおりです。Invalid、Contains、NotContains、BeginsWith、EndsWith。|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_プロパティ_ > text|条件付き書式のテキスト値。|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_プロパティ_ > ランク|数値のランクに対する 1 から 1000、またはパーセントのランクに対する 1 から 100 のランク。|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_プロパティ_> タイプ|上位または下位のランクに基づく値を書式設定します。使用可能な値は次のとおりです。Invalid、TopItems、TopPercent、BottomItems、BottomPercent。|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_関係_> の形式|Format オブジェクト (条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用です。|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_関係_> ルール|この条件付き書式の Rule オブジェクトを表します。読み取り専用です。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_プロパティ_ > axisColor|軸の線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_プロパティ_ > axisFormat|Excel のデータ バーの軸を決定する方法を表します。使用可能な値は次のとおりです。Automatic、None、CellMidPoint。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_プロパティ_ > barDirection|データ バーの図の基準となる方向を表します。使用可能な値は次のとおりです。Context、LeftToRight、RightToLeft。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_プロパティ_ > showDataBarOnly|True の場合、データ バーが適用されているセルの値を非表示にします。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_関係_ > lowerBoundRule|データ バーの下限値 (および該当する場合はその計算方法) を構成するルール。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_関係_ > negativeFormat|Excel データ バーの軸の左側のすべての値を表します。読み取り専用です。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_関係_ > positiveFormat|Excel データ バーの軸の右側のすべての値を表します。読み取り専用です。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_関係_ > upperBoundRule|データ バーの上限値 (および該当する場合はその計算方法) を構成するルール。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_プロパティ_ > reverseIconOrder|True の場合、IconSet のアイコンの順序を反転します。カスタム アイコンを使用する場合には、この設定は適用できません。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_プロパティ_ > showIconOnly|True の場合、値を非表示にし、アイコンのみを表示します。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_プロパティ_ > style|設定した場合、条件付き書式の IconSet オプションを表示します。使用可能な値は次のとおりです。Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_リレーションシップ_ > criteria|条件付きアイコンの規則と潜在的なカスタム アイコンの、抽出条件と IconSets の配列。最初の条件として、カスタム アイコンのみを変更することができますが、設定された場合に、型、数式、演算子は無視されます。|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_関係_> の形式|Format オブジェクト (条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用です。|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_関係_> ルール|条件付き書式のルールです。|1.6|
|[範囲](/javascript/api/excel/excel.range)|_関係_ > conditionalFormats|範囲を交差する ConditionalFormats のコレクション。読み取り専用です。|1.6|
|[範囲](/javascript/api/excel/excel.range)|_メソッド_ > calculate()|ワークシート上のセルの範囲を計算します。|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_関係_> の形式|Format オブジェクト (条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用です。|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_関係_> ルール|条件付き書式のルールです。|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_関係_> の形式|Format オブジェクト (条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用です。|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_関係_> ルール|TopBottom の条件付き書式の条件。|1.6|
|[ブック](/javascript/api/excel/excel.workbook)|_関係_ > internalTest|内部使用専用です。読み取り専用です。|1.6|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_メソッド_ > calculate(markAllDirty: bool)|ワークシート上のすべてのセルを計算します。|1.6|

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
* `getNextColumn()` |||UNTRANSLATED_CONTENT_START|||and `getPreviousColumn()`, `getLast() on table column.|||UNTRANSLATED_CONTENT_END|||
* `getActiveWorksheet()` |||UNTRANSLATED_CONTENT_START|||on the workbook.|||UNTRANSLATED_CONTENT_END|||
* `getRange(address: string)` |||UNTRANSLATED_CONTENT_START|||off of workbook.|||UNTRANSLATED_CONTENT_END|||
* `getBoundingRange(ranges: )` は、指定した範囲を包含する、最小の範囲オブジェクトを取得します。たとえば、"B2:C5" から "D10:E15" までの隣接する範囲は、"B2:E15" になります。
* `getCount()` |||UNTRANSLATED_CONTENT_START|||on various collections such as named item, worksheet, table, etc. to get number of items in a collection.|||UNTRANSLATED_CONTENT_END||| `workbook.worksheets.getCount()`
* `getFirst()` |||UNTRANSLATED_CONTENT_START|||and `getLast()` and get last on various collection such as tworksheet, able column, chart points, range view collection.|||UNTRANSLATED_CONTENT_END|||
* `getNext()` |||UNTRANSLATED_CONTENT_START|||and `getPrevious()` on worksheet, table column collection.|||UNTRANSLATED_CONTENT_END|||
* `getRangeR1C1()` は、特定の行インデックスと列インデックスから開始し、一定数の行と列にわたる、range オブジェクトを取得します。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|_プロパティ_ > id|カスタム XML パーツの ID。読み取り専用です。|1.5|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|_プロパティ_ > 名前空間 Uri|カスタム XML パーツの名前空間 URI。読み取り専用です。|1.5|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|_メソッド_> delete()|カスタム XML パーツを削除します。|1.5|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|_メソッド_ > getXml()|カスタム XML パーツのすべての XML コンテンツを取得します。|1.5|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|_メソッド_ > setXml(xml: string)|カスタム XML パーツのすべての XML コンテンツを設定します。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_プロパティ_ > items|customXmlPart オブジェクトのコレクション。読み取り専用です。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_メソッド_ > add(xml: string)|ブックに新しいカスタム XML 部分を追加します。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_メソッド_ > getByNamespace(namespaceUri: string)|名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_メソッド_> getCount()|コレクション内にある CustomXml パーツの数を取得します。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_メソッド_> getItem(id: string)|ID に基づいて、カスタム XML パーツを取得します。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_メソッド_ > getItemOrNullObject(id: string)|ID に基づいて、カスタム XML パーツを取得します。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_プロパティ_ > items|customXmlPartScoped オブジェクトのコレクション。読み取り専用です。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_メソッド_> getCount()|コレクション内にある CustomXML パーツの数を取得します。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_メソッド_> getItem(id: string)|ID に基づいて、カスタム XML パーツを取得します。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_メソッド_> getItemOrNullObject(id: string)|ID に基づいて、カスタム XML パーツを取得します。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_メソッド_ > getOnlyItem()|コレクションに含まれる項目が 1 つだけの場合、このメソッドはこれを返します。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_メソッド_ > getOnlyItemOrNullObject()|コレクションに含まれる項目が 1 つだけの場合、このメソッドはこれを返します。|1.5|
|[ブック](/javascript/api/excel/excel.workbook)|_関係_ > 空|このブックに含まれる、カスタム XML パーツのコレクションを表します。読み取り専用です。|1.5|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_メソッド_ > getNext(visibleOnly: bool)|これに続くワークシートを取得します。これに続くワークシートがない場合、このメソッドによってエラーがスローされます。|1.5|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_メソッド_ > getNextOrNullObject(visibleOnly: bool)|これに続くワークシートを取得します。これに続くワークシートがない場合、このメソッドによって null オブジェクトが返されます。|1.5|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_メソッド_ > getPrevious(visibleOnly: bool)|この前に来るワークシートを取得します。この前のワークシートがない場合、このメソッドによってエラーがスローされます。|1.5|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_メソッド_ > getPreviousOrNullObject(visibleOnly: bool)|この前に来るワークシートを取得します。この前のワークシートがない場合、このメソッドによって null オブジェクトが返されます。|1.5|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_メソッド_ > getFirst(visibleOnly: bool)|コレクション内の最初のワークシートを取得します。|1.5|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_メソッド_ > getLast(visibleOnly: bool)|コレクション内の最後のワークシートを取得します。|1.5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Excel JavaScript API 1.4 の新機能
要件セット 1.4 の Excel JavaScript API に新たに追加された点は次のとおりです。

### <a name="named-item-add-and-new-properties"></a>名前付きアイテムの追加と新しいプロパティ

新しいプロパティ:

* `comment`
* `scope` ワークシートまたはブックの対象になるアイテム
* `worksheet` 名前付きアイテムの対象になるワークシートを返します。

新しいメソッド:

* `add(name: string, reference: Range or string, comment: string)`新しい名前を指定したスコープのコレクションに追加します。
* `addFormulaLocal(name: string, formula: string, comment: string)` ユーザーのロケールを数式に使用して、新しい名前を指定したスコープのコレクションに追加します。

### <a name="settings-api-in-in-excel-namespace"></a>Excel の名前空間での Setting API

[Setting](/javascript/api/excel/excel.setting) オブジェクトは、ドキュメントに永続化されている設定のキーと値のペアを表します。ここでは、Excel の名前空間に設定関連の API を追加しました。これは純粋な新機能は提供しませんが、これにより約束ベースのバッチ API 構文を維持することが容易になり、Excel 関連タスクの共通 API に対する依存を減らすことができます。

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
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_> getCount()|コレクション内にあるバインドの数を取得します。|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_> getItemOrNullObject(id: string)|ID によってバインド オブジェクトを取得します。バインディング オブジェクトが存在しない場合は null オブジェクトを返します。|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_メソッド_> getCount()|ワークシート上のグラフの数を返します。|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)||||UNTRANSLATED_CONTENT_START|||_Method_ > getItemOrNullObject(name: string)|||UNTRANSLATED_CONTENT_END||||グラフ名を使用してグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_メソッド_> getCount()|系列内にあるグラフのポイントの数を取得します。|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_メソッド_> getCount()|コレクション内にあるデータ系列の数を取得します。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_プロパティ_ > comment|この名前に関連付けられているコメントを表します。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_プロパティ_ > scope|名前がブックを対象にしているのか、特定のワークシートを対象にしているのかを示します。読み取り専用です。使用可能な値は次のとおりです。Equal、Greater、GreaterEqual、Less、LessEqual、NotEqual。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_リレーションシップ_ > worksheet|名前付きのアイテムの対象になるワークシートを返します。アイテムがブックを対象にしている場合は、エラーをスローします。読み取り専用です。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_リレーションシップ_ > worksheetOrNullObject|名前付きのアイテムの対象になるワークシートを返します。アイテムがブックを対象にしている場合は、null オブジェクトを返します。読み取り専用です。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_メソッド_> delete()|指定された名前を削除します。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_メソッド_> getRangeOrNullObject()|名前に関連付けられている範囲オブジェクトを返します。名前付きアイテムの型が範囲でない場合は、null オブジェクトを返します。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_メソッド_ > 追加 (名前: 参照の文字列: 範囲またはコメントの文字列: 文字列)|新しい名前を指定したスコープのコレクションに追加します。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_メソッド_ > addFormulaLocal (名前: 文字列、数式: コメントの文字列: 文字列)|ユーザーのロケールを数式に使用して、新しい名前を指定したスコープのコレクションに追加します。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_メソッド_> getCount()|コレクション内の名前付きアイテムの数を取得します。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)||||UNTRANSLATED_CONTENT_START|||_Method_ > getItemOrNullObject(name: string)|||UNTRANSLATED_CONTENT_END||||名前を使用して、nameditem オブジェクトを取得します。nameditem オブジェクトが存在しない場合は null オブジェクトを返します。|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_メソッド_> getCount()|コレクション内のピボット テーブルの数を取得します。|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)||||UNTRANSLATED_CONTENT_START|||_Method_ > getItemOrNullObject(name: string)|||UNTRANSLATED_CONTENT_END||||名前を使用してピボットテーブルを取得します。PivotTable が存在しない場合は null オブジェクトを返します。|1.4|
|[範囲](/javascript/api/excel/excel.range)|_メソッド_ > getIntersectionOrNullObject (anotherRange: 文字列または範囲)|指定した範囲の長方形の交差部分を表す Range オブジェクトを取得します。交差部分が見つからない場合は、null オブジェクトを返します。|1.4|
|[範囲](/javascript/api/excel/excel.range)|_メソッド_ > getUsedRangeOrNullObject(valuesOnly: bool)|指定した範囲オブジェクトのうち使用されている範囲を返します。範囲内に使用済みのセルがない場合、この関数は null オブジェクトを返します。|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_メソッド_> getCount()|コレクション内にある RangeView オブジェクトの数を取得します。|1.4|
|[文字列](/javascript/api/excel/excel.setting)|_プロパティ_ > key|Setting の ID を表すキーを返します。読み取り専用。|1.4|
|[文字列](/javascript/api/excel/excel.setting)|_プロパティ_ > value|この設定に格納されている値を表します。|1.4|
|[文字列](/javascript/api/excel/excel.setting)|_メソッド_> delete()|設定を削除します。|1.4|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|_プロパティ_ > items|setting オブジェクトのコレクション。読み取り専用。|1.4|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_ > を追加 (キー: 文字列、値: (任意))|指定した設定をブックに設定または追加します。|1.4|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_> getCount()|コレクション内にある Setting の数を取得します。|1.4|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_> getItem(key: string)|キーから Setting エントリを取得します。|1.4|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_> getItemOrNullObject(key: string)|キーから Setting エントリを取得します。Setting が存在しない場合は null オブジェクトを返します。|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_リレーションシップ_ > settings|SettingsChanged イベントが発生したバインドを表す Setting オブジェクトを取得します。|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_メソッド_ > getCount()]|コレクション内のテーブルの数を取得します。|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_メソッド_ > getItemOrNullObject (キー: 数値または文字列)|名前または ID でテーブルを取得します。テーブルが存在しない場合は null オブジェクトを返します。|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_メソッド_> getCount()|表の列数を取得します。|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_メソッド_> getItemOrNullObject (キー: 数値または文字列)|名前または ID によって、列オブジェクトを取得します。列が存在しない場合は null オブジェクトを返します。|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_メソッド_> getCount()|表の行数を取得します。|1.4|
|[ブック](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > settings|ブックに関連付けられている Setting のコレクションを表します。読み取り専用。|1.4|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_リレーションシップ_ > names|現在のワークシートにスコープされている名前のコレクション。読み取り専用です。|1.4|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_メソッド_> getUsedRangeOrNullObject(valuesOnly: bool)|使用範囲とは、値または書式設定が割り当たっているすべてのセルを包含する最小の範囲です。ワークシート全体が空白の場合、この関数は null オブジェクトを返します。|1.4|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_メソッド_ > getCount(visibleOnly: bool)|コレクション内のワークシートの数を取得します。|1.4|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_メソッド_> getItemOrNullObject(key: string)|名前または ID を使用して、ワークシート オブジェクトを取得します。ワークシートが存在しない場合は null オブジェクトを返します。|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 の新機能

要件セット 1.3 の Excel JavaScript API に新しく追加された点は次のとおりです。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[バインディング](/javascript/api/excel/excel.binding)|_メソッド_> delete()|バインドを削除します。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_ > を追加 (範囲: 範囲または文字列、bindingType: id、文字列: 文字列)|特定の範囲に新しいバインドを追加します。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_ > addFromNamedItem (名前: 文字列、bindingType: id、文字列: 文字列)|ブック内の名前付きアイテムに基づいて新しいバインドを追加します。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_ > addFromSelection (bindingType: id、文字列: 文字列)|現在の選択範囲に基づいて新しいバインドを追加します。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_メソッド_ > getItemOrNull(id: string)|ID を使用してバインド オブジェクトを取得します。バインド オブジェクトが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_メソッド_ > getItemOrNull(name: string)|グラフ名を使用してグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_メソッド_> getItemOrNull(name: string)|nameditem オブジェクトを、名前を使用して取得します。nameditem オブジェクトが存在しない場合、返されたオブジェクトの isNull プロパティは true になります。|1.3|
|[ピボットテーブル](/javascript/api/excel/excel.pivottable)|_プロパティ_ > name|ピボットテーブルの名前。|1.3|
|[ピボットテーブル](/javascript/api/excel/excel.pivottable)|_リレーションシップ_ > worksheet|現在のピボットテーブルを含んでいるワークシート。読み取り専用。|1.3|
|[ピボットテーブル](/javascript/api/excel/excel.pivottable)|_メソッド_ > refresh()|ピボットテーブルを更新します。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_プロパティ_ > items|ピボットテーブル オブジェクトのコレクション。読み取り専用。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)||||UNTRANSLATED_CONTENT_START|||_Method_ > getItem(name: string)|||UNTRANSLATED_CONTENT_END||||名前を使用してピボットテーブルを取得します。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_メソッド_> getItemOrNull(name: string)|名前を使用してピボットテーブルを取得します。ピボットテーブルが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|1.3|
|[範囲](/javascript/api/excel/excel.range)|_メソッド_ > getIntersectionOrNull (anotherRange: 文字列または範囲)|指定した範囲の長方形の交差部分を表す Range オブジェクトを取得します。交差部分が見つからない場合は、null オブジェクトを返します。|1.3|
|[範囲](/javascript/api/excel/excel.range)|_メソッド_ > getVisibleView()|現在の範囲の表示されている行を表します。|1.3|
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
|[rangeView](/javascript/api/excel/excel.rangeview)|_メソッド_> getRange()|現在の RangeView に関連付けられている親の範囲を取得します。|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_プロパティ_ > items|rangeView オブジェクトのコレクション。読み取り専用。|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_メソッド_> getItemAt(index: number)|RangeView のインデックスから RangeView の行番号を取得します。0 を起点とする番号になります。|1.3|
|[文字列](/javascript/api/excel/excel.setting)|_プロパティ_ > key|Setting の ID を表すキーを返します。読み取り専用。|1.3|
|[文字列](/javascript/api/excel/excel.setting)|_メソッド_> delete()|設定を削除します。|1.3|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|_プロパティ_ > items|setting オブジェクトのコレクション。読み取り専用。|1.3|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_> getItem(key: string)|キーから Setting エントリを取得します。|1.3|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_ > getItemOrNull(key: string)|キーから Setting エントリを取得します。Setting が存在しない場合、返されたオブジェクトの isNull プロパティは true になります。|1.3|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|_メソッド_ > 設定 (キー: 値の文字列: 文字列)|指定した設定をブックに設定または追加します。|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_リレーションシップ_ > settingCollection|SettingsChanged イベントが発生したバインドを表す Setting オブジェクトを取得します。|1.3|
|[テーブル](/javascript/api/excel/excel.table)|_プロパティ_ > highlightFirstColumn|最初の列に特別な書式設定が含まれているかどうかを示します。|1.3|
|[テーブル](/javascript/api/excel/excel.table)|_プロパティ_ > highlightLastColumn|最後の列に特別な書式設定が含まれているかどうかを示します。|1.3|
|[テーブル](/javascript/api/excel/excel.table)|_プロパティ_ > showBandedColumns|テーブルを見やすくするため、奇数列を偶数列とは異なる方法で強調表示する書式設定にして、列を縞模様で表示するかどうかを示します。|1.3|
|[テーブル](/javascript/api/excel/excel.table)|_プロパティ_ > showBandedRows|テーブルを見やすくするため、奇数行を偶数行とは異なる方法で強調表示する書式設定にして、行を縞模様で表示するかどうかを示します。|1.3|
|[テーブル](/javascript/api/excel/excel.table)|_プロパティ_ > showFilterButton|フィルター ボタンを各列のヘッダーの上部に表示するかどうかを示します。これは、テーブルにヘッダー行が含まれている場合のみ設定できます。|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_メソッド_ > getItemOrNull (キー: 数値または文字列)|名前または ID を使用してテーブルを取得します。テーブルが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_メソッド_> getItemOrNull (キー: 数値または文字列)|名前または ID を使用して列オブジェクトを取得します。列が存在しない場合、返されたオブジェクトの isNull プロパティは true になります。|1.3|
|[ブック](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > pivotTables|ブックに関連付けられているピボットテーブルのコレクションを表します。読み取り専用。|1.3|
|[ブック](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > settings|ブックに関連付けられている Setting のコレクションを表します。読み取り専用。|1.3|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_リレーションシップ_ > pivotTables|ワークシートの一部になっているピボットテーブルのコレクション。読み取り専用。|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Excel JavaScript API 1.2 の新機能

要件セット 1.2 の Excel JavaScript API に新たに追加された点は次のとおりです。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[グラフ](/javascript/api/excel/excel.chart)|_プロパティ_ > id|コレクション内での位置を基にグラフを取得します。読み取り専用です。|1.2|
|[グラフ](/javascript/api/excel/excel.chart)|_リレーションシップ_ > worksheet|現在のグラフを含んでいるワークシート。読み取り専用。|1.2|
|[グラフ](/javascript/api/excel/excel.chart)|_メソッド_ > getImage (高さ: 幅の番号: 番号、fittingMode: 文字列)|指定したサイズに合わせてグラフを拡大、縮小することで、グラフを Base64 でエンコードされた画像としてレンダリングします。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_リレーションシップ_ > criteria|指定した列に現在適用されているフィルターです。読み取り専用です。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_ > apply(criteria: FilterCriteria)|指定した列に、指定されたフィルター条件を適用します。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_ > applyBottomItemsFilter(count: number)|指定した数の要素の列に "下位アイテム" フィルターを適用します。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_ > applyBottomPercentFilter(percent: number)]|指定した割合の要素の列に "下位パーセント" フィルターを適用します。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_ > applyCellColorFilter(color: string)|指定した色の列に "セルの色" フィルターを適用します。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_ > applyCustomFilter (抽出条件 1: 文字列、検索条件 2: 工程、文字列: 文字列)|指定した条件の文字列の列に "アイコン" フィルターを適用します。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_ > applyDynamicFilter(criteria: string)|列に "動的" フィルターを適用します。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_ > applyFontColorFilter(color: string)|指定した色の列に "フォントの色" フィルターを適用します。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_ > applyIconFilter(icon: Icon)|指定したアイコンの列に "アイコン" フィルターを適用します。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_ > applyTopItemsFilter(count: number)|指定した数の要素の列に "上位アイテム" フィルターを適用します。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_ > applyTopPercentFilter(percent: number)|指定した割合の要素の列に "上位パーセント" フィルターを適用します。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_ > applyValuesFilter (値: ())|指定した値の列に "値" フィルターを適用します。|1.2|
|[フィルター](/javascript/api/excel/excel.filter)|_メソッド_> clear()|指定した列のフィルターをクリアします。|1.2|
|[ごと](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > color|セルをフィルター処理するために使用する HTML カラー文字列。「CellColor」フィルターおよび「fontColor」フィルターと併用します。|1.2|
|[ごと](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > criterion1|データをフィルター処理するために使用する最初の条件。「カスタム」フィルター処理の場合には、演算子として使用されます。|1.2|
|[ごと](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > criterion2|データをフィルター処理するために使用する 2 番目の条件。「カスタム」フィルター処理の場合には、演算子としてのみ使用されます。|1.2|
|[ごと](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ >dynamicCriteria|この列に適用する Excel.DynamicFilterCriteria の動的条件。「動的」フィルター処理で使用します。使用可能な値は次のいずれかです。Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|1.2|
|[ごと](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > filterOn|値を表示したままにするかどうかを判別するために、フィルターで使用するプロパティ。使用可能な値は次のとおりです。BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|1.2|
|[ごと](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > operator|"カスタム" フィルター処理を使用するときに、条件 1 と条件 2 と結合との使用する演算子。使用可能な値は次のとおりです。And、Or。|1.2|
|[ごと](/javascript/api/excel/excel.filtercriteria)|_プロパティ_ > values|"値" フィルター処理の一部として使用する値のセット。|1.2|
|[ごと](/javascript/api/excel/excel.filtercriteria)|_リレーションシップ_ > icon|セルをフィルター処理するために使用するアイコン。「アイコン」フィルター処理で使用します。|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_プロパティ_ > date|データのフィルター処理に使用する ISO8601 形式の日付です。|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_プロパティ_ > specificity|データを保持するのに、日付をどの程度詳細に使用するか。たとえば、date が 2005-04-02 で "month" に設定した場合、フィルター操作では 2005 年 4 月の日付データを含むすべての行が保持されます。使用可能な値は次のとおりです。Year、Month、Day、Hour、Minute、Second。|1.2|
|[FormatProtection](/javascript/api/excel/excel.formatprotection)|_プロパティ_ > formulaHidden|Excel が範囲内のセルの数式を非表示にするかどうかを示します。null 値は、範囲全体に一様な数式非表示設定がないことを表します。|1.2|
|[FormatProtection](/javascript/api/excel/excel.formatprotection)|_プロパティ_ > locked|Excel がオブジェクト内のセルをロックするかどうかを示します。null 値は、範囲全体に一様なロック設定がないことを表します。|1.2|
|[アイコン](/javascript/api/excel/excel.icon)|_プロパティ_ > index|指定したセット内のアイコンのインデックスを表します。|1.2|
|[アイコン](/javascript/api/excel/excel.icon)|_プロパティ_ > set|アイコンがその一部であるセットを表します。使用可能な値は次のとおりです。Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.2|
|[範囲](/javascript/api/excel/excel.range)|_プロパティ_ > columnHidden|現在の範囲のすべての列が非表示になっているかどうかを表します。|1.2|
|[範囲](/javascript/api/excel/excel.range)|_プロパティ_ > formulasR1C1|R1C1 スタイル表記の数式を表します。|1.2|
|[範囲](/javascript/api/excel/excel.range)|_プロパティ_ > hidden|現在の範囲のすべてのセルが非表示になっているかどうかを表します。読み取り専用です。|1.2|
|[範囲](/javascript/api/excel/excel.range)|_プロパティ_ > rowHidden|現在の範囲のすべての行が非表示になっているかどうかを表します。|1.2|
|[範囲](/javascript/api/excel/excel.range)|_リレーションシップ_ > sort|現在の範囲について、範囲の並べ替えを表します。読み取り専用。|1.2|
|[範囲](/javascript/api/excel/excel.range)|_メソッド_ > merge(across: bool)|範囲内のセルをワークシートの 1 つの領域に結合します。|1.2|
|[範囲](/javascript/api/excel/excel.range)|_メソッド_ > unmerge()|範囲内のセルを結合解除して別々のセルにします。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_プロパティ_ > columnWidth|範囲内のすべての列の幅を取得または設定します。列の幅が均一でない場合は、null が返されます。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_プロパティ_ > rowHeight|範囲内のすべての行の高さを取得または設定します。行の高さが均一でない場合は、null が返されます。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_リレーションシップ_ > protection|範囲に対する書式保護オブジェクトを返します。読み取り専用です。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_メソッド_ > autofitColumns()|現在の列のデータに基づいて、現在の範囲の列の幅を最適な幅に変更します。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_メソッド_ > autofitRows()|現在の行のデータに基づいて、現在の範囲の行の高さを最適な高さに変更します。|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_プロパティ_ > address|現在の範囲の表示されている行を表します。|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_メソッド_ > 適用 (フィールド: SortField、matchCase: bool、hasHeaders: bool、印刷の向き: メソッドの文字列: 文字列)|並べ替え操作を実行します。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_プロパティ_ > ascending|昇順で並べ替えるかどうかを表します。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_プロパティ_ > color|並べ替えがフォントまたはセルの色で行われる場合に、条件の対象となる色を表します。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_プロパティ_ > dataOption|このフィールドのその他の並べ替えオプションを表します。使用可能な値は次のとおりです。Normal、TextAsNumber。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_プロパティ_ > key|条件の対象とする列 (または行。並べ替えの方向によって異なります) を表します。最初の列 (または行) からのオフセットとして表します。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_プロパティ_ > sortOn|この条件の並べ替えの種類を表します。使用可能な値は次のとおりです。Value、CellColor、FontColor、Icon。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_リレーションシップ_ > icon|並べ替えがセルのアイコンで行われる場合に、条件の対象となるアイコンを表します。|1.2|
|[テーブル](/javascript/api/excel/excel.table)|_リレーションシップ_ > sort|テーブル内の並べ替えを表します。読み取り専用。|1.2|
|[テーブル](/javascript/api/excel/excel.table)|_リレーションシップ_ > worksheet|現在のテーブルを含んでいるワークシート。読み取り専用です。|1.2|
|[テーブル](/javascript/api/excel/excel.table)|_メソッド_ > clearFilters()|現在テーブルに適用されているすべてのフィルターをクリアします。|1.2|
|[テーブル](/javascript/api/excel/excel.table)|_メソッド_ > convertToRange()|テーブルを通常の範囲のセルに変換します。すべてのデータが保持されます。|1.2|
|[テーブル](/javascript/api/excel/excel.table)|_メソッド_ > reapplyFilters()|現在テーブルにあるすべてのフィルターを再適用します。|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_リレーションシップ_ > filter|列に適用されるフィルターを取得します。読み取り専用です。|1.2|
|[だけを並べ替える](/javascript/api/excel/excel.tablesort)|_プロパティ_ > matchCase|大文字小文字の区別が、テーブルの最後の並べ替え操作に影響を与えたかどうかを表します。読み取り専用です。|1.2|
|[だけを並べ替える](/javascript/api/excel/excel.tablesort)|_プロパティ_ > method|テーブルの並べ替えで最後に使用した中国語文字の順序付け方法を表します。読み取り専用です。使用可能な値は次のとおりです。PinYin、StrokeCount。|1.2|
|[だけを並べ替える](/javascript/api/excel/excel.tablesort)|_リレーションシップ_ > fields|テーブルの最後の並べ替えに使用する現在の条件を表します。読み取り専用です。|1.2|
|[だけを並べ替える](/javascript/api/excel/excel.tablesort)|_メソッド_ > 適用 (フィールド: SortField、matchCase: ブール値、メソッド: 文字列)|並べ替え操作を実行します。|1.2|
|[だけを並べ替える](/javascript/api/excel/excel.tablesort)|_メソッド_> clear()|テーブルに現在設定されている並べ替えをクリアします。これにより表の順序が変更されることはありませんが、ヘッダーのボタンの状態がクリアされます。|1.2|
|[だけを並べ替える](/javascript/api/excel/excel.tablesort)|_メソッド_ > reapply()|テーブルに、現在の並べ替えパラメーターを再適用します。|1.2|
|[ブック](/javascript/api/excel/excel.workbook)|_リレーションシップ_ > functions|このブックを含む Excel アプリケーションのインスタンスを表します。読み取り専用。|1.2|
|[ワークシート](/javascript/api/excel/excel.worksheet)|_リレーションシップ_ > protection|ワークシートのシート保護オブジェクトを返します。読み取り専用です。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_プロパティ_ > protected|ワークシートが保護されているかどうかを示します。読み取り専用。読み取り専用。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_リレーションシップ_ > options|シートの保護のオプション。読み取り専用。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_メソッド_ > protect(options: WorksheetProtectionOptions)|ワークシートを保護します。ワークシートが保護されている場合は失敗します。|1.2|
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

## <a name="excel-javascript-api-11"></a>Excel の JavaScript API 1.1

OneNote JavaScript API 1.1 は、API の最初のバージョンです。 API について詳しくは、[OneNote JavaScript API](/javascript/api/excel) リファレンスのトピックをご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
