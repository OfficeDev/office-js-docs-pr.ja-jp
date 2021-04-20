---
title: Excel JavaScript API を使用した基本的なプログラミングの概念
description: Excel JavaScript API を使用して、Excel 用アドインをビルドします。
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: dde7dc66e0746fc4d9cf91ed3df824fab05c109d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292600"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a>Excel JavaScript API を使用した基本的なプログラミングの概念

この記事では、[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) を使用して Excel 2016 以降のアドインをビルドする方法について説明します。 ここでは API の使用の基本となる中心概念について説明し、広い範囲に対する読み取り、書き込み、一定範囲内すべてのセルの更新など、特定のタスクを実行するためのガイダンスを提供します。

> [!IMPORTANT]
> Excel API の非同期性とブックでの動作方法については、「[Using the application-specific API model (アプリケーション固有の API モデルの使用)](../develop/application-specific-api-model.md)」を参照してください。  

## <a name="officejs-apis-for-excel"></a>Excel 用の Office.js API

Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して、Excel のオブジェクトを操作します。

* **Excel JavaScript API**:Office 2016 で導入された [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。

* **共通 API**: Office 2013 で導入された[共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。

Excel 2016 以降を対象にしたアドインでは、機能の大部分を Excel JavaScript API を使用して開発する可能性がありますが、共通 API のオブジェクトも使用します。 例:

* [Context](/javascript/api/office/office.context): `Context`Context`contentLanguage` オブジェクトは、アドインのランタイム環境を表し、API の主要なオブジェクトへのアクセスを提供します。 これは `officeTheme` や `host` などのブック構成の詳細で構成され、`platform` や `requirements.isSetSupported()` などのアドインのランタイム環境に関する情報も提供します。 さらに、 メソッドも提供されます。これを使用すると、指定した要件セットが、アドインが実行されている Excel アプリケーションでサポートされているかどうかを確認できます。
* [Document](/javascript/api/office/office.document): `Document` オブジェクトは `getFileAsync()` メソッドを提供します。これを使用すると、アドインが実行されている Excel ファイルをダウンロードできます。

次の図は、Excel JavaScript API または共通 API を使用するタイミングを示しています。

![Excel JS API と共通 API の違いを示す画像](../images/excel-js-api-common-api.png)

## <a name="object-model"></a>オブジェクト モデル

Excel API について理解するには、ブックの構成要素が互いにどのように関連しているかを理解する必要があります。

* **ブック** には、1 つ以上の **ワークシート** が含まれます。
* **ワークシート** では、**Range** オブジェクトを介してセルにアクセスできます。
* **Range** は、連続したセルのグループを表します。
* **Range** は、**表**、**グラフ**、**図形**、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。
* **ワークシート** には、個々のシートに存在するデータ オブジェクトのコレクションが含まれます。
* **ブック** には、**ブック** 全体のデータ オブジェクト (**表** など) の一部のコレクションが含まれます。

### <a name="ranges"></a>範囲

範囲とは、ブック内の連続したセルのグループのことです。 アドインでは、範囲を定義するのに通常 A1 形式の表記が使用されます (例: **B3** は、列 **B**、行 **3** の単一のセルで、**C2:F4** は、列 **C** から **F**、行 **2** から **4** までのセル)。

範囲には `values`、`formulas`、`format` の 3 つの主要なプロパティがあります。 これらのプロパティで、セルの値、評価する数式、およびセルの視覚的な書式設定を取得または設定します。

#### <a name="range-sample"></a>サンプル範囲

次のサンプルで、売上記録の作成方法を示します。 この関数は、`Range` オブジェクトを使用して、値、数式、書式を設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    var headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    var headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    var productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    var dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    var totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    var totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    return context.sync();
});
```

このサンプルは、現在のワークシートに次のデータを作成します。

![値の行、数式の列、書式設定されたヘッダーを示す売上記録。](../images/excel-overview-range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>グラフ、表、およびその他のデータ オブジェクト

Excel JavaScript API を使用することにより、Excel 内でデータ構造やビジュアル化を作成および操作できます。 表とグラフの 2 つのオブジェクトが頻繁に使用されますが、API はピボットテーブル、図形、画像などもサポートしています。

#### <a name="creating-a-table"></a>表の作成

データが入力された範囲を使用することにより、表を作成します。 書式設定とテーブル コントロール (フィルターなど) が自動的に範囲に適用されます。

次のサンプルでは、前のサンプルの範囲を使用して表を作成します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

前のデータを含むワークシート上でこのサンプル コードを使用すると、次のテーブルが作成されます。

![前の売上記録から作成された表。](../images/excel-overview-table-sample.png)

#### <a name="creating-a-chart"></a>グラフの作成

グラフを作成すると、範囲内のデータを視覚化できます。 この API は、さまざまな種類のグラフをサポートしています。いずれのグラフも、必要に応じてカスタマイズできます。

次のサンプルでは 3 つの品目の簡単な縦棒グラフが作成され、ワークシートの上端から 100 ピクセル下に配置されます。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

前の表を含むワークシート上でこのサンプルを実行すると、次のグラフが作成されます。

![前の売上記録の 3 つの品目の数量が表示されている縦棒グラフ。](../images/excel-overview-chart-sample.png)

## <a name="run-options"></a>実行オプション

`Excel.run` には、[RunOptions](/javascript/api/excel/excel.runoptions) オブジェクトを使用するオーバーロードがあります。 これには、関数の実行時にプラットフォームの動作に影響を与えるプロパティのセットが含まれています。 次のプロパティが現在サポートされています。

* `delayForCellEdit`: ユーザーがセル編集モードを終了するまでバッチ要求を延期するかどうかを指定します。 **true** の場合、バッチ要求は延期され、ユーザーがセル編集モードを終了した時点で実行されます。 **false** の場合、バッチ要求は、ユーザーがセル編集モードにある場合、自動的に失敗します (ユーザーにエラーが表示されます)。 `delayForCellEdit` プロパティが指定されていない場合の既定の動作は、このプロパティが **false** の場合と同じ動作となります。

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="null-or-blank-property-values"></a>null または空白のプロパティ値

`null` と空の文字列は、Excel JavaScript API では特別な意味を持ちます。 これらは、空のセル、書式設定なし、既定値を表すために使用されます。 このセクションでは、プロパティの取得や設定を行うときに `null` や空の文字列を使用する方法について詳しく説明します。

### <a name="null-input-in-2-d-array"></a>2 次元配列での null の入力

Excel では、範囲は 2 次元配列で表され、最初のディメンションは行、2 番目のディメンションは列を示します。 範囲内の特定のセルだけに値、数値書式、または数式を設定するには、2 次元配列内のそのセルに値、数値書式、または数式を指定し、2 次元配列内のその他のすべてのセルに `null` を指定します。

たとえば、範囲内の 1 つのセルの数値書式を更新し、範囲内の他のセルすべての既存の数値書式を保持する場合、更新するセルに新しい数値書式を指定し、他のセルすべてに `null` を指定します。 次のコード スニペットでは、範囲内の 4 番目のセルに新しい数値書式を設定し、その前の 3 つのセルについては数値書式を変更せずに保持します。

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a>プロパティに対する null の入力

`null` は単一プロパティに有効な入力ではありません。たとえば、次のコード スニペットは、範囲の `values` プロパティを `null` に設定できないため無効です。

```js
range.values = null;
```

同様に、次のコード スニペットは、`null` が `color` プロパティで有効な値ではないため無効です。

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a>応答内の null プロパティ値

指定の範囲に複数の値がある場合、`size` および `color` などの書式設定プロパティでは、応答に `null` 値が含まれます。 たとえば、範囲を取得してその `format.font.color` プロパティを読み込む場合:

* 範囲内のすべてのセルのフォントの色が同じ場合、`range.format.font.color` がその色を指定します。
* 範囲内に複数のフォントの色がある場合、`range.format.font.color` は `null` です。

### <a name="blank-input-for-a-property"></a>プロパティに対する空白の入力

プロパティに空白の値 (`''` の間にスペースのない 2 つの引用符) を指定すると、プロパティをクリアまたはリセットする指示として解釈されます。例:

* 範囲の `values` プロパティに空白の値を指定すると、範囲のコンテンツはクリアされます。
* `numberFormat` プロパティに空白の値を指定すると、数値書式は `General` にリセットされます。
* `formula` プロパティと `formulaLocale` プロパティに空白の値を指定すると、数式の値はクリアされます。

### <a name="blank-property-values-in-the-response"></a>応答内の空白のプロパティ値

読み取り操作では、応答内の空白のプロパティ値 (`''` の間にスペースのない、2 つの引用符) は、セルにデータまたは値がないことを示します。 次の 1 番目の例では、範囲内の最初と最後のセルにデータがありません。 2 番目の例では、範囲内の最初の 2 つのセルに数式がありません。

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## <a name="requirement-sets"></a>要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインはランタイム チェックを実行できます。または、マニフェストで指定されている要件セットを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを確認できます。 サポートされている各プラットフォームで使用できる特定の要件セットを確認するには、「[Excel JavaScript API の要件セット](../reference/requirement-sets/excel-api-requirement-sets.md)」を参照してください。

### <a name="checking-for-requirement-set-support-at-runtime"></a>実行時に要件セットのサポートを確認する

次のコード サンプルは、アドインが実行されている Office アプリケーションが指定された API の要件セットをサポートしているかどうかを確認する方法を示しています。

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>マニフェストで要件セットのサポートを定義する

アドインのマニフェストで [Requirements 要素](../reference/manifest/requirements.md) を使用して、アドインをアクティブにするために必要な最小要件セットや API メソッド (またはその両方) を指定できます。 Office アプリケーションやプラットフォームが、マニフェストの `Requirements` 要素で指定した要件セットまたは API メソッドをサポートしない場合、アドインはそのアプリケーションまたはプラットフォームでは実行されず、**[個人用アドイン]** に表示されるアドインの一覧にも表示されません。

次のコード サンプルは、アドインが ExcelApi 要件セットのバージョン 1.3 以上をサポートする Office クライアント アプリケーションのすべてで読み込まれる必要があることを指定する、アドインのマニフェストの `Requirements` 要素を示しています。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> Excel on the web、Windows、iPad などの Office アプリケーションのプラットフォームすべてでアドインを使用できるようにするには、マニフェストで要件セットのサポートを定義するのではなく、実行時に要件のサポートを確認することをお勧めします。

### <a name="requirement-sets-for-the-officejs-common-api"></a>Office.js 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」をご覧ください。

## <a name="handle-errors"></a>エラーを処理する

API エラーが発生すると、API はコードとメッセージを含む `error` オブジェクトを返します。 エラーの処理に関する詳細と、API エラーの一覧については、「[エラー処理](excel-add-ins-error-handling.md)」を参照してください。

## <a name="see-also"></a>関連項目

* [最初の Excel アドインをビルドする](../quickstarts/excel-quickstart-jquery.md)
* [Excel アドインのコード サンプル](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel の JavaScript API を使用した、パフォーマンスの最適化](../excel/performance.md)
* [Excel JavaScript API リファレンス](../reference/overview/excel-add-ins-reference-overview.md)
* [一般的なコーディングの問題と、予期しないプラットフォームの動作](../develop/common-coding-issues.md)
