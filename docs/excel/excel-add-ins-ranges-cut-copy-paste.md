---
title: JavaScript API を使用して範囲を切り取り、コピー Excel貼り付ける
description: JavaScript API を使用して範囲を切り取り、コピー、貼りExcel説明します。
ms.date: 02/16/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 3d55e4d868a15c35ab9c68c799865560547e8188
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745100"
---
# <a name="cut-copy-and-paste-ranges-using-the-excel-javascript-api"></a>JavaScript API を使用して範囲を切り取り、コピー Excel貼り付ける

この記事では、JavaScript API を使用して範囲を切り取り、コピー、貼り付けるExcel説明します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、「Excel[。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="copy-and-paste"></a>Copy and paste

[Range.copyFrom](/javascript/api/excel/excel.range#excel-excel-range-copyfrom-member(1)) メソッドは、ユーザー UI の **コピー** と **貼** り付Excelします。 宛先は、呼 `Range` び出される `copyFrom` オブジェクトです。 コピーされるソースは、範囲または範囲を表す文字列のアドレスとして渡されます。

次のコード サンプルでは、**A1:E1** のデータを **G1** で始まる範囲にコピーします (この貼り付けは **G1:K1** で終わります)。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    // Copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1").
    sheet.getRange("G1").copyFrom("A1:E1");
    await context.sync();
});
```

`Range.copyFrom` には、省略可能なパラメーターが 3 つあります。

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` では、ソースからコピー先にコピーされるデータを指定します。

- `Excel.RangeCopyType.formulas` ソース セル内の数式を転送し、それらの数式の範囲の相対位置を保持します。 任意の数式以外のエントリはそのままコピーされます。
- `Excel.RangeCopyType.values` では、データ値と、数式の場合は数式の結果をコピーします。
- `Excel.RangeCopyType.formats` では、フォント、色、およびその他の書式設定を含む、範囲の書式設定をコピーしますが、値はコピーしません。
- `Excel.RangeCopyType.all` (既定のオプション) は、データと書式設定の両方をコピーし、セルの数式が見つかった場合は保持します。

`skipBlanks` では、空白セルをコピー先にコピーするかどうかを設定します。 true の場合、`copyFrom` ではソースの範囲にある空白セルはスキップされます。
スキップされたセルでは、コピー先の範囲内の対応するセルにある既存のデータを上書きすることはありません。 既定値は false です。

`transpose` では、ソースの場所へのデータの行と列の入れ替えを行うかどうかを決定します。
行と列を入れ替える範囲は対角線で反転されるため、行 **1**、**2**、**3** が列 **A**、**B**、**C** になります。

次のコード サンプルと画像は、この動作をシンプルなシナリオで示しています。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    // Copy a range, omitting the blank cells so existing data is not overwritten in those cells.
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // Copy a range, including the blank cells which will overwrite existing data in the target cells.
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    await context.sync();
});
```

### <a name="data-before-range-is-copied-and-pasted"></a>範囲がコピーおよび貼り付けされる前のデータ

![範囲のコピー Excel実行する前のデータ。](../images/excel-range-copyfrom-skipblanks-before.png)

### <a name="data-after-range-is-copied-and-pasted"></a>範囲がコピーおよび貼り付けされた後のデータ

![範囲のExcelが実行された後のデータ。](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="cut-and-paste-move-cells"></a>セルの切り取りと貼り付け (移動)

[Range.moveTo メソッドは](/javascript/api/excel/excel.range#excel-excel-range-moveto-member(1))、ブック内の新しい場所にセルを移動します。 このセルの移動動作は、セルを移動するときに、範囲 [](https://support.microsoft.com/office/803d65eb-6a3e-4534-8c6f-ff12d1c4139e)の境界線をドラッグするか、切り取りおよび貼り付けアクションを実行する場合 **と****同じように動作** します。 範囲の書式設定と値の両方が、パラメーターとして指定された場所に移動 `destinationRange` されます。

次のコード サンプルでは、メソッドを使用して範囲を移動 `Range.moveTo` します。 移動先の範囲がソースより小さい場合は、ソース コンテンツを含む範囲に拡張されます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    await context.sync();
});
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して重複Excel削除する](excel-add-ins-ranges-remove-duplicates.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
