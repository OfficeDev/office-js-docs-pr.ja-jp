---
title: JavaScript API を使用して選択した範囲を設定Excel取得する
description: JavaScript API を使用して、Excel JavaScript API を使用して選択した範囲を設定および取得するExcel説明します。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9517c072fae92b1b541a52b1805834c2bb429dd3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745364"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a>JavaScript API を使用して選択した範囲を設定Excel取得する

この記事では、JavaScript API を使用して選択した範囲を設定および取得するExcel説明します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、「Excel[。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a>選択範囲を設定する

次のコード サンプルは、作業中のワークシートの範囲 **B2:E6** を選択します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:E6");

    range.select();

    await context.sync();
});
```

### <a name="selected-range-b2e6"></a>選択範囲 B2:E6

![[選択した範囲] Excel。](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a>選択範囲を取得する

次のコード サンプルでは、選択した範囲を取得し、そのプロパティを読 `address` み込み、コンソールにメッセージを書き込みます。

```js
await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    range.load("address");

    await context.sync();
    
    console.log(`The address of the selected range is "${range.address}"`);
});
```

## <a name="select-the-edge-of-a-used-range"></a>使用範囲の端を選択する

[Range.getRangeEdge](/javascript/api/excel/excel.range#excel-excel-range-getrangeedge-member(1)) メソッドと [Range.getExtendedRange](/javascript/api/excel/excel.range#excel-excel-range-getextendedrange-member(1)) メソッドを使用すると、アドインはキーボード選択ショートカットの動作をレプリケートし、現在選択されている範囲に基づいて使用範囲のエッジを選択できます。 使用範囲の詳細については、「使用範囲の [取得」を参照してください](excel-add-ins-ranges-get.md#get-used-range)。

次のスクリーンショットでは、使用範囲は、各セルの **値が C5:F12 の表です**。 この表の外側の空のセルは、使用範囲の外側です。

![C5:F12 のデータが含Excel。](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a>現在使用されている範囲の端にあるセルを選択する

次のコード サンプルは、 `Range.getRangeEdge` メソッドを使用して、現在使用されている範囲の最も遠い端にあるセルを上方向に選択する方法を示しています。 このアクションは、範囲が選択されている間に Ctrl + 上矢印キーのキーボード ショートカットを使用した結果と一致します。

```js
await Excel.run(async (context) => {
    // Get the selected range.
    let range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    let direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    let activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Up arrow key keyboard shortcut while a range is selected.
    let rangeEdge = range.getRangeEdge(
      direction,
      activeCell
    );
    rangeEdge.select();

    await context.sync();
});
```

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a>使用範囲の端にあるセルを選択する前に

次のスクリーンショットは、使用範囲と、使用範囲内で選択した範囲を示しています。 使用範囲は、 **C5:F12** のデータを含むテーブルです。 この表の中で、 **範囲 D8:E9 が** 選択されています。 この選択は、 *メソッドを実行* する前の前の状態 `Range.getRangeEdge` です。

![C5:F12 のデータが含Excel。 範囲 D8:E9 が選択されています。](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a>使用範囲の端にあるセルを選択した後

次のスクリーンショットは、前のスクリーンショットと同じ表を示し、 **C5:F12** の範囲のデータを示しています。 この表の中で、 **範囲 D5 が** 選択されています。 この選択は *、メソッド* を実行した後 `Range.getRangeEdge` の状態の後で、使用範囲の端にあるセルを上方向に選択します。

![C5:F12 のデータが含Excel。 範囲 D5 が選択されています。](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a>現在の範囲から使用範囲の最も遠い端までのすべてのセルを選択する

次のコード サンプル `Range.getExtendedRange` は、メソッドを使用して、現在選択されている範囲から使用範囲の最も遠い端まで、下方向のすべてのセルを選択する方法を示しています。 このアクションは、範囲が選択されている間に Ctrl + Shift +下矢印キーのキーボード ショートカットを使用した結果と一致します。

```js
await Excel.run(async (context) => {
    // Get the selected range.
    let range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    let direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    let activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.
    let extendedRange = range.getExtendedRange(
      direction,
      activeCell
    );
    extendedRange.select();

    await context.sync();
});
```

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a>現在の範囲から使用範囲の端までのすべてのセルを選択する前に

次のスクリーンショットは、使用範囲と、使用範囲内で選択した範囲を示しています。 使用範囲は、 **C5:F12** のデータを含むテーブルです。 この表の中で、 **範囲 D8:E9 が** 選択されています。 この選択は、 *メソッドを実行* する前の前の状態 `Range.getExtendedRange` です。

![C5:F12 のデータが含Excel。 範囲 D8:E9 が選択されています。](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a>現在の範囲から使用範囲の端までのすべてのセルを選択した後

次のスクリーンショットは、前のスクリーンショットと同じ表を示し、 **C5:F12** の範囲のデータを示しています。 この表の中で、 **範囲 D8:E12 が** 選択されています。 この選択は *、メソッド*`Range.getExtendedRange`を実行した後の状態の後で、現在の範囲から下方向の使用範囲の端までのすべてのセルを選択します。

![C5:F12 のデータが含Excel。 範囲 D8:E12 が選択されています。](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲の値、テキスト、または数式を設定Excel取得する](excel-add-ins-ranges-set-get-values.md)
- [JavaScript API を使用して範囲Excel設定する](excel-add-ins-ranges-set-format.md)
