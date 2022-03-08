---
title: JavaScript API を使用して重複Excel削除する
description: JavaScript API を使用して重複Excelする方法について説明します。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 80e1227e06f177d0e37cc2750a7830c727a59436
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340576"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a>JavaScript API を使用して重複Excel削除する

この記事では、JavaScript API を使用して範囲内の重複エントリを削除するコード サンプルExcel示します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、「Excel[。Range クラス](/javascript/api/excel/excel.range)。

## <a name="remove-rows-with-duplicate-entries"></a>重複するエントリがある行を削除する

[Range.removeDuplicates メソッド](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1))は、指定した列に重複するエントリがある行を削除します。 メソッドは、最も低い値のインデックスから範囲の最も高い値のインデックス (上から下) の範囲の各行を通過します。 任意の行で、指定された 1 つまたは複数の列が範囲より前に表示されている場合、その行は削除されます。 範囲にある削除された行の下の行が上に移動します。 `removeDuplicates` は、範囲外にあるセルの位置には影響しません。

`removeDuplicates` は、どの重複をチェックするかを示す列インデックスを表す `number[]` を受け取ります。 この配列は、0 から始まり、ワークシートではなく範囲を基準にしています。 このメソッドは、最初の行がヘッダーであるかどうかを指定するブール型パラメーターも取ります。 **true** の場合、重複について考慮するとき最初の行は無視されます。 この `removeDuplicates` メソッドは、削除された `RemoveDuplicatesResult` 行の数と残りの一意の行数を指定するオブジェクトを返します。

範囲のメソッドを使用する場合は `removeDuplicates` 、次の念に従います。

- `removeDuplicates` は、関数の結果ではなくセルの値を考慮します。 2 つの異なる関数が同じ結果として評価される場合、セルの値は重複と見なしません。
- 空のセルは、`removeDuplicates` に無視されることはありません。 空のセルの値は、その他の値と同様に扱われます。 つまり、範囲に含まれる空の行は `RemoveDuplicatesResult` に含まれることになります。

次のコード サンプルは、最初の列に重複する値を持つエントリの削除を示しています。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B2:D11");

    let deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    await context.sync();

    console.log(deleteResult.removed + " entries with duplicate names removed.");
    console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
});
```

### <a name="data-before-duplicate-entries-are-removed"></a>重複するエントリが削除される前のデータ

![範囲の remove duplicates メソッドExcel実行する前に、データ内のデータを指定します。](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a>重複するエントリが削除された後のデータ

![範囲のExcel重複するメソッドが実行された後のデータ。](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲を切り取り、コピー Excel貼り付ける](excel-add-ins-ranges-cut-copy-paste.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
