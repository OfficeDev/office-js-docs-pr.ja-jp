---
title: Excel JavaScript API を使用して重複を削除する
description: Excel JavaScript API を使用して重複を削除する方法について説明します。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9ece7c9f35b341dbb8d0d90e8ca4bda5215580ed
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889143"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a>Excel JavaScript API を使用して重複を削除する

この記事では、Excel JavaScript API を使用して、範囲内の重複するエントリを削除するコード サンプルを提供します。 オブジェクトがサポートする `Range` プロパティとメソッドの完全な一覧については、 [Excel.Range クラス](/javascript/api/excel/excel.range)を参照してください。

## <a name="remove-rows-with-duplicate-entries"></a>重複するエントリを含む行を削除する

[Range.removeDuplicates](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1)) メソッドは、指定した列のエントリが重複する行を削除します。 このメソッドは、最も低い値のインデックスから範囲の最も高い値のインデックス (上から下) までの範囲内の各行を通過します。 任意の行で、指定された 1 つまたは複数の列が範囲より前に表示されている場合、その行は削除されます。 範囲にある削除された行の下の行が上に移動します。 `removeDuplicates` は、範囲外にあるセルの位置には影響しません。

`removeDuplicates` は、どの重複をチェックするかを示す列インデックスを表す `number[]` を受け取ります。 この配列は、0 から始まり、ワークシートではなく範囲を基準にしています。 このメソッドは、最初の行がヘッダーであるかどうかを示すブール型パラメーターも取り込みます。 重複を検討するときに、先頭行が無視される場合 `true`。 このメソッドは `removeDuplicates` 、 `RemoveDuplicatesResult` 削除された行数と残りの一意の行数を指定するオブジェクトを返します。

範囲の `removeDuplicates` メソッドを使用する場合は、次の点に注意してください。

- `removeDuplicates` は、関数の結果ではなくセルの値を考慮します。 2 つの異なる関数が同じ結果として評価される場合、セルの値は重複と見なしません。
- 空のセルは、`removeDuplicates` に無視されることはありません。 空のセルの値は、その他の値と同様に扱われます。 つまり、範囲に含まれる空の行は `RemoveDuplicatesResult` に含まれることになります。

次のコード サンプルは、最初の列の値が重複するエントリの削除を示しています。

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

![範囲の remove duplicates メソッドが実行される前の Excel のデータ。](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a>重複するエントリが削除された後のデータ

![範囲の remove duplicates メソッドが実行された後の Excel のデータ。](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してセルを操作する](excel-add-ins-cells.md)
- [Excel JavaScript API を使用して範囲を切り取り、コピー、貼り付ける](excel-add-ins-ranges-cut-copy-paste.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
