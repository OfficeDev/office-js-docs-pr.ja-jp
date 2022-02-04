---
title: JavaScript API を使用して重複Excel削除する
description: JavaScript API を使用して重複Excelする方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="remove-duplicates-using-the-excel-javascript-api"></a>JavaScript API を使用して重複Excel削除する

この記事では、JavaScript API を使用して範囲内の重複エントリを削除するExcel示します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、「Excel[。Range クラス](/javascript/api/excel/excel.range)。

## <a name="remove-rows-with-duplicate-entries"></a>重複するエントリがある行を削除する

[Range.removeDuplicates メソッド](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1))は、指定した列に重複するエントリがある行を削除します。 メソッドは、最も低い値のインデックスから範囲の最も高い値のインデックス (上から下) の範囲の各行を通過します。 任意の行で、指定された 1 つまたは複数の列が範囲より前に表示されている場合、その行は削除されます。 範囲にある削除された行の下の行が上に移動します。 `removeDuplicates` は、範囲外にあるセルの位置には影響しません。

`removeDuplicates` は、どの重複をチェックするかを示す列インデックスを表す `number[]` を受け取ります。 この配列は、0 から始まり、ワークシートではなく範囲を基準にしています。 このメソッドは、最初の行がヘッダーであるかどうかを指定するブール型パラメーターも取ります。 **true** の場合、重複について考慮するとき最初の行は無視されます。 この `removeDuplicates` メソッドは、削除された `RemoveDuplicatesResult` 行の数と残りの一意の行数を指定するオブジェクトを返します。

範囲のメソッドを使用する場合は `removeDuplicates` 、次の念に従います。

- `removeDuplicates` は、関数の結果ではなくセルの値を考慮します。 2 つの異なる関数が同じ結果として評価される場合、セルの値は重複と見なしません。
- 空のセルは、`removeDuplicates` に無視されることはありません。 空のセルの値は、その他の値と同様に扱われます。 つまり、範囲に含まれる空の行は `RemoveDuplicatesResult` に含まれることになります。

次のコード サンプルは、最初の列に重複する値を持つエントリの削除を示しています。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

### <a name="data-before-duplicate-entries-are-removed"></a>重複するエントリが削除される前のデータ

![範囲のExcelの重複メソッドが実行される前のデータ。](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a>重複するエントリが削除された後のデータ

![範囲のExcel重複するメソッドが実行された後のデータ。](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲を切り取り、コピー Excel貼り付ける](excel-add-ins-ranges-cut-copy-paste.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
