---
title: Excel アドインのトラブルシューティング
description: Excel アドインの開発エラーをトラブルシューティングする方法について説明します。
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 0efc8b4d25d9d748975146e187104972e4ad58a9
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270729"
---
# <a name="troubleshooting-excel-add-ins"></a>Excel アドインのトラブルシューティング

この記事では、Excel に固有の問題のトラブルシューティングについて説明します。 ページの下部にあるフィードバック ツールを使用して、記事に追加できるその他の問題を提案してください。

## <a name="api-limitations-when-the-active-workbook-switches"></a>アクティブなブックが切り替えらた場合の API の制限事項

Excel 用アドインは、一度に 1 つのブックで動作することを目的とします。 アドインを実行しているブックとは別のブックがフォーカスを得た場合、エラーが発生する可能性があります。 これは、フォーカスが変更された時点で特定のメソッドが呼び出されている場合にのみ発生します。

このブックの切り替えによって影響を受ける API は次のとおりです。

|Excel JavaScript API | スローされたエラー |
|--|--|
| `Chart.activate` | GeneralException |
| `Range.select` | GeneralException |
| `Table.clearFilters` | GeneralException |
| `Workbook.getActiveCell`  | InvalidSelection|
| `Workbook.getSelectedRange` | InvalidSelection|
| `Workbook.getSelectedRanges`  | InvalidSelection|
| `Worksheet.activate` | GeneralException |
| `Worksheet.delete`  | InvalidSelection|
| `Worksheet.gridlines` | GeneralException |
| `Worksheet.showHeadings` | GeneralException |
| `WorksheetCollection.add` | GeneralException |
| `WorksheetFreezePanes.freezeAt` | GeneralException |
| `WorksheetFreezePanes.freezeColumns` | GeneralException |
| `WorksheetFreezePanes.freezeRows` | GeneralException |
| `WorksheetFreezePanes.getLocationOrNullObject`| GeneralException |
| `WorksheetFreezePanes.unfreeze` | GeneralException |

> [!NOTE]
> これは、Windows または Mac で開いている複数の Excel ブックにのみ適用されます。

## <a name="coauthoring"></a>共同編集

共同 [編集環境のイベントで](co-authoring-in-excel-add-ins.md) 使用するパターンについては、Excel アドインの共同編集を参照してください。 また、次のような特定の API を使用する場合のマージ競合の可能性について説明します [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) 。

## <a name="known-issues"></a>既知の問題

### <a name="binding-events-return-temporary-binding-obects"></a>バインド イベントが一時的 `Binding` なオベクトを返す

[BindingDataChangedEventArgs.binding と](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding)はどちらも、イベントを発生したオブジェクトの ID を含む一時オブジェクト `Binding` `Binding` を返します。 この ID を使用して `BindingCollection.getItem(id)` 、イベントを発生 `Binding` したオブジェクトを取得します。

次のコード サンプルは、この一時バインド ID を使用して関連オブジェクトを取得する方法を示 `Binding` しています。 サンプルでは、イベント リスナーがバインドに割り当てられます。 リスナーは、イベントが `getBindingId` トリガーされるとメソッド `onDataChanged` を呼び出します。 メソッド `getBindingId` は、一時オブジェクトの ID `Binding` を使用して、イベントを発生 `Binding` したオブジェクトを取得します。

```js
Excel.run(function (context) {
    // Retrieve your binding.
    var binding = context.workbook.bindings.getItemAt(0);

    return context.sync().then(function () {
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);

        return context.sync();
    });
});

function getBindingId(eventArgs) {
    return Excel.run(function (context) {
        // Get the temporary binding object and load its ID. 
        var tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        var originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>セルの `useStandardHeight` 形式 `useStandardWidth` と問題

Web 上の Excel [では、useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) プロパティ `CellPropertiesFormat` が正しく動作しません。 Excel on the web UI の問題により、このプラットフォームでは高さが不正確に計算されるプロパティ `useStandardHeight` `true` を設定します。 たとえば、Excel on the web では標準の高さ **14** **が 14.25** に変更されます。

すべてのプラットフォームで [、useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) プロパティと [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) プロパティは、設定のみを `CellPropertiesFormat` 目的とします `true` 。 これらのプロパティを設定 `false` する効果はありません。 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>`getImage`Excel for Mac でサポートされていない Range メソッド

Range [getImage メソッド](/javascript/api/excel/excel.range#getImage__) は、Excel for Mac では現在サポートされていません。 現在 [の状態については、「OfficeDev/office-js issue #235」](https://github.com/OfficeDev/office-js/issues/235) を参照してください。

### <a name="range-return-character-limit"></a>範囲の戻り値の文字制限

[Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_)メソッドと[Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_)メソッドのアドレス文字列の制限は 8192 文字です。 この制限を超えると、アドレス文字列は 8192 文字に切り捨てられて表示されます。

## <a name="see-also"></a>関連項目

- [アドインを使用したOfficeエラーのトラブルシューティング](../testing/troubleshoot-development-errors.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)
