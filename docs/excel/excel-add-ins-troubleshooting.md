---
title: アドインExcelトラブルシューティング
description: アドインの開発エラーをトラブルシューティングするExcel説明します。
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: b90d8cfdb4696445655122a2fa7eb74d1c87fa2f
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671465"
---
# <a name="troubleshooting-excel-add-ins"></a>アドインExcelトラブルシューティング

この記事では、ユーザーに固有のトラブルシューティングの問題についてExcel。 ページの下部にあるフィードバック ツールを使用して、記事に追加できるその他の問題を提案してください。

## <a name="api-limitations-when-the-active-workbook-switches"></a>アクティブなブックが切り替え時の API の制限

一度に 1 Excelブックを操作することを目的としたアドインです。 アドインを実行しているブックとは別のブックがフォーカスを獲得すると、エラーが発生する可能性があります。 これは、特定のメソッドがフォーカスの変更時に呼び出される過程にある場合にのみ発生します。

次の API は、このブックスイッチの影響を受ける。

|Excel JavaScript API | スローされるエラー |
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
> これは、ユーザーまたは Mac で開Excel複数のブックにのみWindows適用されます。

## <a name="coauthoring"></a>共同編集

共同[編集環境でイベントと](co-authoring-in-excel-add-ins.md)Excelするパターンについては、「Coauthoring in Excelアドイン」を参照してください。 この記事では、 などの特定の API を使用する場合の潜在的なマージ競合について説明します [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add_index__values_) 。

## <a name="known-issues"></a>既知の問題

### <a name="binding-events-return-temporary-binding-obects"></a>バインド イベントは一時的 `Binding` な obects を返します

[BindingDataChangedEventArgs.binding と](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding)の両方が、イベントを発生したオブジェクトの ID を含む一時オブジェクト `Binding` `Binding` を返します。 この ID を使用して `BindingCollection.getItem(id)` 、イベントを発生 `Binding` したオブジェクトを取得します。

次のコード サンプルは、この一時的なバインド ID を使用して関連するオブジェクトを取得する方法を示 `Binding` しています。 サンプルでは、イベント リスナーがバインドに割り当てられます。 イベントがトリガーされると `getBindingId` 、リスナー `onDataChanged` はメソッドを呼び出します。 メソッド `getBindingId` は、一時オブジェクトの ID `Binding` を使用して、イベントを発生 `Binding` したオブジェクトを取得します。

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

[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight)プロパティは、このプロパティで正 `CellPropertiesFormat` Excel on the web。 UI の問題のため、このExcel on the webで高さを不正確に `useStandardHeight` 計算するプロパティ `true` を設定します。 たとえば、標準の高さ **14** は **14.25** に変更Excel on the web。

すべてのプラットフォームで [、useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) プロパティと [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) プロパティは、 に設定される `CellPropertiesFormat` のみです `true` 。 これらのプロパティを設定 `false` すると、効果はありません。 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>Range `getImage` メソッドは、サポートされていないExcel for Mac

Range [getImage](/javascript/api/excel/excel.range#getImage__)メソッドは現在、このメソッドではExcel for Mac。 現在 [の状態については、「OfficeDev/office-js issue #235」](https://github.com/OfficeDev/office-js/issues/235) を参照してください。

### <a name="range-return-character-limit"></a>範囲の戻り値の文字制限

[Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_)メソッドと[Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_)メソッドのアドレス文字列の制限は 8192 文字です。 この制限を超えると、アドレス文字列は 8192 文字に切り詰まれます。

## <a name="see-also"></a>関連項目

- [アドインを使用したOfficeのトラブルシューティング](../testing/troubleshoot-development-errors.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)
