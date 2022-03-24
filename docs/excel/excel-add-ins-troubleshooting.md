---
title: アドインExcelトラブルシューティング
description: アドインの開発エラーをトラブルシューティングするExcel説明します。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: aabbb3d8b62101eacb2ac51684a3d1f6c16e84a4
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745507"
---
# <a name="troubleshooting-excel-add-ins"></a>アドインExcelトラブルシューティング

この記事では、ユーザー固有のトラブルシューティングの問題についてExcel。 ページの下部にあるフィードバック ツールを使用して、記事に追加できるその他の問題を提案してください。

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

共同編集環境でイベントとExcelするパターンについては、「[Coauthoring](co-authoring-in-excel-add-ins.md) in Excelアドイン」を参照してください。 この記事では、 などの特定の API を使用する場合の潜在的なマージ競合について説明します [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1))。

## <a name="known-issues"></a>既知の問題

### <a name="binding-events-return-temporary-binding-obects"></a>バインド イベントは一時的 `Binding` な obects を返します

[BindingDataChangedEventArgs.binding と](/javascript/api/excel/excel.bindingdatachangedeventargs#excel-excel-bindingdatachangedeventargs-binding-member) [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#excel-excel-bindingselectionchangedeventargs-binding-member) `Binding` の両方が、イベントを発生したオブジェクトの ID `Binding` を含む一時オブジェクトを返します。 この ID を使用して、 `BindingCollection.getItem(id)` イベントを発生 `Binding` したオブジェクトを取得します。

次のコード サンプルは、この一時的なバインド ID を使用して関連するオブジェクトを取得する方法を示 `Binding` しています。 サンプルでは、イベント リスナーがバインドに割り当てられます。 イベントがトリガーされると、 `getBindingId` リスナーはメソッド `onDataChanged` を呼び出します。 メソッド `getBindingId` は、一時オブジェクトの ID を使用 `Binding` して、イベントを発生した `Binding` オブジェクトを取得します。

```js
async function run() {
    await Excel.run(async (context) => {
        // Retrieve your binding.
        let binding = context.workbook.bindings.getItemAt(0);
    
        await context.sync();
    
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);
        await context.sync();
    });
}

async function getBindingId(eventArgs) {
    await Excel.run(async (context) => {
        // Get the temporary binding object and load its ID. 
        let tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        let originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>セルの形式 `useStandardHeight` と `useStandardWidth` 問題

[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member) プロパティは`CellPropertiesFormat`、このプロパティで正Excel on the web。 UI の問題のため、このExcel on the web `useStandardHeight` `true`で高さを不正確に計算するプロパティを設定します。 たとえば、標準の高さ 14 は **14.25** に変更Excel on the web。

すべてのプラットフォームで、 [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member) プロパティと [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardwidth-member) `CellPropertiesFormat` プロパティは、 に設定されるのみです `true`。 これらのプロパティを設定すると `false` 、効果はありません。

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>Range `getImage` メソッドは、サポートされていないExcel for Mac

Range [getImage](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1)) メソッドは現在、このメソッドでExcel for Mac。 現在 [の状態については、「OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) 」を参照してください。

### <a name="range-return-character-limit"></a>範囲の戻り値の文字制限

[Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) メソッドと [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getranges-member(1)) メソッドのアドレス文字列の制限は 8192 文字です。 この制限を超えると、アドレス文字列は 8192 文字に切り詰まれます。

## <a name="see-also"></a>関連項目

- [Office アドインでの開発エラーのトラブルシューティング](../testing/troubleshoot-development-errors.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)
