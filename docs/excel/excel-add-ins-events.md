---
title: Excel JavaScript API を使用してイベントを操作する
description: JavaScript オブジェクトのイベントExcel一覧です。 これには、イベント ハンドラーと関連付けられたパターンの使用に関する情報が含まれます。
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 596d8738b4c4953a937825e6c7294b2478ae59f7
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936763"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してイベントを操作する

この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。

## <a name="events-in-excel"></a>Excel のイベント

Excel ブックで特定の種類の変更が発生するたびに、イベント通知がトリガーされます。 Excel JavaScript API を使用すると、イベント ハンドラーを登録できます。このハンドラーによって、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できるようになります。 現時点でサポートされているイベントは次のとおりです。

| イベント | 説明 | サポートされているオブジェクト |
|:---------------|:-------------|:-----------|
| `onActivated` | オブジェクトがアクティブ化されたときに発生します。 | [**Chart**](/javascript/api/excel/excel.chart#onActivated)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onActivated)、[**Shape**](/javascript/api/excel/excel.shape#onActivated)、[**Worksheet**](/javascript/api/excel/excel.worksheet#onActivated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onActivated) |
| `onActivated` | ブックがアクティブ化されると発生します。 | [**Workbook**](/javascript/api/excel/excel.workbook#onActivated) |
| `onAdded` | オブジェクトがコレクションに追加されたときに発生します。 | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onAdded), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onAdded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onAdded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onAdded) |
| `onAutoSaveSettingChanged` | ブックで `autoSave` の設定が変更されると発生します。 | [**Workbook**](/javascript/api/excel/excel.workbook#onAutoSaveSettingChanged) |
| `onCalculated` | ワークシートの計算が完了したとき (あるいはコレクションのすべてのワークシートが完了したとき) に発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onCalculated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onCalculated) |
| `onChanged` | 個々のセルまたはコメントのデータが変更された場合に発生します。 | [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onChanged), [**Table**](/javascript/api/excel/excel.table#onChanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onChanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onChanged) |
| `onColumnSorted` | 1 つ以上の列を並べ替えたときに発生します。 これは、左から右に並べ替えを実行したときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onColumnSorted)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onColumnSorted) |
| `onDataChanged` | バインド内でデータまたは書式設定が変更されるときに発生します。 | [**Binding**](/javascript/api/excel/excel.binding#onDataChanged) |
| `onDeactivated` | オブジェクトが非アクティブ化されたときに発生します。 | [**Chart**](/javascript/api/excel/excel.chart#onDeactivated)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onDeactivated)、[**Shape**](/javascript/api/excel/excel.shape#onDeactivated)、[**Worksheet**](/javascript/api/excel/excel.worksheet#onDeactivated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onDeactivated) |
| `onDeleted` | オブジェクトがコレクションから削除されたときに発生します。 | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onDeleted), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onDeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onDeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onDeleted) |
| `onFormatChanged` | ワークシートで書式設定が変更されたときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onFormatChanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormatChanged) |
| `onFormulaChanged` | 数式が変更された場合に発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onFormulaChanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged) |
| `onRowSorted` | 1 つ以上の行を並べ替えたときに発生します。 これは、上から下に並べ替えを実行したときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onRowSorted)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onRowSorted) |
| `onSelectionChanged` | アクティブなセルまたは選択範囲が変更されたときに発生します。 | [**Binding**](/javascript/api/excel/excel.binding#onSelectionChanged), [**Table**](/javascript/api/excel/excel.table#onSelectionChanged), [**Workbook**](/javascript/api/excel/excel.workbook#onSelectionChanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onSelectionChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onSelectionChanged) |
| `onRowHiddenChanged` | 特定のワークシート上の行非表示状態が変更されたときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onRowHiddenChanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onRowHiddenChanged) |
| `onSettingsChanged` | ドキュメント内の設定が変更されるときに発生します。 | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#onSettingsChanged) |
| `onSingleClicked` | ワークシートで左クリック / タップされたアクションが発生したときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onSingleClicked)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onSingleClicked) |

### <a name="events-in-preview"></a>プレビューでのイベント

> [!NOTE]
> 次のイベントは現在、パブリック プレビューでのみ利用できます。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| イベント | 説明 | サポートされているオブジェクト |
|:---------------|:-------------|:-----------|
| `onFiltered` | フィルターがオブジェクトに適用されたときに発生します。 | [**Table**](/javascript/api/excel/excel.table#onFiltered)、[**TableCollection**](/javascript/api/excel/excel.tablecollection#onFiltered)、[**Worksheet**](/javascript/api/excel/excel.worksheet#onFiltered)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFiltered) |

### <a name="event-triggers"></a>イベント トリガー

Excel ブックのイベントは、次の事項でトリガーできます。

- ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作
- ブックを変更する Office アドイン (JavaScript) コード
- ブックを変更する VBA アドイン (マクロ) コード

Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。

### <a name="lifecycle-of-an-event-handler"></a>イベント ハンドラーのライフサイクル

アドインがイベント ハンドラーを登録すると、そのイベント ハンドラーが作成されます。 アドインがイベント ハンドラーを登録解除するか、アドインが更新、再読み込み、または閉じられると、イベント ハンドラーは破棄されます。 イベント ハンドラーは Excel ファイルの一部として保持されず、また Excel on the web のセッション間でも保持されません。

> [!CAUTION]
> イベントが登録されているオブジェクト (`onChanged` イベントが登録されているテーブルなど) が削除されると、イベント ハンドラーはトリガーされませんが、アドインまたは Excel セッションが更新または閉じるまではメモリで維持されます。

### <a name="events-and-coauthoring"></a>イベントと共同編集

[共同編集機能](co-authoring-in-excel-add-ins.md)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。共同編集でトリガーできるイベント (`onChanged` など) の場合、対応する **Event** オブジェクトには **source** プロパティが含まれるようになります。このプロパティは、イベントが現在のユーザーによってローカルにトリガーされた (`event.source = Local`) ものか、リモートの共同作成者によってトリガーされた (`event.source = Remote`) ものかを示します。

## <a name="register-an-event-handler"></a>イベント ハンドラーの登録

次のコード例では、ワークシートの `onChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録します。 このコードでは、そのワークシートでデータが変更されたときに、`handleChange` 関数を実行するように指定しています。

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a>イベントの処理

前の例で示したように、イベント ハンドラーの登録時には、特定のイベントが発生したときに実行する関数を指定します。 その関数は、目的のシナリオに必要なアクションを実行するように設計できます。 次のコード例は、イベントに関する情報を単にコンソールに出力するイベント ハンドラー関数を示しています。

```js
function handleChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Change type of event: " + event.changeType);
                console.log("Address of event: " + event.address);
                console.log("Source of event: " + event.source);
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a>イベント ハンドラーを削除する

次のコード例では、ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。 また、そのイベント ハンドラーを削除するために、後から呼び出すことができる `remove()` 関数も定義しています。 イベント ハンドラーを `RequestContext` 削除するには、イベント ハンドラーの作成に使用する必要があります。 

```js
var eventResult;

Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);

function handleSelectionChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Address of current selection: " + event.address);
            });
    }).catch(errorHandlerFunction);
}

function remove() {
    return Excel.run(eventResult.context, function (context) {
        eventResult.remove();

        return context.sync()
            .then(function() {
                eventResult = null;
                console.log("Event handler successfully removed.");
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="enable-and-disable-events"></a>イベントの有効化と無効化

イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。
たとえば、アプリがイベントを受信する必要がないことや、複数エンティティの一括編集を実行中にイベントを無視できるすることがあります。

イベントは[ランタイム](/javascript/api/excel/excel.runtime) レベルで有効または無効にできます。
`enableEvents` プロパティは、イベントが発生したかどうかと、イベント ハンドラーがアクティブになったかどうかを判別します。

次のコード サンプルは、イベントのオンとオフを切り替える方法を示しています。

```js
Excel.run(function (context) {
    context.runtime.load("enableEvents");
    return context.sync()
        .then(function () {
            var eventBoolean = !context.runtime.enableEvents;
            context.runtime.enableEvents = eventBoolean;
            if (eventBoolean) {
                console.log("Events are currently on.");
            } else {
                console.log("Events are currently off.");
            }
        }).then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
