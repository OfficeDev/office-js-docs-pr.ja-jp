---
title: Excel JavaScript API を使用してイベントを操作する
description: Excel JavaScript オブジェクトのイベントのリスト。 これには、イベントハンドラーと関連付けられたパターンの使用に関する情報が含まれます。
ms.date: 08/18/2020
localization_priority: Normal
ms.openlocfilehash: 9c1610dc06af56ed436f1832baab395cbe9de971
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408461"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してイベントを操作する

この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。

## <a name="events-in-excel"></a>Excel のイベント

Excel ブックで特定の種類の変更が発生するたびに、イベント通知がトリガーされます。 Excel JavaScript API を使用すると、イベント ハンドラーを登録できます。このハンドラーによって、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できるようになります。 現時点でサポートされているイベントは次のとおりです。

| イベント | 説明 | サポートされているオブジェクト |
|:---------------|:-------------|:-----------|
| `onActivated` | オブジェクトがアクティブ化されたときに発生します。 | [**Chart**](/javascript/api/excel/excel.chart#onactivated)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated)、[**Shape**](/javascript/api/excel/excel.shape#onactivated)、[**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated) |
| `onAdded` | オブジェクトがコレクションに追加されたときに発生します。 | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded)、[**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded) |
| `onAutoSaveSettingChanged` | ブックで `autoSave` の設定が変更されると発生します。 | [**Workbook**](/javascript/api/excel/excel.workbook#onautosavesettingchanged) |
| `onCalculated` | ワークシートの計算が完了したとき (あるいはコレクションのすべてのワークシートが完了したとき) に発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated) |
| `onChanged` | セル内のデータが変更されたときに発生します。 | [**Table**](/javascript/api/excel/excel.table#onchanged)、[**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged)、[**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged) |
| `onColumnSorted` | 1 つ以上の列を並べ替えたときに発生します。 これは、左から右に並べ替えを実行したときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted) |
| `onDataChanged` | バインド内でデータまたは書式設定が変更されるときに発生します。 | [**Binding**](/javascript/api/excel/excel.binding#ondatachanged) |
| `onDeactivated` | オブジェクトが非アクティブ化されたときに発生します。 | [**Chart**](/javascript/api/excel/excel.chart#ondeactivated)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated)、[**Shape**](/javascript/api/excel/excel.shape#ondeactivated)、[**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated) |
| `onDeleted` | オブジェクトがコレクションから削除されたときに発生します。 | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted)、[**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted) |
| `onFormatChanged` | ワークシートで書式設定が変更されたときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged) |
| `onRowSorted` | 1 つ以上の行を並べ替えたときに発生します。 これは、上から下に並べ替えを実行したときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted) |
| `onSelectionChanged` | アクティブなセルまたは選択範囲が変更されたときに発生します。 | [**Binding**](/javascript/api/excel/excel.binding#onselectionchanged)、 [**Table**](/javascript/api/excel/excel.table#onselectionchanged)、 [**Workbook**](/javascript/api/excel/excel.workbook#onselectionchanged)、 [**Worksheet、Worksheet**](/javascript/api/excel/excel.worksheet#onselectionchanged)[**コレクション**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged) |
| `onRowHiddenChanged` | 特定のワークシート上の行非表示状態が変更されたときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged) |
| `onSettingsChanged` | ドキュメント内の設定が変更されるときに発生します。 | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#onsettingschanged) |
| `onSingleClicked` | ワークシートで左クリック / タップされたアクションが発生したときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked) |

### <a name="events-in-preview"></a>プレビューでのイベント

> [!NOTE]
> 次のイベントは現在、パブリック プレビューでのみ利用できます。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| イベント | 説明 | サポートされているオブジェクト |
|:---------------|:-------------|:-----------|
| `onFiltered` | フィルターがオブジェクトに適用されたときに発生します。 | [**Table**](/javascript/api/excel/excel.table#onfiltered)、[**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered)、[**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered) |

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

次のコード例では、ワークシートの `onChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録します。 このコードでは、そのワークシートでデータが変更されたときに、`handleDataChange` 関数を実行するように指定しています。

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

次のコード例では、ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。 また、そのイベント ハンドラーを削除するために、後から呼び出すことができる `remove()` 関数も定義しています。 `RequestContext`イベントハンドラーの作成に使用されたを削除するには、を使用する必要があることに注意してください。 

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

- [Office アドインでの Excel JavaScript オブジェクトモデル](excel-add-ins-core-concepts.md)
