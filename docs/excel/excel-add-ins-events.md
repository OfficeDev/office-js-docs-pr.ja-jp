---
title: Excel JavaScript API を使用してイベントを操作する
description: ''
ms.date: 10/17/2018
localization_priority: Priority
ms.openlocfilehash: 58bb6c01babc19840444a4bee9daef03ad9a7df5
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386527"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してイベントを操作する 

この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。 

## <a name="events-in-excel"></a>Excel のイベント

Excel ブックで特定の種類の変更が発生するたびに、イベント通知がトリガーされます。 Excel JavaScript API を使用すると、イベント ハンドラーを登録できます。このハンドラーによって、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できるようになります。 現時点でサポートされているイベントは次のとおりです。

| イベント | 説明 | サポートされているオブジェクト |
|:---------------|:-------------|:-----------|
| `onAdded` | オブジェクトが追加されたときに発生するイベント。 | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onDeleted` | オブジェクトが削除されたときに発生するイベント。 | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onActivated` | オブジェクトがアクティブ化されたときに発生するイベント。 | [**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart)、[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onDeactivated` | オブジェクトが非アクティブ化されたときに発生するイベント。 | [**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart)、[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onCalculated` | ワークシートの計算が完了したとき (あるいはコレクションのすべてのワークシートが完了したとき) に発生するイベント。 | [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onChanged` | セル内のデータが変更されたときに発生するイベント。 | [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、[**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table)、[**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection) |
| `onDataChanged` | バインド内でデータまたは書式設定が変更されるときに発生するイベント。 | [**Binding**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | アクティブなセルまたは選択範囲が変更されたときに発生するイベント。 | [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、[**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table)、[**Binding**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSettingsChanged` | ドキュメント内の設定が変更されるときに発生するイベント。 | [**SettingCollection**](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a>イベント トリガー

Excel ブックのイベントは、次の事項でトリガーできます。

- ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作
- ブックを変更する Office アドイン (JavaScript) コード
- ブックを変更する VBA アドイン (マクロ) コード

Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。

### <a name="lifecycle-of-an-event-handler"></a>イベント ハンドラーのライフサイクル

アドインがイベント ハンドラーを登録すると、そのイベント ハンドラーが作成されます。 アドインがイベント ハンドラーを登録解除するか、アドインが更新、再読み込み、または閉じられると、イベント ハンドラーは破棄されます。 イベント ハンドラーは Excel ファイルの一部として保持されず、また Excel Online のセッション間でも保持されません。

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

次のコード例では、ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。 また、そのイベント ハンドラーを削除するために、後から呼び出すことができる `remove()` 関数も定義しています。

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

イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。 たとえば、アプリがイベントを受信する必要がないことや、複数エンティティの一括編集を実行中にイベントを無視できるすることがあります。 

イベントは[ランタイム](https://docs.microsoft.com/javascript/api/excel/excel.runtime) レベルで有効または無効にできます。 `enableEvents` プロパティは、イベントが発生したかどうかと、イベント ハンドラーがアクティブになったかどうかを判別します。 

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

- [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
