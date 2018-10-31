---
title: Excel JavaScript API を使用してイベントを操作する
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: c3fbdf27dcbedf0d006973e6ebc2e01b02e6cec2
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/19/2018
ms.locfileid: "25639939"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してイベントを操作する 

この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。 

## <a name="events-in-excel"></a>Excel のイベント

Excel ブック内で特定の種類の変更が起こるたびに、イベント通知が発生します。Excel の JavaScript API を使用すれば、イベント ハンドラーを登録して、特定のイベントが発生したときに、アドインにより指定された関数が自動的に実行されるようにすることができます。現在、次のイベントがサポートされています。

| イベント | 説明 | サポートされているオブジェクト |
|:---------------|:-------------|:-----------|
| `onAdded` | オブジェクトが追加されたときに発生するイベント。 | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onDeleted` | オブジェクトが削除されたときに発生するイベント。 | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、 [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onActivated` | オブジェクトがアクティブ化されたときに発生するイベント。 | [**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart)、[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onDeactivated` | オブジェクトが非アクティブ化されたときに発生するイベント。 | [**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart)、[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onCalculated` | ワークシートの計算が終了した（またはコレクションのすべてのワークシートが終了した）ときに発生するイベント。 | [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onChanged` | セル内のデータが変更されたときに発生するイベント。 | [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、 [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table)、 [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection) |
| `onDataChanged` | バインディング内のデータまたは書式設定が変更されたときに発生するイベント。 | [**バインディング**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | アクティブなセルまたは選択範囲が変更されたときに発生するイベント。 | [ **Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、[**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table)、[ **Binding**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSettingsChanged` | ドキュメント内の設定が変更されたときに発生するイベント。 | [**SettingCollection**](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a>イベント トリガー

Excel ブック内のイベントは、次の事項でトリガーできます。

- ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作
- ブックを変更する Office アドイン (JavaScript) コード
- ブックを変更する VBA アドイン (マクロ) コード

Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。

### <a name="lifecycle-of-an-event-handler"></a>イベント ハンドラーのライフサイクル

イベント ハンドラーは、アドインでイベント ハンドラーを登録するときに作成されます。アドインでイベント ハンドラーの登録を解除したとき、またはアドインが更新、再読み込みされた場合、閉じられたときに破棄されます。イベント ハンドラーは、Excel ファイルのまたは複数の Excel のオンラインでのセッションの一部として保持はされません。


> [!CAUTION]
> イベントが登録されているオブジェクトが削除されたとき（たとえば`onChanged` イベントが登録されたテーブル）、 イベントハンドラがトリガーしなくなりますが、アドインセッションまたはExcel のセッションを更新または閉じるまでメモリに残ります。

### <a name="events-and-coauthoring"></a>イベントと共同編集

[共同編集機能](co-authoring-in-excel-add-ins.md)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。共同編集でトリガーされ得るイベント (`onChanged` など) の場合、対応する **Event** オブジェクトには **source** プロパティが含まれるようになります。このプロパティは、イベントが現在のユーザーによってローカルにトリガーされた (`event.source = Local`) ものか、リモートの共同作成者によってトリガーされた (`event.source = Remote`) ものかを示します。

## <a name="register-an-event-handler"></a>イベント ハンドラーの登録

次のコード サンプルでは、**Sample** という名前のワークシート内の `onChanged` イベントのイベント ハンドラーを登録します。コードでは、ワークシート内でデータが変更されたとき、`handleDataChange` 関数が実行されるよう指定しています。

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

前の例で示したように、イベント ハンドラーの登録時には、特定のイベントが発生したときに実行する関数を指定します。その関数は、シナリオが必要とする任意のアクションを実行するように作成できます。次のコード サンプルは、単にイベントの情報をコンソールに出力するイベント ハンドラー関数を示します。 

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

次のコード例では、**Sample** という名前ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。また、その後にそのイベント ハンドラーを削除するために呼び出さすことのできる `remove()` 関数も定義しています。

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

イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。たとえば、アプリケーションがイベントを受け取る必要がない可能性、または複数のエンティティを一括編集しているときにイベントを無視できる可能性があります。 

イベントは、 [ランタイム](https://docs.microsoft.com/javascript/api/excel/excel.runtime) レベルで有効または無効にされます。`enableEvents`プロパティは、イベントが発生し、そのハンドラーがアクティブ化されているかどうかを決定します。 

次のコード サンプルは、イベントのオンとオフを切り替える方法を示します。

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