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
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="6dd2e-102">Excel JavaScript API を使用してイベントを操作する</span><span class="sxs-lookup"><span data-stu-id="6dd2e-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="6dd2e-103">この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="6dd2e-104">Excel のイベント</span><span class="sxs-lookup"><span data-stu-id="6dd2e-104">Events in Excel</span></span>

<span data-ttu-id="6dd2e-p101">Excel ブック内で特定の種類の変更が起こるたびに、イベント通知が発生します。Excel の JavaScript API を使用すれば、イベント ハンドラーを登録して、特定のイベントが発生したときに、アドインにより指定された関数が自動的に実行されるようにすることができます。現在、次のイベントがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="6dd2e-108">イベント</span><span class="sxs-lookup"><span data-stu-id="6dd2e-108">Event</span></span> | <span data-ttu-id="6dd2e-109">説明</span><span class="sxs-lookup"><span data-stu-id="6dd2e-109">Description</span></span> | <span data-ttu-id="6dd2e-110">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="6dd2e-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="6dd2e-111">オブジェクトが追加されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="6dd2e-112">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="6dd2e-112">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="6dd2e-113">オブジェクトが削除されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="6dd2e-114">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、 [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="6dd2e-114">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onActivated` | <span data-ttu-id="6dd2e-115">オブジェクトがアクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="6dd2e-116">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart)、[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="6dd2e-116">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeactivated` | <span data-ttu-id="6dd2e-117">オブジェクトが非アクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="6dd2e-118">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart)、[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="6dd2e-118">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span></span> |
| `onCalculated` | <span data-ttu-id="6dd2e-119">ワークシートの計算が終了した（またはコレクションのすべてのワークシートが終了した）ときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-119">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="6dd2e-120">[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="6dd2e-120">**WorksheetCollection**, [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="6dd2e-121">セル内のデータが変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-121">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="6dd2e-122">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、 [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table)、 [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span><span class="sxs-lookup"><span data-stu-id="6dd2e-122">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="6dd2e-123">バインディング内のデータまたは書式設定が変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-123">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="6dd2e-124">**バインディング**</span><span class="sxs-lookup"><span data-stu-id="6dd2e-124">**Binding**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | <span data-ttu-id="6dd2e-125">アクティブなセルまたは選択範囲が変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="6dd2e-126">[ **Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、[**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table)、[ **Binding**](https://docs.microsoft.com/javascript/api/excel/excel.binding)</span><span class="sxs-lookup"><span data-stu-id="6dd2e-126">**Worksheet**, [Table](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="6dd2e-127">ドキュメント内の設定が変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-127">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="6dd2e-128">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="6dd2e-128">**settingCollection**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="6dd2e-129">イベント トリガー</span><span class="sxs-lookup"><span data-stu-id="6dd2e-129">Event triggers</span></span>

<span data-ttu-id="6dd2e-130">Excel ブック内のイベントは、次の事項でトリガーできます。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-130">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="6dd2e-131">ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作</span><span class="sxs-lookup"><span data-stu-id="6dd2e-131">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="6dd2e-132">ブックを変更する Office アドイン (JavaScript) コード</span><span class="sxs-lookup"><span data-stu-id="6dd2e-132">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="6dd2e-133">ブックを変更する VBA アドイン (マクロ) コード</span><span class="sxs-lookup"><span data-stu-id="6dd2e-133">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="6dd2e-134">Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-134">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="6dd2e-135">イベント ハンドラーのライフサイクル</span><span class="sxs-lookup"><span data-stu-id="6dd2e-135">Lifecycle of an event handler</span></span>

<span data-ttu-id="6dd2e-p102">イベント ハンドラーは、アドインでイベント ハンドラーを登録するときに作成されます。アドインでイベント ハンドラーの登録を解除したとき、またはアドインが更新、再読み込みされた場合、閉じられたときに破棄されます。イベント ハンドラーは、Excel ファイルのまたは複数の Excel のオンラインでのセッションの一部として保持はされません。
</span><span class="sxs-lookup"><span data-stu-id="6dd2e-p102">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

> [!CAUTION]
> <span data-ttu-id="6dd2e-139">イベントが登録されているオブジェクトが削除されたとき（たとえば`onChanged` イベントが登録されたテーブル）、 イベントハンドラがトリガーしなくなりますが、アドインセッションまたはExcel のセッションを更新または閉じるまでメモリに残ります。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-139">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="6dd2e-140">イベントと共同編集</span><span class="sxs-lookup"><span data-stu-id="6dd2e-140">Events and coauthoring</span></span>

<span data-ttu-id="6dd2e-p103">[共同編集機能](co-authoring-in-excel-add-ins.md)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。共同編集でトリガーされ得るイベント (`onChanged` など) の場合、対応する **Event** オブジェクトには **source** プロパティが含まれるようになります。このプロパティは、イベントが現在のユーザーによってローカルにトリガーされた (`event.source = Local`) ものか、リモートの共同作成者によってトリガーされた (`event.source = Remote`) ものかを示します。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="6dd2e-143">イベント ハンドラーの登録</span><span class="sxs-lookup"><span data-stu-id="6dd2e-143">Register an event handler</span></span>

<span data-ttu-id="6dd2e-p104">次のコード サンプルでは、**Sample** という名前のワークシート内の `onChanged` イベントのイベント ハンドラーを登録します。コードでは、ワークシート内でデータが変更されたとき、`handleDataChange` 関数が実行されるよう指定しています。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-p104">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="6dd2e-146">イベントの処理</span><span class="sxs-lookup"><span data-stu-id="6dd2e-146">Handle an event</span></span>

<span data-ttu-id="6dd2e-p105">前の例で示したように、イベント ハンドラーの登録時には、特定のイベントが発生したときに実行する関数を指定します。その関数は、シナリオが必要とする任意のアクションを実行するように作成できます。次のコード サンプルは、単にイベントの情報をコンソールに出力するイベント ハンドラー関数を示します。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-p105">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="6dd2e-150">イベント ハンドラーを削除する</span><span class="sxs-lookup"><span data-stu-id="6dd2e-150">Remove an event handler</span></span>

<span data-ttu-id="6dd2e-p106">次のコード例では、**Sample** という名前ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。また、その後にそのイベント ハンドラーを削除するために呼び出さすことのできる `remove()` 関数も定義しています。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-p106">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="6dd2e-153">イベントの有効化と無効化</span><span class="sxs-lookup"><span data-stu-id="6dd2e-153">Enable and disable events</span></span>

<span data-ttu-id="6dd2e-p107">イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。たとえば、アプリケーションがイベントを受け取る必要がない可能性、または複数のエンティティを一括編集しているときにイベントを無視できる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-p107">The performance of an add-in may be improved by disabling events. For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span> 

<span data-ttu-id="6dd2e-p108">イベントは、 [ランタイム](https://docs.microsoft.com/javascript/api/excel/excel.runtime) レベルで有効または無効にされます。`enableEvents`プロパティは、イベントが発生し、そのハンドラーがアクティブ化されているかどうかを決定します。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-p108">Events are enabled and disabled at the [runtime](https://docs.microsoft.com/javascript/api/excel/excel.runtime) level. The `enableEvents` property determines if events are fired and their handlers are activated.</span></span> 

<span data-ttu-id="6dd2e-158">次のコード サンプルは、イベントのオンとオフを切り替える方法を示します。</span><span class="sxs-lookup"><span data-stu-id="6dd2e-158">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="6dd2e-159">関連項目</span><span class="sxs-lookup"><span data-stu-id="6dd2e-159">See also</span></span>

- [<span data-ttu-id="6dd2e-160">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="6dd2e-160">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)