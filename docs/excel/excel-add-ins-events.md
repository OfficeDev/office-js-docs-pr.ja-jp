---
title: Excel JavaScript API を使用してイベントを操作する
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 08653a84c051709d16371d89672d3f7ebe2030b7
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872019"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="a85b6-102">Excel JavaScript API を使用してイベントを操作する</span><span class="sxs-lookup"><span data-stu-id="a85b6-102">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="a85b6-103">この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。</span><span class="sxs-lookup"><span data-stu-id="a85b6-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="a85b6-104">Excel のイベント</span><span class="sxs-lookup"><span data-stu-id="a85b6-104">Events in Excel</span></span>

<span data-ttu-id="a85b6-p101">Excel ブックで特定の種類の変更が発生するたびに、イベント通知がトリガーされます。 Excel JavaScript API を使用すると、イベント ハンドラーを登録できます。このハンドラーによって、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できるようになります。 現時点でサポートされているイベントは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="a85b6-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="a85b6-108">イベント</span><span class="sxs-lookup"><span data-stu-id="a85b6-108">Event</span></span> | <span data-ttu-id="a85b6-109">説明</span><span class="sxs-lookup"><span data-stu-id="a85b6-109">Description</span></span> | <span data-ttu-id="a85b6-110">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="a85b6-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="a85b6-111">オブジェクトが追加されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="a85b6-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="a85b6-112">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="a85b6-112">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="a85b6-113">オブジェクトが削除されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="a85b6-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="a85b6-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="a85b6-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onActivated` | <span data-ttu-id="a85b6-115">オブジェクトがアクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="a85b6-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="a85b6-116">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="a85b6-116">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeactivated` | <span data-ttu-id="a85b6-117">オブジェクトが非アクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="a85b6-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="a85b6-118">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="a85b6-118">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onCalculated` | <span data-ttu-id="a85b6-119">ワークシートの計算が完了したとき (あるいはコレクションのすべてのワークシートが完了したとき) に発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="a85b6-119">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="a85b6-120">[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="a85b6-120">[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onChanged` | <span data-ttu-id="a85b6-121">セル内のデータが変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="a85b6-121">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="a85b6-122">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)</span><span class="sxs-lookup"><span data-stu-id="a85b6-122">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="a85b6-123">バインド内でデータまたは書式設定が変更されるときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="a85b6-123">Event that occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="a85b6-124">**Binding**</span><span class="sxs-lookup"><span data-stu-id="a85b6-124">**Binding**</span></span>](/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | <span data-ttu-id="a85b6-125">アクティブなセルまたは選択範囲が変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="a85b6-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="a85b6-126">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**Table**](/javascript/api/excel/excel.table)、[**Binding**](/javascript/api/excel/excel.binding)</span><span class="sxs-lookup"><span data-stu-id="a85b6-126">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**Binding**](/javascript/api/excel/excel.binding)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="a85b6-127">ドキュメント内の設定が変更されるときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="a85b6-127">Event that occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="a85b6-128">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="a85b6-128">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="a85b6-129">イベント トリガー</span><span class="sxs-lookup"><span data-stu-id="a85b6-129">Event triggers</span></span>

<span data-ttu-id="a85b6-130">Excel ブックのイベントは、次の事項でトリガーできます。</span><span class="sxs-lookup"><span data-stu-id="a85b6-130">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="a85b6-131">ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作</span><span class="sxs-lookup"><span data-stu-id="a85b6-131">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="a85b6-132">ブックを変更する Office アドイン (JavaScript) コード</span><span class="sxs-lookup"><span data-stu-id="a85b6-132">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="a85b6-133">ブックを変更する VBA アドイン (マクロ) コード</span><span class="sxs-lookup"><span data-stu-id="a85b6-133">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="a85b6-134">Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="a85b6-134">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="a85b6-135">イベント ハンドラーのライフサイクル</span><span class="sxs-lookup"><span data-stu-id="a85b6-135">Lifecycle of an event handler</span></span>

<span data-ttu-id="a85b6-136">アドインがイベント ハンドラーを登録すると、そのイベント ハンドラーが作成されます。</span><span class="sxs-lookup"><span data-stu-id="a85b6-136">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="a85b6-137">アドインがイベント ハンドラーを登録解除するか、アドインが更新、再読み込み、または閉じられると、イベント ハンドラーは破棄されます。</span><span class="sxs-lookup"><span data-stu-id="a85b6-137">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="a85b6-138">イベント ハンドラーは Excel ファイルの一部として保持されず、また Excel Online のセッション間でも保持されません。</span><span class="sxs-lookup"><span data-stu-id="a85b6-138">Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.</span></span>

> [!CAUTION]
> <span data-ttu-id="a85b6-139">イベントが登録されているオブジェクト (`onChanged` イベントが登録されているテーブルなど) が削除されると、イベント ハンドラーはトリガーされませんが、アドインまたは Excel セッションが更新または閉じるまではメモリで維持されます。</span><span class="sxs-lookup"><span data-stu-id="a85b6-139">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="a85b6-140">イベントと共同編集</span><span class="sxs-lookup"><span data-stu-id="a85b6-140">Events and coauthoring</span></span>

<span data-ttu-id="a85b6-p103">[共同編集機能](co-authoring-in-excel-add-ins.md)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。共同編集でトリガーできるイベント (`onChanged` など) の場合、対応する **Event** オブジェクトには **source** プロパティが含まれるようになります。このプロパティは、イベントが現在のユーザーによってローカルにトリガーされた (`event.source = Local`) ものか、リモートの共同作成者によってトリガーされた (`event.source = Remote`) ものかを示します。</span><span class="sxs-lookup"><span data-stu-id="a85b6-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="a85b6-143">イベント ハンドラーの登録</span><span class="sxs-lookup"><span data-stu-id="a85b6-143">Register an event handler</span></span>

<span data-ttu-id="a85b6-p104">次のコード例では、ワークシートの `onChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録します。 このコードでは、そのワークシートでデータが変更されたときに、`handleDataChange` 関数を実行するように指定しています。</span><span class="sxs-lookup"><span data-stu-id="a85b6-p104">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="a85b6-146">イベントの処理</span><span class="sxs-lookup"><span data-stu-id="a85b6-146">Handle an event</span></span>

<span data-ttu-id="a85b6-p105">前の例で示したように、イベント ハンドラーの登録時には、特定のイベントが発生したときに実行する関数を指定します。 その関数は、目的のシナリオに必要なアクションを実行するように設計できます。 次のコード例は、イベントに関する情報を単にコンソールに出力するイベント ハンドラー関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="a85b6-p105">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="a85b6-150">イベント ハンドラーを削除する</span><span class="sxs-lookup"><span data-stu-id="a85b6-150">Remove an event handler</span></span>

<span data-ttu-id="a85b6-p106">次のコード例では、ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。 また、そのイベント ハンドラーを削除するために、後から呼び出すことができる `remove()` 関数も定義しています。</span><span class="sxs-lookup"><span data-stu-id="a85b6-p106">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="a85b6-153">イベントの有効化と無効化</span><span class="sxs-lookup"><span data-stu-id="a85b6-153">Enable and disable events</span></span>

<span data-ttu-id="a85b6-154">イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="a85b6-154">The performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="a85b6-155">たとえば、アプリがイベントを受信する必要がないことや、複数エンティティの一括編集を実行中にイベントを無視できるすることがあります。</span><span class="sxs-lookup"><span data-stu-id="a85b6-155">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="a85b6-156">イベントは[ランタイム](/javascript/api/excel/excel.runtime) レベルで有効または無効にできます。</span><span class="sxs-lookup"><span data-stu-id="a85b6-156">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="a85b6-157">`enableEvents` プロパティは、イベントが発生したかどうかと、イベント ハンドラーがアクティブになったかどうかを判別します。</span><span class="sxs-lookup"><span data-stu-id="a85b6-157">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="a85b6-158">次のコード サンプルは、イベントのオンとオフを切り替える方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="a85b6-158">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="a85b6-159">関連項目</span><span class="sxs-lookup"><span data-stu-id="a85b6-159">See also</span></span>

- [<span data-ttu-id="a85b6-160">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="a85b6-160">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
