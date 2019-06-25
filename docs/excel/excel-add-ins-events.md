---
title: Excel JavaScript API を使用してイベントを操作する
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 3757b4e094bae29bc02f2b360144513fbc7ad1ab
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127787"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="61c4d-102">Excel JavaScript API を使用してイベントを操作する</span><span class="sxs-lookup"><span data-stu-id="61c4d-102">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="61c4d-103">この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="61c4d-104">Excel のイベント</span><span class="sxs-lookup"><span data-stu-id="61c4d-104">Events in Excel</span></span>

<span data-ttu-id="61c4d-p101">Excel ブックで特定の種類の変更が発生するたびに、イベント通知がトリガーされます。 Excel JavaScript API を使用すると、イベント ハンドラーを登録できます。このハンドラーによって、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できるようになります。 現時点でサポートされているイベントは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="61c4d-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="61c4d-108">イベント</span><span class="sxs-lookup"><span data-stu-id="61c4d-108">Event</span></span> | <span data-ttu-id="61c4d-109">説明</span><span class="sxs-lookup"><span data-stu-id="61c4d-109">Description</span></span> | <span data-ttu-id="61c4d-110">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="61c4d-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="61c4d-111">オブジェクトがアクティブ化されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-111">Occurs when an object is activated.</span></span> | <span data-ttu-id="61c4d-112">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**Shape**](/javascript/api/excel/excel.shape)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="61c4d-112">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.shape), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onAdded` | <span data-ttu-id="61c4d-113">オブジェクトがコレクションに追加されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-113">Occurs when a view is added to the collection.</span></span> | <span data-ttu-id="61c4d-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="61c4d-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="61c4d-115">ブックで `autoSave` の設定が変更されると発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-115">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="61c4d-116">**Workbook**</span><span class="sxs-lookup"><span data-stu-id="61c4d-116">**Workbook**</span></span>](/javascript/api/excel/excel.workbook) |
| `onCalculated` | <span data-ttu-id="61c4d-117">ワークシートの計算が完了したとき (あるいはコレクションのすべてのワークシートが完了したとき) に発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-117">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="61c4d-118">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="61c4d-118">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="61c4d-119">セル内のデータが変更されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-119">Occurs when data within cells is changed.</span></span> | <span data-ttu-id="61c4d-120">[**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="61c4d-120">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="61c4d-121">バインド内でデータまたは書式設定が変更されるときに発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="61c4d-122">**Binding**</span><span class="sxs-lookup"><span data-stu-id="61c4d-122">**Binding**</span></span>](/javascript/api/excel/excel.binding) |
| `onDeactivated` | <span data-ttu-id="61c4d-123">オブジェクトが非アクティブ化されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-123">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="61c4d-124">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**Shape**](/javascript/api/excel/excel.shape)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="61c4d-124">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.shape), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeleted` | <span data-ttu-id="61c4d-125">オブジェクトがコレクションから削除されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-125">Occurs when an item is deleted from the specified collection.</span></span> | <span data-ttu-id="61c4d-126">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="61c4d-126">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFormatChanged` | <span data-ttu-id="61c4d-127">ワークシートで書式設定が変更されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-127">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="61c4d-128">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="61c4d-128">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="61c4d-129">アクティブなセルまたは選択範囲が変更されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-129">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="61c4d-130">[**Binding**](/javascript/api/excel/excel.binding)、[**Table**](/javascript/api/excel/excel.table)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="61c4d-130">[**Table**](/javascript/api/excel/excel.binding), [**TableCollection**](/javascript/api/excel/excel.table), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="61c4d-131">ドキュメント内の設定が変更されるときに発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-131">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="61c4d-132">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="61c4d-132">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection) |

### <a name="events-in-preview"></a><span data-ttu-id="61c4d-133">プレビューでのイベント</span><span class="sxs-lookup"><span data-stu-id="61c4d-133">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="61c4d-134">次のイベントは現在、パブリック プレビューでのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="61c4d-134">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="61c4d-135">イベント</span><span class="sxs-lookup"><span data-stu-id="61c4d-135">Event</span></span> | <span data-ttu-id="61c4d-136">説明</span><span class="sxs-lookup"><span data-stu-id="61c4d-136">Description</span></span> | <span data-ttu-id="61c4d-137">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="61c4d-137">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onFiltered` | <span data-ttu-id="61c4d-138">フィルターがオブジェクトに適用されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-138">Occurs when filter is applied on an object.</span></span> | <span data-ttu-id="61c4d-139">[**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="61c4d-139">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="61c4d-140">イベント トリガー</span><span class="sxs-lookup"><span data-stu-id="61c4d-140">Event triggers</span></span>

<span data-ttu-id="61c4d-141">Excel ブックのイベントは、次の事項でトリガーできます。</span><span class="sxs-lookup"><span data-stu-id="61c4d-141">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="61c4d-142">ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作</span><span class="sxs-lookup"><span data-stu-id="61c4d-142">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="61c4d-143">ブックを変更する Office アドイン (JavaScript) コード</span><span class="sxs-lookup"><span data-stu-id="61c4d-143">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="61c4d-144">ブックを変更する VBA アドイン (マクロ) コード</span><span class="sxs-lookup"><span data-stu-id="61c4d-144">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="61c4d-145">Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="61c4d-145">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="61c4d-146">イベント ハンドラーのライフサイクル</span><span class="sxs-lookup"><span data-stu-id="61c4d-146">Lifecycle of an event handler</span></span>

<span data-ttu-id="61c4d-147">アドインがイベント ハンドラーを登録すると、そのイベント ハンドラーが作成されます。</span><span class="sxs-lookup"><span data-stu-id="61c4d-147">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="61c4d-148">アドインがイベント ハンドラーを登録解除するか、アドインが更新、再読み込み、または閉じられると、イベント ハンドラーは破棄されます。</span><span class="sxs-lookup"><span data-stu-id="61c4d-148">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="61c4d-149">イベント ハンドラーは Excel ファイルの一部として保持されず、また Excel on the web のセッション間でも保持されません。</span><span class="sxs-lookup"><span data-stu-id="61c4d-149">Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.</span></span>

> [!CAUTION]
> <span data-ttu-id="61c4d-150">イベントが登録されているオブジェクト (`onChanged` イベントが登録されているテーブルなど) が削除されると、イベント ハンドラーはトリガーされませんが、アドインまたは Excel セッションが更新または閉じるまではメモリで維持されます。</span><span class="sxs-lookup"><span data-stu-id="61c4d-150">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="61c4d-151">イベントと共同編集</span><span class="sxs-lookup"><span data-stu-id="61c4d-151">Events and coauthoring</span></span>

<span data-ttu-id="61c4d-p104">[共同編集機能](co-authoring-in-excel-add-ins.md)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。共同編集でトリガーできるイベント (`onChanged` など) の場合、対応する **Event** オブジェクトには **source** プロパティが含まれるようになります。このプロパティは、イベントが現在のユーザーによってローカルにトリガーされた (`event.source = Local`) ものか、リモートの共同作成者によってトリガーされた (`event.source = Remote`) ものかを示します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="61c4d-154">イベント ハンドラーの登録</span><span class="sxs-lookup"><span data-stu-id="61c4d-154">Register an event handler</span></span>

<span data-ttu-id="61c4d-p105">次のコード例では、ワークシートの `onChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録します。 このコードでは、そのワークシートでデータが変更されたときに、`handleDataChange` 関数を実行するように指定しています。</span><span class="sxs-lookup"><span data-stu-id="61c4d-p105">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="61c4d-157">イベントの処理</span><span class="sxs-lookup"><span data-stu-id="61c4d-157">Handle an event</span></span>

<span data-ttu-id="61c4d-p106">前の例で示したように、イベント ハンドラーの登録時には、特定のイベントが発生したときに実行する関数を指定します。 その関数は、目的のシナリオに必要なアクションを実行するように設計できます。 次のコード例は、イベントに関する情報を単にコンソールに出力するイベント ハンドラー関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="61c4d-p106">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="61c4d-161">イベント ハンドラーを削除する</span><span class="sxs-lookup"><span data-stu-id="61c4d-161">Remove an event handler</span></span>

<span data-ttu-id="61c4d-p107">次のコード例では、ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。 また、そのイベント ハンドラーを削除するために、後から呼び出すことができる `remove()` 関数も定義しています。</span><span class="sxs-lookup"><span data-stu-id="61c4d-p107">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="61c4d-164">イベントの有効化と無効化</span><span class="sxs-lookup"><span data-stu-id="61c4d-164">Enable and disable events</span></span>

<span data-ttu-id="61c4d-165">イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="61c4d-165">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="61c4d-166">たとえば、アプリがイベントを受信する必要がないことや、複数エンティティの一括編集を実行中にイベントを無視できるすることがあります。</span><span class="sxs-lookup"><span data-stu-id="61c4d-166">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="61c4d-167">イベントは[ランタイム](/javascript/api/excel/excel.runtime) レベルで有効または無効にできます。</span><span class="sxs-lookup"><span data-stu-id="61c4d-167">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="61c4d-168">`enableEvents` プロパティは、イベントが発生したかどうかと、イベント ハンドラーがアクティブになったかどうかを判別します。</span><span class="sxs-lookup"><span data-stu-id="61c4d-168">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="61c4d-169">次のコード サンプルは、イベントのオンとオフを切り替える方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="61c4d-169">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="61c4d-170">関連項目</span><span class="sxs-lookup"><span data-stu-id="61c4d-170">See also</span></span>

- [<span data-ttu-id="61c4d-171">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="61c4d-171">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
