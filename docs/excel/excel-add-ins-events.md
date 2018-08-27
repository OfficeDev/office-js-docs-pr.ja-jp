---
title: Excel JavaScript API を使用してイベントを操作する
description: ''
ms.date: 05/25/2018
ms.openlocfilehash: 3d94a36a60220b856795b8d0abf5387fcb8c1bad
ms.sourcegitcommit: e1c92ba882e6eb03a165867c6021a6aa742aa310
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/20/2018
ms.locfileid: "22925627"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="95d9e-102">Excel JavaScript API を使用してイベントを操作する</span><span class="sxs-lookup"><span data-stu-id="95d9e-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="95d9e-103">この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。</span><span class="sxs-lookup"><span data-stu-id="95d9e-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="95d9e-104">Excel のイベント</span><span class="sxs-lookup"><span data-stu-id="95d9e-104">Events in Excel</span></span>

<span data-ttu-id="95d9e-105">Excel ブックで特定の種類の変更が発生するたびに、イベント通知がトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="95d9e-105">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="95d9e-106">Excel JavaScript API を使用すると、イベント ハンドラーを登録できます。このハンドラーによって、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できるようになります。</span><span class="sxs-lookup"><span data-stu-id="95d9e-106">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="95d9e-107">現時点でサポートされているイベントは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="95d9e-107">The following events are currently supported.</span></span>

| <span data-ttu-id="95d9e-108">イベント</span><span class="sxs-lookup"><span data-stu-id="95d9e-108">Event</span></span> | <span data-ttu-id="95d9e-109">説明</span><span class="sxs-lookup"><span data-stu-id="95d9e-109">Description</span></span> | <span data-ttu-id="95d9e-110">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="95d9e-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="95d9e-111">オブジェクトが追加されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-111">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="95d9e-112">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="95d9e-112">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onDeleted` | <span data-ttu-id="95d9e-113">オブジェクトが削除されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-113">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="95d9e-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="95d9e-114">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onActivated` | <span data-ttu-id="95d9e-115">オブジェクトがアクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="95d9e-116">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection)、 [**Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="95d9e-116">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onDeactivated` | <span data-ttu-id="95d9e-117">オブジェクトが非アクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="95d9e-118">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection)、 [**Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="95d9e-118">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="95d9e-119">セル内のデータが変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="95d9e-120">[**worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet)、 [**Table**](https://dev.office.com/reference/add-ins/excel/table)、 [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span><span class="sxs-lookup"><span data-stu-id="95d9e-120">[**Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Table**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="95d9e-121">バインディングでデータまたは書式設定が変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="95d9e-122">**Binding**</span><span class="sxs-lookup"><span data-stu-id="95d9e-122">**Binding**</span></span>](https://dev.office.com/reference/add-ins/excel/binding) |
| `onSelectionChanged` | <span data-ttu-id="95d9e-123">アクティブなセルまたは選択範囲が変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-123">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="95d9e-124">[ **Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet)、[**Table**](https://dev.office.com/reference/add-ins/excel/table)、[ **Binding**](https://dev.office.com/reference/add-ins/excel/binding)</span><span class="sxs-lookup"><span data-stu-id="95d9e-124">**Worksheet**, [Table](https://dev.office.com/reference/add-ins/excel/worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="95d9e-125">ドキュメント内の設定が変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-125">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="95d9e-126">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="95d9e-126">**SettingCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/settingcollection) |

## <a name="preview-beta-events-in-excel"></a><span data-ttu-id="95d9e-127">Excel でのプレビュー（ベータ）イベント</span><span class="sxs-lookup"><span data-stu-id="95d9e-127">Preview (Beta) Events in Excel</span></span>

> [!NOTE]
> <span data-ttu-id="95d9e-128">これらのイベントは現在、公開プレビュー (ベータ版) でのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="95d9e-128">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="95d9e-129">これらの機能を使用するには、Office.js CDN のベータ ライブラリを使用する必要があります。 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="95d9e-129">To use these features, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

| <span data-ttu-id="95d9e-130">イベント</span><span class="sxs-lookup"><span data-stu-id="95d9e-130">Event</span></span> | <span data-ttu-id="95d9e-131">説明</span><span class="sxs-lookup"><span data-stu-id="95d9e-131">Description</span></span> | <span data-ttu-id="95d9e-132">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="95d9e-132">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="95d9e-133">グラフが追加されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-133">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="95d9e-134">**ChartCollection**</span><span class="sxs-lookup"><span data-stu-id="95d9e-134">**chartCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onDeleted` | <span data-ttu-id="95d9e-135">グラフが削除されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-135">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="95d9e-136">**ChartCollection**</span><span class="sxs-lookup"><span data-stu-id="95d9e-136">**chartCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onActivated` | <span data-ttu-id="95d9e-137">グラフがアクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-137">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="95d9e-138">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)、 [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="95d9e-138">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |
| `onDeactivated` | <span data-ttu-id="95d9e-139">グラフが非アクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-139">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="95d9e-140">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)、 [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="95d9e-140">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |
| `onCalculated` | <span data-ttu-id="95d9e-141">ワークシートの計算が終了した（またはコレクションのすべてのワークシートが終了した）ときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="95d9e-141">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="95d9e-142">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)、 [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="95d9e-142">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="95d9e-143">イベント トリガー</span><span class="sxs-lookup"><span data-stu-id="95d9e-143">Event triggers</span></span>

<span data-ttu-id="95d9e-144">Excel ブックのイベントは、次の事項でトリガーできます。</span><span class="sxs-lookup"><span data-stu-id="95d9e-144">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="95d9e-145">ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作</span><span class="sxs-lookup"><span data-stu-id="95d9e-145">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="95d9e-146">ブックを変更する Office アドイン (JavaScript) コード</span><span class="sxs-lookup"><span data-stu-id="95d9e-146">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="95d9e-147">ブックを変更する VBA アドイン (マクロ) コード</span><span class="sxs-lookup"><span data-stu-id="95d9e-147">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="95d9e-148">Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="95d9e-148">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="95d9e-149">イベント ハンドラーのライフサイクル</span><span class="sxs-lookup"><span data-stu-id="95d9e-149">Lifecycle of an event handler</span></span>

<span data-ttu-id="95d9e-p103">イベント ハンドラーは、アドインでイベント ハンドラーを登録するときに作成され、アドインでイベント ハンドラーの登録を解除したとき、またはアドインが閉じられたときに破棄されます。イベント ハンドラーは、Excel ファイルの一部として保持されません。</span><span class="sxs-lookup"><span data-stu-id="95d9e-p103">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="95d9e-152">イベントと共同編集</span><span class="sxs-lookup"><span data-stu-id="95d9e-152">Events and coauthoring</span></span>

<span data-ttu-id="95d9e-p104">[共同編集機能](co-authoring-in-excel-add-ins.md)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。共同編集でトリガーできるイベント (`onChanged` など) の場合、対応する **Event** オブジェクトには **source** プロパティが含まれるようになります。このプロパティは、イベントが現在のユーザーによってローカルにトリガーされた (`event.source = Local`) ものか、リモートの共同作成者によってトリガーされた (`event.source = Remote`) ものかを示します。</span><span class="sxs-lookup"><span data-stu-id="95d9e-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="95d9e-155">イベント ハンドラーの登録</span><span class="sxs-lookup"><span data-stu-id="95d9e-155">Register an event handler</span></span>

<span data-ttu-id="95d9e-156">次のコード例では、ワークシートの `onChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録します。</span><span class="sxs-lookup"><span data-stu-id="95d9e-156">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="95d9e-157">このコードでは、そのワークシートでデータが変更されたときに、`handleDataChange` 関数を実行するように指定しています。</span><span class="sxs-lookup"><span data-stu-id="95d9e-157">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="95d9e-158">イベントの処理</span><span class="sxs-lookup"><span data-stu-id="95d9e-158">Handle an event</span></span>

<span data-ttu-id="95d9e-159">前の例で示したように、イベント ハンドラーの登録時には、特定のイベントが発生したときに実行する関数を指定します。</span><span class="sxs-lookup"><span data-stu-id="95d9e-159">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="95d9e-160">その関数は、目的のシナリオに必要なアクションを実行するように設計できます。</span><span class="sxs-lookup"><span data-stu-id="95d9e-160">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="95d9e-161">次のコード例は、イベントに関する情報を単にコンソールに出力するイベント ハンドラー関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="95d9e-161">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="95d9e-162">イベント ハンドラーを削除する</span><span class="sxs-lookup"><span data-stu-id="95d9e-162">Remove an event handler</span></span>

<span data-ttu-id="95d9e-163">次のコード例では、ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。</span><span class="sxs-lookup"><span data-stu-id="95d9e-163">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="95d9e-164">また、そのイベント ハンドラーを削除するために、後から呼び出すことができる `remove()` 関数も定義しています。</span><span class="sxs-lookup"><span data-stu-id="95d9e-164">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="95d9e-165">イベントの有効化と無効化</span><span class="sxs-lookup"><span data-stu-id="95d9e-165">Enable and disable events</span></span>

> [!NOTE]
> <span data-ttu-id="95d9e-166">この機能は現在、公開プレビュー (ベータ版) でのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="95d9e-166">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="95d9e-167">これを使用するには、Office.js CDN のベータ版のライブラリを参照する必要があります: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="95d9e-167">To use these features, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

<span data-ttu-id="95d9e-168">イベントは、 [ランタイム](https://docs.microsoft.com/en-us/javascript/api/excel/excel.runtime?view=office-js) レベルでオンとオフになっています。</span><span class="sxs-lookup"><span data-stu-id="95d9e-168">Events are turned on and off at the [runtime](https://docs.microsoft.com/en-us/javascript/api/excel/excel.runtime?view=office-js) level.</span></span> <span data-ttu-id="95d9e-169"> `enableEvents` プロパティは、イベントが発生し、そのハンドラーがアクティブ化されるかどうかを決定します。</span><span class="sxs-lookup"><span data-stu-id="95d9e-169">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span> <span data-ttu-id="95d9e-170">イベントをオフにすると、パフォーマンスが特に重要な場合、あるいは複数のエンティティを編集していて、完了するまでイベントの発生を避けたい場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="95d9e-170">Turning events off is useful when performance is critical or when editing multiple entities and want to avoid firing events until you have finished.</span></span>

<span data-ttu-id="95d9e-171">イベントをオンとオフを切り替える方法を次のコード例に示します。</span><span class="sxs-lookup"><span data-stu-id="95d9e-171">The following code sample shows how to toggle events on and off.</span></span>

```typescript
async function toggleEvents() {
    await Excel.run(async (context) => {
        context.runtime.load("enableEvents");
        await context.sync();
        const eventBoolean = !context.runtime.enableEvents
        context.runtime.enableEvents = eventBoolean;
        if (eventBoolean) {
            console.log("Events are currently on.");
        } else {
            console.log("Events are currently off.");
        }
        await context.sync();
    });
}
```

## <a name="see-also"></a><span data-ttu-id="95d9e-172">関連項目</span><span class="sxs-lookup"><span data-stu-id="95d9e-172">See also</span></span>

- [<span data-ttu-id="95d9e-173">Excel JavaScript API の中心概念</span><span class="sxs-lookup"><span data-stu-id="95d9e-173">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="95d9e-174">Excel JavaScript API オープン仕様</span><span class="sxs-lookup"><span data-stu-id="95d9e-174">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)