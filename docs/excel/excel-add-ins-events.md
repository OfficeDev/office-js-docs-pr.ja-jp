---
title: Excel JavaScript API を使用してイベントを操作する
description: ''
ms.date: 05/25/2018
ms.openlocfilehash: fbeb0e6efabe37afb0f73ab8e7448d8cf01ebace
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943979"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="bbc05-102">Excel JavaScript API を使用してイベントを操作する</span><span class="sxs-lookup"><span data-stu-id="bbc05-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="bbc05-103">この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。</span><span class="sxs-lookup"><span data-stu-id="bbc05-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="bbc05-104">Excel のイベント</span><span class="sxs-lookup"><span data-stu-id="bbc05-104">Events in Excel</span></span>

<span data-ttu-id="bbc05-105">Excel ブックで特定の種類の変更が発生するたびに、イベント通知がトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="bbc05-105">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="bbc05-106">Excel JavaScript API を使用すると、イベント ハンドラーを登録できます。このハンドラーによって、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できるようになります。</span><span class="sxs-lookup"><span data-stu-id="bbc05-106">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="bbc05-107">現時点でサポートされているイベントは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="bbc05-107">The following events are currently supported.</span></span>

| <span data-ttu-id="bbc05-108">イベント</span><span class="sxs-lookup"><span data-stu-id="bbc05-108">Event</span></span> | <span data-ttu-id="bbc05-109">説明</span><span class="sxs-lookup"><span data-stu-id="bbc05-109">Description</span></span> | <span data-ttu-id="bbc05-110">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="bbc05-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="bbc05-111">オブジェクトが追加されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-111">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="bbc05-112">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="bbc05-112">**WorksheetCollection**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onDeleted` | <span data-ttu-id="bbc05-113">オブジェクトが削除されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-113">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="bbc05-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="bbc05-114">**WorksheetCollection**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onActivated` | <span data-ttu-id="bbc05-115">オブジェクトがアクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="bbc05-116">[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="bbc05-116">**WorksheetCollection**, [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeactivated` | <span data-ttu-id="bbc05-117">オブジェクトが非アクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="bbc05-118">[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="bbc05-118">**WorksheetCollection**, [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="bbc05-119">セル内のデータが変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="bbc05-120">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、 [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table)、 [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span><span class="sxs-lookup"><span data-stu-id="bbc05-120">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="bbc05-121">バインディング内でデータまたは書式設定が変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="bbc05-122">**バインディング**</span><span class="sxs-lookup"><span data-stu-id="bbc05-122">**Binding**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | <span data-ttu-id="bbc05-123">アクティブなセルまたは選択範囲が変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-123">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="bbc05-124">[ **Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、[**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table)、[ **Binding**](https://docs.microsoft.com/javascript/api/excel/excel.binding)</span><span class="sxs-lookup"><span data-stu-id="bbc05-124">**Worksheet**, [Table](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="bbc05-125">ドキュメント内の設定が変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-125">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="bbc05-126">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="bbc05-126">**SettingCollection**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

## <a name="preview-beta-events-in-excel"></a><span data-ttu-id="bbc05-127">Excel でのプレビュー（ベータ）イベント</span><span class="sxs-lookup"><span data-stu-id="bbc05-127">Preview (Beta) Events in Excel</span></span>

> [!NOTE]
> <span data-ttu-id="bbc05-128">これらのイベントは現在、公開プレビュー (ベータ版) でのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="bbc05-128">These samples use APIs currently available only in public preview (beta).</span></span> <span data-ttu-id="bbc05-129">これらの機能を使用するには、Office.js CDN のベータ ライブラリを使用する必要があります。 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="bbc05-129">To use these features, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

| <span data-ttu-id="bbc05-130">イベント</span><span class="sxs-lookup"><span data-stu-id="bbc05-130">Event</span></span> | <span data-ttu-id="bbc05-131">説明</span><span class="sxs-lookup"><span data-stu-id="bbc05-131">Description</span></span> | <span data-ttu-id="bbc05-132">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="bbc05-132">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="bbc05-133">グラフが追加されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-133">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="bbc05-134">**ChartCollection**</span><span class="sxs-lookup"><span data-stu-id="bbc05-134">**chartCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onDeleted` | <span data-ttu-id="bbc05-135">グラフが削除されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-135">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="bbc05-136">**ChartCollection**</span><span class="sxs-lookup"><span data-stu-id="bbc05-136">**chartCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onActivated` | <span data-ttu-id="bbc05-137">グラフがアクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-137">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="bbc05-138">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)、 [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="bbc05-138">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |
| `onDeactivated` | <span data-ttu-id="bbc05-139">グラフが非アクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-139">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="bbc05-140">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)、 [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="bbc05-140">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |
| `onCalculated` | <span data-ttu-id="bbc05-141">ワークシートの計算が終了した（またはコレクションのすべてのワークシートが終了した）ときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="bbc05-141">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="bbc05-142">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)、[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="bbc05-142">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="bbc05-143">イベント トリガー</span><span class="sxs-lookup"><span data-stu-id="bbc05-143">Event triggers</span></span>

<span data-ttu-id="bbc05-144">Excel ブックのイベントは、次の事項でトリガーできます。</span><span class="sxs-lookup"><span data-stu-id="bbc05-144">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="bbc05-145">ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作</span><span class="sxs-lookup"><span data-stu-id="bbc05-145">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="bbc05-146">ブックを変更する Office アドイン (JavaScript) コード</span><span class="sxs-lookup"><span data-stu-id="bbc05-146">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="bbc05-147">ブックを変更する VBA アドイン (マクロ) コード</span><span class="sxs-lookup"><span data-stu-id="bbc05-147">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="bbc05-148">Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="bbc05-148">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="bbc05-149">イベント ハンドラーのライフサイクル</span><span class="sxs-lookup"><span data-stu-id="bbc05-149">Lifecycle of an event handler</span></span>

<span data-ttu-id="bbc05-p103">イベント ハンドラーは、アドインでイベント ハンドラーを登録するときに作成され、アドインでイベント ハンドラーの登録を解除したとき、またはアドインが閉じられたときに破棄されます。イベント ハンドラーは、Excel ファイルの一部として保持されません。</span><span class="sxs-lookup"><span data-stu-id="bbc05-p103">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="bbc05-152">イベントと共同編集</span><span class="sxs-lookup"><span data-stu-id="bbc05-152">Events and coauthoring</span></span>

<span data-ttu-id="bbc05-p104">[共同編集機能](co-authoring-in-excel-add-ins.md)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。共同編集でトリガーできるイベント (`onChanged` など) の場合、対応する **Event** オブジェクトには **source** プロパティが含まれるようになります。このプロパティは、イベントが現在のユーザーによってローカルにトリガーされた (`event.source = Local`) ものか、リモートの共同作成者によってトリガーされた (`event.source = Remote`) ものかを示します。</span><span class="sxs-lookup"><span data-stu-id="bbc05-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="bbc05-155">イベント ハンドラーの登録</span><span class="sxs-lookup"><span data-stu-id="bbc05-155">Register an event handler</span></span>

<span data-ttu-id="bbc05-156">次のコード例では、ワークシートの `onChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録します。</span><span class="sxs-lookup"><span data-stu-id="bbc05-156">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="bbc05-157">このコードでは、そのワークシートでデータが変更されたときに、`handleDataChange` 関数を実行するように指定しています。</span><span class="sxs-lookup"><span data-stu-id="bbc05-157">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="bbc05-158">イベントの処理</span><span class="sxs-lookup"><span data-stu-id="bbc05-158">Handle an event</span></span>

<span data-ttu-id="bbc05-159">前の例で示したように、イベント ハンドラーの登録時には、特定のイベントが発生したときに実行する関数を指定します。</span><span class="sxs-lookup"><span data-stu-id="bbc05-159">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="bbc05-160">その関数は、目的のシナリオに必要なアクションを実行するように設計できます。</span><span class="sxs-lookup"><span data-stu-id="bbc05-160">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="bbc05-161">次のコード例は、イベントに関する情報を単にコンソールに出力するイベント ハンドラー関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="bbc05-161">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="bbc05-162">イベント ハンドラーを削除する</span><span class="sxs-lookup"><span data-stu-id="bbc05-162">Remove an event handler</span></span>

<span data-ttu-id="bbc05-163">次のコード例では、ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。</span><span class="sxs-lookup"><span data-stu-id="bbc05-163">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="bbc05-164">また、そのイベント ハンドラーを削除するために、後から呼び出すことができる `remove()` 関数も定義しています。</span><span class="sxs-lookup"><span data-stu-id="bbc05-164">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="bbc05-165">イベントの有効化と無効化</span><span class="sxs-lookup"><span data-stu-id="bbc05-165">Enable and disable events</span></span>

> [!NOTE]
> <span data-ttu-id="bbc05-166">この機能は現在、公開プレビュー (ベータ版) でのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="bbc05-166">The copyFrom function is currently available only in public preview (beta).</span></span> <span data-ttu-id="bbc05-167">これを使用するには、Office.js CDN のベータ版のライブラリを参照する必要があります: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="bbc05-167">To use these features, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

<span data-ttu-id="bbc05-168">イベントを無効にするアドインのパフォーマンスを向上させる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="bbc05-168">The performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="bbc05-169">たとえば、アプリケーションがイベントを受け取る必要がない可能性、または複数のエンティティを一括編集しているときにイベントを無視する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="bbc05-169">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span> 

<span data-ttu-id="bbc05-170">イベントは、 [ランタイム](https://docs.microsoft.com/javascript/api/excel/excel.runtime) レベルで有効と無効にされます。</span><span class="sxs-lookup"><span data-stu-id="bbc05-170">Events are turned on and off at the [runtime](https://docs.microsoft.com/javascript/api/excel/excel.runtime) level.</span></span> <span data-ttu-id="bbc05-171">|||UNTRANSLATED_CONTENT_START|||The `enableEvents` property determines if events are fired and their handlers are activated.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="bbc05-171">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span> 

<span data-ttu-id="bbc05-172">イベントをオンとオフを切り替える方法を次のコード例に示します。</span><span class="sxs-lookup"><span data-stu-id="bbc05-172">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="bbc05-173">関連項目</span><span class="sxs-lookup"><span data-stu-id="bbc05-173">See also</span></span>

- [<span data-ttu-id="bbc05-174">Excel JavaScript API の中心概念</span><span class="sxs-lookup"><span data-stu-id="bbc05-174">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="bbc05-175">Excel JavaScript API オープン仕様</span><span class="sxs-lookup"><span data-stu-id="bbc05-175">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)