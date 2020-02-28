---
title: Excel JavaScript API を使用してイベントを操作する
description: Excel JavaScript オブジェクトのイベントのリスト。 これには、イベントハンドラーと関連付けられたパターンの使用に関する情報が含まれます。
ms.date: 02/11/2020
localization_priority: Normal
ms.openlocfilehash: f1a1faf9acc370e7183a078aeeba34019e54900f
ms.sourcegitcommit: d85efbf41a3382ca7d3ab08f2c3f0664d4b26c53
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/28/2020
ms.locfileid: "42327776"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="8ca64-104">Excel JavaScript API を使用してイベントを操作する</span><span class="sxs-lookup"><span data-stu-id="8ca64-104">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="8ca64-105">この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-105">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="8ca64-106">Excel のイベント</span><span class="sxs-lookup"><span data-stu-id="8ca64-106">Events in Excel</span></span>

<span data-ttu-id="8ca64-p102">Excel ブックで特定の種類の変更が発生するたびに、イベント通知がトリガーされます。 Excel JavaScript API を使用すると、イベント ハンドラーを登録できます。このハンドラーによって、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できるようになります。 現時点でサポートされているイベントは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8ca64-p102">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="8ca64-110">イベント</span><span class="sxs-lookup"><span data-stu-id="8ca64-110">Event</span></span> | <span data-ttu-id="8ca64-111">説明</span><span class="sxs-lookup"><span data-stu-id="8ca64-111">Description</span></span> | <span data-ttu-id="8ca64-112">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="8ca64-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="8ca64-113">オブジェクトがアクティブ化されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-113">Occurs when an object is activated.</span></span> | <span data-ttu-id="8ca64-114">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**Shape**](/javascript/api/excel/excel.shape)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-114">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAdded` | <span data-ttu-id="8ca64-115">オブジェクトがコレクションに追加されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-115">Occurs when an object is added to the collection.</span></span> | <span data-ttu-id="8ca64-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="8ca64-117">ブックで `autoSave` の設定が変更されると発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-117">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="8ca64-118">**Workbook**</span><span class="sxs-lookup"><span data-stu-id="8ca64-118">**Workbook**</span></span>](/javascript/api/excel/excel.workbook) |
| `onCalculated` | <span data-ttu-id="8ca64-119">ワークシートの計算が完了したとき (あるいはコレクションのすべてのワークシートが完了したとき) に発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-119">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="8ca64-120">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-120">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="8ca64-121">セル内のデータが変更されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-121">Occurs when data within cells is changed.</span></span> | <span data-ttu-id="8ca64-122">[**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-122">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onColumnSorted` | <span data-ttu-id="8ca64-123">1 つ以上の列を並べ替えたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-123">Occurs when one or more columns have been sorted.</span></span> <span data-ttu-id="8ca64-124">これは、左から右に並べ替えを実行したときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-124">This happens as the result of a left-to-right sort operation.</span></span> | <span data-ttu-id="8ca64-125">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-125">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="8ca64-126">バインド内でデータまたは書式設定が変更されるときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-126">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="8ca64-127">**Binding**</span><span class="sxs-lookup"><span data-stu-id="8ca64-127">**Binding**</span></span>](/javascript/api/excel/excel.binding) |
| `onDeactivated` | <span data-ttu-id="8ca64-128">オブジェクトが非アクティブ化されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-128">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="8ca64-129">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**Shape**](/javascript/api/excel/excel.shape)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-129">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="8ca64-130">オブジェクトがコレクションから削除されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-130">Occurs when an object is deleted from the collection.</span></span> | <span data-ttu-id="8ca64-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFormatChanged` | <span data-ttu-id="8ca64-132">ワークシートで書式設定が変更されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-132">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="8ca64-133">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-133">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onRowSorted` | <span data-ttu-id="8ca64-134">1 つ以上の行を並べ替えたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-134">Occurs when one or more rows have been sorted.</span></span> <span data-ttu-id="8ca64-135">これは、上から下に並べ替えを実行したときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-135">This happens as the result of a top-to-bottom sort operation.</span></span> | <span data-ttu-id="8ca64-136">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-136">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="8ca64-137">アクティブなセルまたは選択範囲が変更されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-137">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="8ca64-138">[**Binding**](/javascript/api/excel/excel.binding)、[**Table**](/javascript/api/excel/excel.table)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-138">[**Binding**](/javascript/api/excel/excel.binding), [**Table**](/javascript/api/excel/excel.table),  [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="8ca64-139">ドキュメント内の設定が変更されるときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-139">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="8ca64-140">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="8ca64-140">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection) |
| `onSingleClicked` | <span data-ttu-id="8ca64-141">ワークシートで左クリック / タップされたアクションが発生したときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-141">Occurs when left-clicked/tapped action occurs in the worksheet.</span></span> | <span data-ttu-id="8ca64-142">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-142">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |

> [!WARNING]
> <span data-ttu-id="8ca64-143">`onSelectionChanged` は現在不安定です。</span><span class="sxs-lookup"><span data-stu-id="8ca64-143">`onSelectionChanged` is currently unstable.</span></span> <span data-ttu-id="8ca64-144">`onSelectionChanged` を確実に使用するための回避策があります。</span><span class="sxs-lookup"><span data-stu-id="8ca64-144">There is a workaround to reliably use `onSelectionChanged`.</span></span> <span data-ttu-id="8ca64-145">HTML ホーム ページの `<head>` セクションに次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-145">Add the following code to the `<head>` section of your HTML home page:</span></span>
>
> ```HTML
> <script> MutationObserver=null; </script>
> ```
>
> <span data-ttu-id="8ca64-146">この問題に関する説明は、「[office-js GitHub リポジトリ](https://github.com/OfficeDev/office-js/issues/533)」にあります。</span><span class="sxs-lookup"><span data-stu-id="8ca64-146">A full discussion of the issue can be found on the [office-js GitHub repo](https://github.com/OfficeDev/office-js/issues/533).</span></span>

### <a name="events-in-preview"></a><span data-ttu-id="8ca64-147">プレビューでのイベント</span><span class="sxs-lookup"><span data-stu-id="8ca64-147">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="8ca64-148">次のイベントは現在、パブリック プレビューでのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="8ca64-148">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="8ca64-149">イベント</span><span class="sxs-lookup"><span data-stu-id="8ca64-149">Event</span></span> | <span data-ttu-id="8ca64-150">説明</span><span class="sxs-lookup"><span data-stu-id="8ca64-150">Description</span></span> | <span data-ttu-id="8ca64-151">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="8ca64-151">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onFiltered` | <span data-ttu-id="8ca64-152">フィルターがオブジェクトに適用されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-152">Occurs when a filter is applied to an object.</span></span> | <span data-ttu-id="8ca64-153">[**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-153">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onRowHiddenChanged` | <span data-ttu-id="8ca64-154">特定のワークシート上の行非表示状態が変更されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-154">Occurs when the row-hidden state changes on a specific worksheet.</span></span> | <span data-ttu-id="8ca64-155">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="8ca64-155">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="8ca64-156">イベント トリガー</span><span class="sxs-lookup"><span data-stu-id="8ca64-156">Event triggers</span></span>

<span data-ttu-id="8ca64-157">Excel ブックのイベントは、次の事項でトリガーできます。</span><span class="sxs-lookup"><span data-stu-id="8ca64-157">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="8ca64-158">ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作</span><span class="sxs-lookup"><span data-stu-id="8ca64-158">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="8ca64-159">ブックを変更する Office アドイン (JavaScript) コード</span><span class="sxs-lookup"><span data-stu-id="8ca64-159">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="8ca64-160">ブックを変更する VBA アドイン (マクロ) コード</span><span class="sxs-lookup"><span data-stu-id="8ca64-160">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="8ca64-161">Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="8ca64-161">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="8ca64-162">イベント ハンドラーのライフサイクル</span><span class="sxs-lookup"><span data-stu-id="8ca64-162">Lifecycle of an event handler</span></span>

<span data-ttu-id="8ca64-163">アドインがイベント ハンドラーを登録すると、そのイベント ハンドラーが作成されます。</span><span class="sxs-lookup"><span data-stu-id="8ca64-163">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="8ca64-164">アドインがイベント ハンドラーを登録解除するか、アドインが更新、再読み込み、または閉じられると、イベント ハンドラーは破棄されます。</span><span class="sxs-lookup"><span data-stu-id="8ca64-164">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="8ca64-165">イベント ハンドラーは Excel ファイルの一部として保持されず、また Excel on the web のセッション間でも保持されません。</span><span class="sxs-lookup"><span data-stu-id="8ca64-165">Event handlers do not persist as part of the Excel file, or across sessions with Excel on the web.</span></span>

> [!CAUTION]
> <span data-ttu-id="8ca64-166">イベントが登録されているオブジェクト (`onChanged` イベントが登録されているテーブルなど) が削除されると、イベント ハンドラーはトリガーされませんが、アドインまたは Excel セッションが更新または閉じるまではメモリで維持されます。</span><span class="sxs-lookup"><span data-stu-id="8ca64-166">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="8ca64-167">イベントと共同編集</span><span class="sxs-lookup"><span data-stu-id="8ca64-167">Events and coauthoring</span></span>

<span data-ttu-id="8ca64-p108">[共同編集機能](co-authoring-in-excel-add-ins.md)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。共同編集でトリガーできるイベント (`onChanged` など) の場合、対応する **Event** オブジェクトには **source** プロパティが含まれるようになります。このプロパティは、イベントが現在のユーザーによってローカルにトリガーされた (`event.source = Local`) ものか、リモートの共同作成者によってトリガーされた (`event.source = Remote`) ものかを示します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-p108">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="8ca64-170">イベント ハンドラーの登録</span><span class="sxs-lookup"><span data-stu-id="8ca64-170">Register an event handler</span></span>

<span data-ttu-id="8ca64-p109">次のコード例では、ワークシートの `onChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録します。 このコードでは、そのワークシートでデータが変更されたときに、`handleDataChange` 関数を実行するように指定しています。</span><span class="sxs-lookup"><span data-stu-id="8ca64-p109">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="8ca64-173">イベントの処理</span><span class="sxs-lookup"><span data-stu-id="8ca64-173">Handle an event</span></span>

<span data-ttu-id="8ca64-p110">前の例で示したように、イベント ハンドラーの登録時には、特定のイベントが発生したときに実行する関数を指定します。 その関数は、目的のシナリオに必要なアクションを実行するように設計できます。 次のコード例は、イベントに関する情報を単にコンソールに出力するイベント ハンドラー関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="8ca64-p110">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="8ca64-177">イベント ハンドラーを削除する</span><span class="sxs-lookup"><span data-stu-id="8ca64-177">Remove an event handler</span></span>

<span data-ttu-id="8ca64-178">次のコード例では、ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。</span><span class="sxs-lookup"><span data-stu-id="8ca64-178">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="8ca64-179">また、そのイベント ハンドラーを削除するために、後から呼び出すことができる `remove()` 関数も定義しています。</span><span class="sxs-lookup"><span data-stu-id="8ca64-179">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span> <span data-ttu-id="8ca64-180">イベントハンドラーの`RequestContext`作成に使用されたを削除するには、を使用する必要があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="8ca64-180">Note that the `RequestContext` used to create the event handler is needed to remove it.</span></span> 

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="8ca64-181">イベントの有効化と無効化</span><span class="sxs-lookup"><span data-stu-id="8ca64-181">Enable and disable events</span></span>

<span data-ttu-id="8ca64-182">イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="8ca64-182">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="8ca64-183">たとえば、アプリがイベントを受信する必要がないことや、複数エンティティの一括編集を実行中にイベントを無視できるすることがあります。</span><span class="sxs-lookup"><span data-stu-id="8ca64-183">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="8ca64-184">イベントは[ランタイム](/javascript/api/excel/excel.runtime) レベルで有効または無効にできます。</span><span class="sxs-lookup"><span data-stu-id="8ca64-184">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="8ca64-185">`enableEvents` プロパティは、イベントが発生したかどうかと、イベント ハンドラーがアクティブになったかどうかを判別します。</span><span class="sxs-lookup"><span data-stu-id="8ca64-185">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="8ca64-186">次のコード サンプルは、イベントのオンとオフを切り替える方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8ca64-186">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="8ca64-187">関連項目</span><span class="sxs-lookup"><span data-stu-id="8ca64-187">See also</span></span>

- [<span data-ttu-id="8ca64-188">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="8ca64-188">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
