---
title: Excel JavaScript API を使用してイベントを操作する
description: JavaScript オブジェクトのイベントExcel一覧です。 これには、イベント ハンドラーと関連付けられたパターンの使用に関する情報が含まれます。
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: c3e1e2c1a0393316166eda6b695d04fbb9876ffd
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290769"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="cb4bf-104">Excel JavaScript API を使用してイベントを操作する</span><span class="sxs-lookup"><span data-stu-id="cb4bf-104">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="cb4bf-105">この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-105">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="cb4bf-106">Excel のイベント</span><span class="sxs-lookup"><span data-stu-id="cb4bf-106">Events in Excel</span></span>

<span data-ttu-id="cb4bf-p102">Excel ブックで特定の種類の変更が発生するたびに、イベント通知がトリガーされます。 Excel JavaScript API を使用すると、イベント ハンドラーを登録できます。このハンドラーによって、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できるようになります。 現時点でサポートされているイベントは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-p102">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="cb4bf-110">イベント</span><span class="sxs-lookup"><span data-stu-id="cb4bf-110">Event</span></span> | <span data-ttu-id="cb4bf-111">説明</span><span class="sxs-lookup"><span data-stu-id="cb4bf-111">Description</span></span> | <span data-ttu-id="cb4bf-112">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb4bf-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="cb4bf-113">オブジェクトがアクティブ化されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-113">Occurs when an object is activated.</span></span> | <span data-ttu-id="cb4bf-114">[**Chart**](/javascript/api/excel/excel.chart#onactivated)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated)、[**Shape**](/javascript/api/excel/excel.shape#onactivated)、[**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-114">[**Chart**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span></span> |
| `onActivated` | <span data-ttu-id="cb4bf-115">ブックがアクティブ化されると発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-115">Occurs when a workbook is activated.</span></span> | [<span data-ttu-id="cb4bf-116">**Workbook**</span><span class="sxs-lookup"><span data-stu-id="cb4bf-116">**Workbook**</span></span>](/javascript/api/excel/excel.workbook#onActivated) |
| `onAdded` | <span data-ttu-id="cb4bf-117">オブジェクトがコレクションに追加されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-117">Occurs when an object is added to the collection.</span></span> | <span data-ttu-id="cb4bf-118">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-118">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="cb4bf-119">ブックで `autoSave` の設定が変更されると発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-119">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="cb4bf-120">**Workbook**</span><span class="sxs-lookup"><span data-stu-id="cb4bf-120">**Workbook**</span></span>](/javascript/api/excel/excel.workbook#onautosavesettingchanged) |
| `onCalculated` | <span data-ttu-id="cb4bf-121">ワークシートの計算が完了したとき (あるいはコレクションのすべてのワークシートが完了したとき) に発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-121">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="cb4bf-122">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-122">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span></span> |
| `onChanged` | <span data-ttu-id="cb4bf-123">個々のセルまたはコメントのデータが変更された場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-123">Occurs when the data of individual cells or comments has changed.</span></span> | <span data-ttu-id="cb4bf-124">[**CommentCollection**](/javascript/api/excel/excel.commentcollection#onchanged), [**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-124">[**CommentCollection**](/javascript/api/excel/excel.commentcollection#onchanged), [**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span></span> |
| `onColumnSorted` | <span data-ttu-id="cb4bf-125">1 つ以上の列を並べ替えたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-125">Occurs when one or more columns have been sorted.</span></span> <span data-ttu-id="cb4bf-126">これは、左から右に並べ替えを実行したときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-126">This happens as the result of a left-to-right sort operation.</span></span> | <span data-ttu-id="cb4bf-127">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-127">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span></span> |
| `onDataChanged` | <span data-ttu-id="cb4bf-128">バインド内でデータまたは書式設定が変更されるときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-128">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="cb4bf-129">**Binding**</span><span class="sxs-lookup"><span data-stu-id="cb4bf-129">**Binding**</span></span>](/javascript/api/excel/excel.binding#ondatachanged) |
| `onDeactivated` | <span data-ttu-id="cb4bf-130">オブジェクトが非アクティブ化されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-130">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="cb4bf-131">[**Chart**](/javascript/api/excel/excel.chart#ondeactivated)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated)、[**Shape**](/javascript/api/excel/excel.shape#ondeactivated)、[**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-131">[**Chart**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span></span> |
| `onDeleted` | <span data-ttu-id="cb4bf-132">オブジェクトがコレクションから削除されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-132">Occurs when an object is deleted from the collection.</span></span> | <span data-ttu-id="cb4bf-133">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-133">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span></span> |
| `onFormatChanged` | <span data-ttu-id="cb4bf-134">ワークシートで書式設定が変更されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-134">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="cb4bf-135">[**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-135">[**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span></span> |
| `onFormulaChanged` | <span data-ttu-id="cb4bf-136">数式が変更された場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-136">Occurs when a formula is changed.</span></span> | <span data-ttu-id="cb4bf-137">[**Worksheet**](/javascript/api/excel/excel.worksheet#onFormulaChanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-137">[**Worksheet**](/javascript/api/excel/excel.worksheet#onFormulaChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged)</span></span> |
| `onRowSorted` | <span data-ttu-id="cb4bf-138">1 つ以上の行を並べ替えたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-138">Occurs when one or more rows have been sorted.</span></span> <span data-ttu-id="cb4bf-139">これは、上から下に並べ替えを実行したときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-139">This happens as the result of a top-to-bottom sort operation.</span></span> | <span data-ttu-id="cb4bf-140">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-140">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="cb4bf-141">アクティブなセルまたは選択範囲が変更されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-141">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="cb4bf-142">[**Binding**](/javascript/api/excel/excel.binding#onselectionchanged), [**Table**](/javascript/api/excel/excel.table#onselectionchanged), [**Workbook**](/javascript/api/excel/excel.workbook#onselectionchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-142">[**Binding**](/javascript/api/excel/excel.binding#onselectionchanged), [**Table**](/javascript/api/excel/excel.table#onselectionchanged), [**Workbook**](/javascript/api/excel/excel.workbook#onselectionchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)</span></span> |
| `onRowHiddenChanged` | <span data-ttu-id="cb4bf-143">特定のワークシート上の行非表示状態が変更されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-143">Occurs when the row-hidden state changes on a specific worksheet.</span></span> | <span data-ttu-id="cb4bf-144">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-144">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="cb4bf-145">ドキュメント内の設定が変更されるときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-145">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="cb4bf-146">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="cb4bf-146">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection#onsettingschanged) |
| `onSingleClicked` | <span data-ttu-id="cb4bf-147">ワークシートで左クリック / タップされたアクションが発生したときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-147">Occurs when left-clicked/tapped action occurs in the worksheet.</span></span> | <span data-ttu-id="cb4bf-148">[**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-148">[**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span></span> |

### <a name="events-in-preview"></a><span data-ttu-id="cb4bf-149">プレビューでのイベント</span><span class="sxs-lookup"><span data-stu-id="cb4bf-149">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="cb4bf-150">次のイベントは現在、パブリック プレビューでのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-150">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="cb4bf-151">イベント</span><span class="sxs-lookup"><span data-stu-id="cb4bf-151">Event</span></span> | <span data-ttu-id="cb4bf-152">説明</span><span class="sxs-lookup"><span data-stu-id="cb4bf-152">Description</span></span> | <span data-ttu-id="cb4bf-153">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb4bf-153">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onFiltered` | <span data-ttu-id="cb4bf-154">フィルターがオブジェクトに適用されたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-154">Occurs when a filter is applied to an object.</span></span> | <span data-ttu-id="cb4bf-155">[**Table**](/javascript/api/excel/excel.table#onfiltered)、[**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered)、[**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span><span class="sxs-lookup"><span data-stu-id="cb4bf-155">[**Table**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="cb4bf-156">イベント トリガー</span><span class="sxs-lookup"><span data-stu-id="cb4bf-156">Event triggers</span></span>

<span data-ttu-id="cb4bf-157">Excel ブックのイベントは、次の事項でトリガーできます。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-157">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="cb4bf-158">ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作</span><span class="sxs-lookup"><span data-stu-id="cb4bf-158">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="cb4bf-159">ブックを変更する Office アドイン (JavaScript) コード</span><span class="sxs-lookup"><span data-stu-id="cb4bf-159">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="cb4bf-160">ブックを変更する VBA アドイン (マクロ) コード</span><span class="sxs-lookup"><span data-stu-id="cb4bf-160">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="cb4bf-161">Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-161">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="cb4bf-162">イベント ハンドラーのライフサイクル</span><span class="sxs-lookup"><span data-stu-id="cb4bf-162">Lifecycle of an event handler</span></span>

<span data-ttu-id="cb4bf-163">アドインがイベント ハンドラーを登録すると、そのイベント ハンドラーが作成されます。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-163">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="cb4bf-164">アドインがイベント ハンドラーを登録解除するか、アドインが更新、再読み込み、または閉じられると、イベント ハンドラーは破棄されます。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-164">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="cb4bf-165">イベント ハンドラーは Excel ファイルの一部として保持されず、また Excel on the web のセッション間でも保持されません。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-165">Event handlers do not persist as part of the Excel file, or across sessions with Excel on the web.</span></span>

> [!CAUTION]
> <span data-ttu-id="cb4bf-166">イベントが登録されているオブジェクト (`onChanged` イベントが登録されているテーブルなど) が削除されると、イベント ハンドラーはトリガーされませんが、アドインまたは Excel セッションが更新または閉じるまではメモリで維持されます。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-166">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="cb4bf-167">イベントと共同編集</span><span class="sxs-lookup"><span data-stu-id="cb4bf-167">Events and coauthoring</span></span>

<span data-ttu-id="cb4bf-p107">[共同編集機能](co-authoring-in-excel-add-ins.md)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。共同編集でトリガーできるイベント (`onChanged` など) の場合、対応する **Event** オブジェクトには **source** プロパティが含まれるようになります。このプロパティは、イベントが現在のユーザーによってローカルにトリガーされた (`event.source = Local`) ものか、リモートの共同作成者によってトリガーされた (`event.source = Remote`) ものかを示します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-p107">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="cb4bf-170">イベント ハンドラーの登録</span><span class="sxs-lookup"><span data-stu-id="cb4bf-170">Register an event handler</span></span>

<span data-ttu-id="cb4bf-p108">次のコード例では、ワークシートの `onChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録します。 このコードでは、そのワークシートでデータが変更されたときに、`handleChange` 関数を実行するように指定しています。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-p108">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="cb4bf-173">イベントの処理</span><span class="sxs-lookup"><span data-stu-id="cb4bf-173">Handle an event</span></span>

<span data-ttu-id="cb4bf-p109">前の例で示したように、イベント ハンドラーの登録時には、特定のイベントが発生したときに実行する関数を指定します。 その関数は、目的のシナリオに必要なアクションを実行するように設計できます。 次のコード例は、イベントに関する情報を単にコンソールに出力するイベント ハンドラー関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-p109">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="cb4bf-177">イベント ハンドラーを削除する</span><span class="sxs-lookup"><span data-stu-id="cb4bf-177">Remove an event handler</span></span>

<span data-ttu-id="cb4bf-178">次のコード例では、ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-178">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="cb4bf-179">また、そのイベント ハンドラーを削除するために、後から呼び出すことができる `remove()` 関数も定義しています。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-179">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span> <span data-ttu-id="cb4bf-180">イベント ハンドラーを `RequestContext` 削除するには、イベント ハンドラーの作成に使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-180">Note that the `RequestContext` used to create the event handler is needed to remove it.</span></span> 

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="cb4bf-181">イベントの有効化と無効化</span><span class="sxs-lookup"><span data-stu-id="cb4bf-181">Enable and disable events</span></span>

<span data-ttu-id="cb4bf-182">イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-182">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="cb4bf-183">たとえば、アプリがイベントを受信する必要がないことや、複数エンティティの一括編集を実行中にイベントを無視できるすることがあります。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-183">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="cb4bf-184">イベントは[ランタイム](/javascript/api/excel/excel.runtime) レベルで有効または無効にできます。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-184">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="cb4bf-185">`enableEvents` プロパティは、イベントが発生したかどうかと、イベント ハンドラーがアクティブになったかどうかを判別します。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-185">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="cb4bf-186">次のコード サンプルは、イベントのオンとオフを切り替える方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="cb4bf-186">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="cb4bf-187">関連項目</span><span class="sxs-lookup"><span data-stu-id="cb4bf-187">See also</span></span>

- [<span data-ttu-id="cb4bf-188">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="cb4bf-188">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
