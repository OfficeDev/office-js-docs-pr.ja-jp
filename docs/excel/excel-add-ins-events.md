---
title: Excel JavaScript API を使用してイベントを操作する
description: ''
ms.date: 05/25/2018
ms.openlocfilehash: 575e4112ed5f55356020eed8327d309fc58cd643
ms.sourcegitcommit: 9685fd83136bd2106f4c5595bda0010bc1b1950b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/19/2018
ms.locfileid: "20596520"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="c10c2-102">Excel JavaScript API を使用してイベントを操作する</span><span class="sxs-lookup"><span data-stu-id="c10c2-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="c10c2-103">この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。</span><span class="sxs-lookup"><span data-stu-id="c10c2-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="c10c2-104">Excel のイベント</span><span class="sxs-lookup"><span data-stu-id="c10c2-104">Events in Excel</span></span>

<span data-ttu-id="c10c2-105">Excel ブックで特定の種類の変更が発生するたびに、イベント通知がトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="c10c2-105">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="c10c2-106">Excel JavaScript API を使用すると、イベント ハンドラーを登録できます。このハンドラーによって、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できるようになります。</span><span class="sxs-lookup"><span data-stu-id="c10c2-106">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="c10c2-107">現時点でサポートされているイベントは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c10c2-107">The following events are currently supported.</span></span>

| <span data-ttu-id="c10c2-108">イベント</span><span class="sxs-lookup"><span data-stu-id="c10c2-108">Event</span></span> | <span data-ttu-id="c10c2-109">説明</span><span class="sxs-lookup"><span data-stu-id="c10c2-109">Description</span></span> | <span data-ttu-id="c10c2-110">サポートされているオブジェクト</span><span class="sxs-lookup"><span data-stu-id="c10c2-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="c10c2-111">オブジェクトが追加されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="c10c2-111">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="c10c2-112">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="c10c2-112">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onDeleted` | <span data-ttu-id="c10c2-113">オブジェクトが削除されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="c10c2-113">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="c10c2-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="c10c2-114">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onActivated` | <span data-ttu-id="c10c2-115">オブジェクトがアクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="c10c2-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="c10c2-116">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection)、 [**Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="c10c2-116">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onDeactivated` | <span data-ttu-id="c10c2-117">オブジェクトが非アクティブ化されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="c10c2-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="c10c2-118">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection)、 [**Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="c10c2-118">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="c10c2-119">セル内のデータが変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="c10c2-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="c10c2-120">[**worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet)、 [**Table**](https://dev.office.com/reference/add-ins/excel/table)、 [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span><span class="sxs-lookup"><span data-stu-id="c10c2-120">[**Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Table**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="c10c2-121">バインディングでデータまたは書式設定が変更されるときに発生します。</span><span class="sxs-lookup"><span data-stu-id="c10c2-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="c10c2-122">**Binding**</span><span class="sxs-lookup"><span data-stu-id="c10c2-122">**Binding**</span></span>](https://dev.office.com/reference/add-ins/excel/binding) |
| `onSelectionChanged` | <span data-ttu-id="c10c2-123">アクティブなセルまたは選択範囲が変更されたときに発生するイベント。</span><span class="sxs-lookup"><span data-stu-id="c10c2-123">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="c10c2-124">[ **worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet)、[**Table**](https://dev.office.com/reference/add-ins/excel/table)、[ **Binding**](https://dev.office.com/reference/add-ins/excel/binding)</span><span class="sxs-lookup"><span data-stu-id="c10c2-124">**Worksheet**, [Table](https://dev.office.com/reference/add-ins/excel/worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="c10c2-125">ドキュメント内の設定が変更されるときに発生します。</span><span class="sxs-lookup"><span data-stu-id="c10c2-125">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="c10c2-126">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="c10c2-126">**SettingCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="c10c2-127">イベント トリガー</span><span class="sxs-lookup"><span data-stu-id="c10c2-127">Event triggers</span></span>

<span data-ttu-id="c10c2-128">Excel ブックのイベントは、次の事項でトリガーできます。</span><span class="sxs-lookup"><span data-stu-id="c10c2-128">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="c10c2-129">ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作</span><span class="sxs-lookup"><span data-stu-id="c10c2-129">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="c10c2-130">ブックを変更する Office アドイン (JavaScript) コード</span><span class="sxs-lookup"><span data-stu-id="c10c2-130">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="c10c2-131">ブックを変更する VBA アドイン (マクロ) コード</span><span class="sxs-lookup"><span data-stu-id="c10c2-131">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="c10c2-132">Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="c10c2-132">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="c10c2-133">イベント ハンドラーのライフサイクル</span><span class="sxs-lookup"><span data-stu-id="c10c2-133">Lifecycle of an event handler</span></span>

<span data-ttu-id="c10c2-p102">イベント ハンドラーは、アドインでイベント ハンドラーを登録するときに作成され、アドインでイベント ハンドラーの登録を解除したとき、またはアドインが閉じられたときに破棄されます。イベント ハンドラーは、Excel ファイルの一部として保持されません。</span><span class="sxs-lookup"><span data-stu-id="c10c2-p102">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="c10c2-136">イベントと共同編集</span><span class="sxs-lookup"><span data-stu-id="c10c2-136">Events and coauthoring</span></span>

<span data-ttu-id="c10c2-p103">[共同編集機能](co-authoring-in-excel-add-ins.md)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。共同編集でトリガーできるイベント (`onChanged` など) の場合、対応する **Event** オブジェクトには **source** プロパティが含まれるようになります。このプロパティは、イベントが現在のユーザーによってローカルにトリガーされた (`event.source = Local`) ものか、リモートの共同作成者によってトリガーされた (`event.source = Remote`) ものかを示します。</span><span class="sxs-lookup"><span data-stu-id="c10c2-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="c10c2-139">イベント ハンドラーの登録</span><span class="sxs-lookup"><span data-stu-id="c10c2-139">Register an event handler</span></span>

<span data-ttu-id="c10c2-140">次のコード例では、ワークシートの `onChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録します。</span><span class="sxs-lookup"><span data-stu-id="c10c2-140">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="c10c2-141">このコードでは、そのワークシートでデータが変更されたときに、`handleDataChange` 関数を実行するように指定しています。</span><span class="sxs-lookup"><span data-stu-id="c10c2-141">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="c10c2-142">イベントの処理</span><span class="sxs-lookup"><span data-stu-id="c10c2-142">Handle an event</span></span>

<span data-ttu-id="c10c2-143">前の例で示したように、イベント ハンドラーの登録時には、特定のイベントが発生したときに実行する関数を指定します。</span><span class="sxs-lookup"><span data-stu-id="c10c2-143">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="c10c2-144">その関数は、目的のシナリオに必要なアクションを実行するように設計できます。</span><span class="sxs-lookup"><span data-stu-id="c10c2-144">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="c10c2-145">次のコード例は、イベントに関する情報を単にコンソールに出力するイベント ハンドラー関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="c10c2-145">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="c10c2-146">イベント ハンドラーを削除する</span><span class="sxs-lookup"><span data-stu-id="c10c2-146">Remove an event handler</span></span>

<span data-ttu-id="c10c2-147">次のコード例では、ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。</span><span class="sxs-lookup"><span data-stu-id="c10c2-147">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="c10c2-148">また、そのイベント ハンドラーを削除するために、後から呼び出すことができる `remove()` 関数も定義しています。</span><span class="sxs-lookup"><span data-stu-id="c10c2-148">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="c10c2-149">関連項目</span><span class="sxs-lookup"><span data-stu-id="c10c2-149">See also</span></span>

- [<span data-ttu-id="c10c2-150">Excel JavaScript API の中心概念</span><span class="sxs-lookup"><span data-stu-id="c10c2-150">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c10c2-151">Excel JavaScript API オープン仕様</span><span class="sxs-lookup"><span data-stu-id="c10c2-151">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)