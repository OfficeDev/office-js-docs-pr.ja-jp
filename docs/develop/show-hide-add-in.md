---
title: アドインの作業ウィンドウを表示Office非表示にする
description: アドインが継続的に実行されている間に、プログラムによってアドインのユーザー インターフェイスを非表示または表示する方法について説明します。
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 20db609a3a6ded5624391f705dab1ad6b8f6e043
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789251"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a><span data-ttu-id="24275-103">アドインの作業ウィンドウを表示Office非表示にする</span><span class="sxs-lookup"><span data-stu-id="24275-103">Show or hide the task pane of your Office Add-in</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="24275-104">関数を呼び出すことによって、Officeアドインの作業ウィンドウを表示 `Office.addin.showAsTaskpane()` できます。</span><span class="sxs-lookup"><span data-stu-id="24275-104">You can show the task pane of your Office Add-in by calling the `Office.addin.showAsTaskpane()` function.</span></span>

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

<span data-ttu-id="24275-105">前のコードでは **、CurrentQuarterSales** という名前の Excel ワークシートがあるシナリオを想定しています。</span><span class="sxs-lookup"><span data-stu-id="24275-105">The previous code assumes a scenario where there is an Excel worksheet named **CurrentQuarterSales**.</span></span> <span data-ttu-id="24275-106">このワークシートがアクティブ化されるたびに、アドインによって作業ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="24275-106">The add-in will make the task pane visible whenever this worksheet is activated.</span></span> <span data-ttu-id="24275-107">このメソッド `onCurrentQuarter` は、ワークシートに登録 [されている Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) イベントのハンドラーです。</span><span class="sxs-lookup"><span data-stu-id="24275-107">The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) event which has been registered for the worksheet.</span></span>

<span data-ttu-id="24275-108">関数を呼び出して作業ウィンドウを非表示 `Office.addin.hide()` にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="24275-108">You can also hide the task pane by calling the `Office.addin.hide()` function.</span></span>

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

<span data-ttu-id="24275-109">前のコードは [、Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated) イベントに登録されているハンドラーです。</span><span class="sxs-lookup"><span data-stu-id="24275-109">The previous code is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated) event.</span></span>

## <a name="additional-details-on-showing-the-task-pane"></a><span data-ttu-id="24275-110">作業ウィンドウの表示に関するその他の詳細</span><span class="sxs-lookup"><span data-stu-id="24275-110">Additional details on showing the task pane</span></span>

<span data-ttu-id="24275-111">呼び出しOffice作業ウィンドウのリソース ID ( ) 値として割り当てたファイルが作業ウィンドウ `Office.addin.showAsTaskpane()` `resid` に表示されます。</span><span class="sxs-lookup"><span data-stu-id="24275-111">When you call `Office.addin.showAsTaskpane()`, Office will display in a task pane the file that you assigned as the resource ID (`resid`) value of the task pane.</span></span> <span data-ttu-id="24275-112">この `resid` 値は、ファイルを開き、要素内 **manifest.xml** することで、割り当 `<SourceLocation>` てまたは変更 `<Action xsi:type="ShowTaskpane">` できます。</span><span class="sxs-lookup"><span data-stu-id="24275-112">This `resid` value can be assigned or changed by opening your **manifest.xml** file and locating `<SourceLocation>` inside the `<Action xsi:type="ShowTaskpane">` element.</span></span>
<span data-ttu-id="24275-113">(詳 [しくは、「共有ランタイムOffice使うアドイン](configure-your-add-in-to-use-a-shared-runtime.md) の構成」をご覧ください)。</span><span class="sxs-lookup"><span data-stu-id="24275-113">(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md) for additional details.)</span></span>

<span data-ttu-id="24275-114">非同期 `Office.addin.showAsTaskpane()` メソッドの場合、コードは関数が完了するまで実行を続行します。</span><span class="sxs-lookup"><span data-stu-id="24275-114">Since `Office.addin.showAsTaskpane()` is an asynchronous method, your code will continue running until the function is complete.</span></span> <span data-ttu-id="24275-115">使用している JavaScript 構文に応じて、キーワードまたはメソッドを使用してこの完了 `await` `then()` を待ちます。</span><span class="sxs-lookup"><span data-stu-id="24275-115">Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.</span></span>

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a><span data-ttu-id="24275-116">共有ランタイムを使用するアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="24275-116">Configure your add-in to use the shared runtime</span></span>

<span data-ttu-id="24275-117">これらのメソッドを `showAsTaskpane()` `hide()` 使用するには、アドインで共有ランタイムを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="24275-117">To use the `showAsTaskpane()` and `hide()` methods, your add-in must use the shared runtime.</span></span> <span data-ttu-id="24275-118">詳細については、「共有ランタイム [を使用Officeアドインを構成する」を参照してください](configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="24275-118">For more information, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="preservation-of-state-and-event-listeners"></a><span data-ttu-id="24275-119">状態リスナーとイベント リスナーの保持</span><span class="sxs-lookup"><span data-stu-id="24275-119">Preservation of state and event listeners</span></span>

<span data-ttu-id="24275-120">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span><span class="sxs-lookup"><span data-stu-id="24275-120">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span></span> <span data-ttu-id="24275-121">アンロードまたは再読み込み (または状態の再初期化) は行されません。</span><span class="sxs-lookup"><span data-stu-id="24275-121">They do not unload or reload it (or reinitialize its state).</span></span>

<span data-ttu-id="24275-122">次のシナリオを考えます。作業ウィンドウはタブで設計されています。</span><span class="sxs-lookup"><span data-stu-id="24275-122">Consider the following scenario: A task pane is designed with tabs.</span></span> <span data-ttu-id="24275-123">アドイン **が** 最初に起動すると、[ホーム] タブが開きます。</span><span class="sxs-lookup"><span data-stu-id="24275-123">The **Home** tab is open when the add-in is first launched.</span></span> <span data-ttu-id="24275-124">ユーザーが [設定] タブを開き、後で何らかのイベントに応答して作業ウィンドウの呼び出し `hide()` をコード化したとします。</span><span class="sxs-lookup"><span data-stu-id="24275-124">Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event.</span></span> <span data-ttu-id="24275-125">別のイベントへの応答 `showAsTaskpane()` として、後でコードが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="24275-125">Still later code calls `showAsTaskpane()` in response to another event.</span></span> <span data-ttu-id="24275-126">作業ウィンドウが再び表示され、[設定] **タブは** 引き続き選択されます。</span><span class="sxs-lookup"><span data-stu-id="24275-126">The task pane will reappear, and the **Settings** tab is still selected.</span></span>

![[ホーム]、[設定]、[お気に入り]、および [アカウント] というラベルの付いた 4 つのタブがある作業ウィンドウのスクリーンショット。](../images/TaskpaneWithTabs.png)

<span data-ttu-id="24275-128">また、作業ウィンドウに登録されているイベント リスナーは、作業ウィンドウが非表示の場合でも引き続き実行されます。</span><span class="sxs-lookup"><span data-stu-id="24275-128">In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.</span></span>

<span data-ttu-id="24275-129">次のシナリオについて考えます。作業ウィンドウには、Excel の登録されたハンドラーと `Worksheet.onActivated` `Worksheet.onDeactivated` **、Sheet1** という名前のシートのイベントがあります。</span><span class="sxs-lookup"><span data-stu-id="24275-129">Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**.</span></span> <span data-ttu-id="24275-130">アクティブ化されたハンドラーによって、作業ウィンドウに緑色のドットが表示されます。</span><span class="sxs-lookup"><span data-stu-id="24275-130">The activated handler causes a green dot to appear in the task pane.</span></span> <span data-ttu-id="24275-131">非アクティブ化されたハンドラーは、ドットを赤色 (既定の状態) に変更します。</span><span class="sxs-lookup"><span data-stu-id="24275-131">The deactivated handler turns the dot red (which is its default state).</span></span> <span data-ttu-id="24275-132">次に、シート `hide()` **1** がアクティブ化されていないときに、ドットが赤のときにコードが呼び出されるとします。</span><span class="sxs-lookup"><span data-stu-id="24275-132">Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red.</span></span> <span data-ttu-id="24275-133">作業ウィンドウが非表示の間、 **シート 1 が** アクティブになります。</span><span class="sxs-lookup"><span data-stu-id="24275-133">While the task pane is hidden, **Sheet1** is activated.</span></span> <span data-ttu-id="24275-134">以降のコードは、 `showAsTaskpane()` 何らかのイベントに応答して呼び出します。</span><span class="sxs-lookup"><span data-stu-id="24275-134">Later code calls `showAsTaskpane()` in response to some event.</span></span> <span data-ttu-id="24275-135">作業ウィンドウが開くと、作業ウィンドウが非表示でもイベント リスナーとハンドラーが実行されたため、ドットは緑色になります。</span><span class="sxs-lookup"><span data-stu-id="24275-135">When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.</span></span>

## <a name="handle-the-visibility-changed-event"></a><span data-ttu-id="24275-136">可視性の変更イベントを処理する</span><span class="sxs-lookup"><span data-stu-id="24275-136">Handle the visibility changed event</span></span>

<span data-ttu-id="24275-137">コードで作業ウィンドウの表示/非表示を変更すると、イベントOffice `showAsTaskpane()` `hide()` トリガー `VisibilityModeChanged` されます。</span><span class="sxs-lookup"><span data-stu-id="24275-137">When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event.</span></span> <span data-ttu-id="24275-138">このイベントを処理すると便利です。</span><span class="sxs-lookup"><span data-stu-id="24275-138">It can be useful to handle this event.</span></span> <span data-ttu-id="24275-139">たとえば、作業ウィンドウにブック内のすべてのシートの一覧が表示されたとします。</span><span class="sxs-lookup"><span data-stu-id="24275-139">For example, suppose the task pane displays a list of all the sheets in a workbook.</span></span> <span data-ttu-id="24275-140">作業ウィンドウが非表示の間に新しいワークシートが追加された場合、作業ウィンドウを表示すると、それ自体は新しいワークシート名をリストに追加しません。</span><span class="sxs-lookup"><span data-stu-id="24275-140">If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list.</span></span> <span data-ttu-id="24275-141">ただし、次のコード例に示すように、コードはイベントに応答して `VisibilityModeChanged` [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets)コレクション内のすべてのワークシートの[Worksheet.name](/javascript/api/excel/excel.worksheet#name)プロパティを再読み込みできます。</span><span class="sxs-lookup"><span data-stu-id="24275-141">But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.</span></span>

<span data-ttu-id="24275-142">イベントのハンドラーを登録するには、JavaScript コンテキストのほとんどの場合のように、"ハンドラーの追加" メソッドOffice使用します。</span><span class="sxs-lookup"><span data-stu-id="24275-142">To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts.</span></span> <span data-ttu-id="24275-143">代わりに、ハンドラーを渡す特別な関数 [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-)があります。</span><span class="sxs-lookup"><span data-stu-id="24275-143">Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span></span> <span data-ttu-id="24275-144">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="24275-144">The following is an example.</span></span> <span data-ttu-id="24275-145">プロパティは `args.visibilityMode` [VisibilityMode 型です](/javascript/api/office/office.visibilitymode)。</span><span class="sxs-lookup"><span data-stu-id="24275-145">Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span></span>

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

<span data-ttu-id="24275-146">この関数は、ハンドラーを登録解除 *する別の関数を* 返します。</span><span class="sxs-lookup"><span data-stu-id="24275-146">The function returns another function that *deregisters* the handler.</span></span> <span data-ttu-id="24275-147">次に、シンプルですが堅牢ではない例を示します。</span><span class="sxs-lookup"><span data-stu-id="24275-147">Here is a simple, but not robust, example:</span></span>

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="24275-148">メソッドは非同期であり、promise を返します。つまり、登録解除ハンドラーを呼び出す前に、コードで promise のフルフィルメントを待 `onVisibilityModeChanged` **つ必要** があります。</span><span class="sxs-lookup"><span data-stu-id="24275-148">The `onVisibilityModeChanged` method is asynchronous and returns a promise, which means that your code needs to await the fulfillment of the promise before it can call the **deregister** handler.</span></span>

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

<span data-ttu-id="24275-149">登録解除関数も非同期で、promise を返します。</span><span class="sxs-lookup"><span data-stu-id="24275-149">The deregister function is also asynchronous and returns a promise.</span></span> <span data-ttu-id="24275-150">したがって、登録解除が完了するまで実行しないコードがある場合は、登録解除関数によって返される promise を待つ必要があります。</span><span class="sxs-lookup"><span data-stu-id="24275-150">So, if you have code that should not run until after the deregistration is complete, then you should await the promise returned by the deregister function.</span></span>

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a><span data-ttu-id="24275-151">関連項目</span><span class="sxs-lookup"><span data-stu-id="24275-151">See also</span></span>

- [<span data-ttu-id="24275-152">共有 JavaScript Office使用する新しいアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="24275-152">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="24275-153">ドキュメントが開Officeアドインでコードを実行する</span><span class="sxs-lookup"><span data-stu-id="24275-153">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
