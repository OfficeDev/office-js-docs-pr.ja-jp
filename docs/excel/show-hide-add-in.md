---
title: 共有ランタイムで Office アドインを表示または非表示にする
description: 連続して実行している間にプログラムでアドインのユーザーインターフェイスを表示または非表示にする方法について説明します。
ms.date: 05/17/2020
localization_priority: Normal
ms.openlocfilehash: e49c47c86a986c85ad12e09666b7ac2fb5411322
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275715"
---
# <a name="show-or-hide-an-office-add-in-in-a-shared-runtime"></a><span data-ttu-id="d2b3b-103">共有ランタイムで Office アドインを表示または非表示にする</span><span class="sxs-lookup"><span data-stu-id="d2b3b-103">Show or hide an Office Add-in in a shared runtime</span></span>

<span data-ttu-id="d2b3b-104">Office アドインには、次のいずれかの部分を含めることができます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-104">An Office Add-in can include any of the following parts:</span></span>

- <span data-ttu-id="d2b3b-105">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="d2b3b-105">A task pane</span></span>
- <span data-ttu-id="d2b3b-106">UI レス関数ファイル (作業ウィンドウや他のユーザーインターフェイス要素を使用しないカスタム関数)</span><span class="sxs-lookup"><span data-stu-id="d2b3b-106">A UI-less function file (custom functions which do not use a task pane or other user interface elements)</span></span>
- <span data-ttu-id="d2b3b-107">Excel カスタム関数</span><span class="sxs-lookup"><span data-stu-id="d2b3b-107">An Excel custom function</span></span>

<span data-ttu-id="d2b3b-108">既定では、各パーツは独自の独立した JavaScript ランタイムで実行され、独自のグローバルオブジェクトとグローバル変数を持ちます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-108">By default, each part runs in its own separate JavaScript runtime, with its own global object and global variables.</span></span>

<span data-ttu-id="d2b3b-109">2つ以上のパーツを含むアドインは、共通の JavaScript ランタイムを共有できます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-109">It's possible for add-ins with two or more parts to share a common JavaScript runtime.</span></span> <span data-ttu-id="d2b3b-110">この共有ランタイム機能を使用すると、アドインの実行中に作業ウィンドウを非表示にし、再び開くことができる新しい Api が有効になります。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-110">This shared runtime feature enables new APIs that hide and reopen the task pane while the add-in runs.</span></span>

## <a name="configure-an-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="d2b3b-111">共有ランタイムを使用するようにアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="d2b3b-111">Configure an add-in to use a shared runtime</span></span>

<span data-ttu-id="d2b3b-112">共有ランタイムを使用するようにアドインを構成するには、「[共有ランタイムを使用するように Office アドインを構成](configure-your-add-in-to-use-a-shared-runtime.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-112">To configure the add-in to use a shared runtime, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="show-and-hide-the-task-pane"></a><span data-ttu-id="d2b3b-113">作業ウィンドウを表示または非表示にする</span><span class="sxs-lookup"><span data-stu-id="d2b3b-113">Show and hide the task pane</span></span>

<span data-ttu-id="d2b3b-114">新しい Api は `Office.addin` プロパティにあります。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-114">The new APIs are in the `Office.addin` property.</span></span> <span data-ttu-id="d2b3b-115">作業ウィンドウを表示するには、コードを呼び出し `Office.addin.showAsTaskpane()` ます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-115">To show the task pane, your code calls `Office.addin.showAsTaskpane()`.</span></span> <span data-ttu-id="d2b3b-116">Office は、作業ウィンドウのリソース ID () に割り当てたページを作業ウィンドウに表示し `resid` ます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-116">Office will display in a task pane the page that you assigned to the resource ID (`resid`) for the task pane.</span></span> <span data-ttu-id="d2b3b-117">これは、 `resid` マニフェスト内のにに割り当てられたです `<SourceLocation>` `<Action xsi:type="ShowTaskpane">` 。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-117">This is the `resid` that you assigned to the `<SourceLocation>` of the `<Action xsi:type="ShowTaskpane">` in the manifest.</span></span> <span data-ttu-id="d2b3b-118">(「[共有ランタイムを使用するために Office アドインを構成する」を](configure-your-add-in-to-use-a-shared-runtime.md)参照してください)。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-118">(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).)</span></span>

<span data-ttu-id="d2b3b-119">これは非同期メソッドなので、完了するまで後続のコードが実行されないように、コードで待機する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-119">This is an asynchronous method, so your code should await it when the subsequent code should not run until it completes.</span></span> <span data-ttu-id="d2b3b-120">この完了は `await` `then()` 、使用している JavaScript 構文に応じて、キーワードまたはメソッドのいずれかを使用して待機します。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-120">Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.</span></span> <span data-ttu-id="d2b3b-121">次の例では、 **CurrentQuarterSales**という名前の Excel ワークシートが存在することを前提としています。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-121">The following assumes that there is an Excel worksheet named **CurrentQuarterSales**.</span></span> <span data-ttu-id="d2b3b-122">このワークシートがアクティブになると、アドインによって作業ウィンドウが表示されるようになります。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-122">The add-in should make the task pane visible whenever this worksheet is activated.</span></span> <span data-ttu-id="d2b3b-123">このメソッド `onCurrentQuarter` は、ワークシートに登録されている、 [onactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated)イベントのハンドラーです。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-123">The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated) event which has been registered for the worksheet.</span></span>

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

<span data-ttu-id="d2b3b-124">作業ウィンドウを非表示にするには、コードを呼び出し `Office.addin.hide()` ます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-124">To hide the task pane, your code calls `Office.addin.hide()`.</span></span> <span data-ttu-id="d2b3b-125">次の例は、 [onDeactivated アクティブ](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated)化イベントに登録されているハンドラーです。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-125">The following example is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated) event.</span></span>

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

### <a name="preservation-of-state-and-event-listeners"></a><span data-ttu-id="d2b3b-126">状態およびイベントリスナーの保持</span><span class="sxs-lookup"><span data-stu-id="d2b3b-126">Preservation of state and event listeners</span></span>

<span data-ttu-id="d2b3b-127">`hide()`メソッドと `showAsTaskpane()` メソッドは、作業ウィンドウの*表示状態*のみを変更します。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-127">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span></span> <span data-ttu-id="d2b3b-128">アンロードまたは再ロードしたり、その状態を再初期化したりすることはありません。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-128">They do not unload or reload it (or reinitialize its state).</span></span>

<span data-ttu-id="d2b3b-129">次のシナリオを考えてみます。作業ウィンドウは、タブで設計されています。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-129">Consider the following scenario: A task pane is designed with tabs.</span></span> <span data-ttu-id="d2b3b-130">[**ホーム**] タブは、アドインを最初に起動したときに開かれています。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-130">The **Home** tab is open when the add-in is first launched.</span></span> <span data-ttu-id="d2b3b-131">ユーザーが [**設定**] タブを開き、後で、あるイベントに応答して、作業ウィンドウの呼び出しのコードを開くとし `hide()` ます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-131">Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event.</span></span> <span data-ttu-id="d2b3b-132">他のイベントに応答して、後でコードを呼び出すこと `showAsTaskpane()` ができます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-132">Still later code calls `showAsTaskpane()` in response to another event.</span></span> <span data-ttu-id="d2b3b-133">作業ウィンドウが再度表示され、[**設定**] タブが選択されたままになります。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-133">The task pane will reappear, and the **Settings** tab is still selected.</span></span>

![[ホーム]、[設定]、[お気に入り]、および [アカウント] というラベルの付いた4つのタブがある作業ウィンドウのスクリーンショット。](../images/TaskpaneWithTabs.png)

<span data-ttu-id="d2b3b-135">さらに、作業ウィンドウに登録されているイベントリスナーは、作業ウィンドウが非表示になっている場合でも、引き続き実行されます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-135">In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.</span></span>

<span data-ttu-id="d2b3b-136">次のシナリオを考えます。この作業ウィンドウには、 `Worksheet.onActivated` Sheet1 というシートの Excel およびイベントのハンドラーが登録されてい `Worksheet.onDeactivated` ます。 **Sheet1**</span><span class="sxs-lookup"><span data-stu-id="d2b3b-136">Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**.</span></span> <span data-ttu-id="d2b3b-137">アクティブ化されたハンドラーによって、作業ウィンドウに緑の点が表示されます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-137">The activated handler causes a green dot to appear in the task pane.</span></span> <span data-ttu-id="d2b3b-138">非アクティブ化されたハンドラーは、ドット red (これは既定の状態) をオフにします。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-138">The deactivated handler turns the dot red (which is its default state).</span></span> <span data-ttu-id="d2b3b-139">`hide()` **Sheet1**がアクティブ化されておらず、ドットが赤の場合は、コードが呼び出されるとします。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-139">Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red.</span></span> <span data-ttu-id="d2b3b-140">作業ウィンドウは非表示になっていますが、 **Sheet1**がアクティブになります。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-140">While the task pane is hidden, **Sheet1** is activated.</span></span> <span data-ttu-id="d2b3b-141">`showAsTaskpane()`イベントに応答して、後でコードを呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-141">Later code calls `showAsTaskpane()` in response to some event.</span></span> <span data-ttu-id="d2b3b-142">作業ウィンドウが開くと、その作業ウィンドウが非表示になっているにもかかわらず、イベントリスナーとハンドラーが実行されるため、ドットは緑になります。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-142">When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.</span></span>

### <a name="handle-visibility-changed-event"></a><span data-ttu-id="d2b3b-143">可視性の変更イベントを処理する</span><span class="sxs-lookup"><span data-stu-id="d2b3b-143">Handle visibility changed event</span></span>

<span data-ttu-id="d2b3b-144">コードによって作業ウィンドウの表示がまたはに変更されると、Office によって `showAsTaskpane()` `hide()` イベントがトリガーさ `VisibilityModeChanged` れます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-144">When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event.</span></span> <span data-ttu-id="d2b3b-145">このイベントを処理すると便利な場合があります。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-145">It can be useful to handle this event.</span></span> <span data-ttu-id="d2b3b-146">たとえば、作業ウィンドウにブック内のすべてのシートの一覧が表示されているとします。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-146">For example, suppose the task pane displays a list of all the sheets in a workbook.</span></span> <span data-ttu-id="d2b3b-147">作業ウィンドウが非表示になっているときに新しいワークシートが追加されても、その作業ウィンドウが表示されないようにするには、リストに新しいワークシート名を追加します。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-147">If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list.</span></span> <span data-ttu-id="d2b3b-148">しかし、 `VisibilityModeChanged` 以下のコード例に示されているように、コードでイベントに応答して、 [Worksheet.name](/javascript/api/excel/excel.worksheet#name)コレクション内のすべてのワークシートのプロパティを再読み込みすることが[できます。](/javascript/api/excel/excel.workbook#worksheets)</span><span class="sxs-lookup"><span data-stu-id="d2b3b-148">But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.</span></span>

<span data-ttu-id="d2b3b-149">イベントのハンドラーを登録するには、ほとんどの Office JavaScript コンテキストでの "add handler" メソッドは使用しません。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-149">To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts.</span></span> <span data-ttu-id="d2b3b-150">代わりに、ハンドラーを渡すための特殊な関数[onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-)が用意されています。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-150">Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span></span> <span data-ttu-id="d2b3b-151">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-151">The following is an example.</span></span> <span data-ttu-id="d2b3b-152">このプロパティの `args.visibilityMode` 型は[VisibilityMode](/javascript/api/office/office.visibilitymode)であることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-152">Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span></span>

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

<span data-ttu-id="d2b3b-153">この関数は、ハンドラーを*deregisters*する別の関数を返します。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-153">The function returns another function that *deregisters* the handler.</span></span> <span data-ttu-id="d2b3b-154">この例は、単純ですが、堅牢ではありません。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-154">Here is a simple, but not robust, example:</span></span>

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="d2b3b-155">この `onVisibilityModeChanged` メソッドは非同期です。つまり、コードから返される*登録*解除ハンドラーを呼び出す場合は、登録解除ハンドラーを呼び出す前に、が完了していることを `onVisibilityModeChanged` 確認する必要があり `onVisibilityModeChanged` ます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-155">The `onVisibilityModeChanged` method is asynchronous which means that if your code calls the *deregister* handler that `onVisibilityModeChanged` returns, you should ensure that `onVisibilityModeChanged` has completed before calling the deregister handler.</span></span> <span data-ttu-id="d2b3b-156">そのための1つの方法は、 `await` 次の例のように、メソッド呼び出しでキーワードを使用することです。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-156">One way to do that is to use the `await` keyword on the method call as in the following example.</span></span>

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

<span data-ttu-id="d2b3b-157">ES2015 の JavaScript のみを使用する場合は、次の例に示すように、コードでメソッドを使用して、 `then` 返された Promise オブジェクトが解決されるまで待機し、返された関数をグローバル変数に代入することができます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-157">If you want to use only pre-ES2015 JavaScript, your code can use the `then` method to wait until the returned Promise object has resolved and assign the returned function to a global variable as in the following example.</span></span>

```javascript
var removeVisibilityModeHandler;

Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
}).then(function(removeHandler) {
        removeVisibilityModeHandler = removeHandler;
    });

// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="d2b3b-158">登録解除関数は、それ自体が非同期です。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-158">The deregister function is itself asynchronous.</span></span> <span data-ttu-id="d2b3b-159">そのため、登録解除の完了後に実行してはならないコードがある場合は、次の例に示すように、登録解除関数を `await` キーワードまたはメソッドで待機する必要があり `then` ます。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-159">So, if you have code that should not run until after the deregistration is complete, then the deregister function should also be awaited with either the `await` keyword or with a `then` method as in the following examples.</span></span>

<span data-ttu-id="d2b3b-160">ハンドラーを登録解除するには、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="d2b3b-160">To deregister the handler:</span></span>

```javascript
await removeVisibilityModeHandler();
// subsequent code here

// or use pre-ES2015 syntax:
removeVisibilityModeHandler().then(function () {
        // subsequent code here
    })
```
