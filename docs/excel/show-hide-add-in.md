---
title: 共有ランタイムで Office アドインを表示または非表示にする
description: 連続して実行している間にプログラムによってアドインの UI を表示または非表示にする方法について説明します。
ms.date: 03/02/2020
localization_priority: Normal
ms.openlocfilehash: c028823be165723cad3c0b314b53fe7e618188b2
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/04/2020
ms.locfileid: "42413796"
---
# <a name="show-or-hide-an-office-add-in-in-a-shared-runtime-preview"></a><span data-ttu-id="fb3c1-103">共有ランタイムで Office アドインを表示または非表示にする (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="fb3c1-103">Show or hide an Office Add-in in a shared runtime (preview)</span></span>

<span data-ttu-id="fb3c1-104">Office アドインには、次のいずれかの部分を含めることができます。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-104">An Office Add-in can include any of the following parts:</span></span>

- <span data-ttu-id="fb3c1-105">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="fb3c1-105">A task pane</span></span>
- <span data-ttu-id="fb3c1-106">UI レス関数ファイル</span><span class="sxs-lookup"><span data-stu-id="fb3c1-106">A UI-less function file</span></span>
- <span data-ttu-id="fb3c1-107">Excel カスタム関数</span><span class="sxs-lookup"><span data-stu-id="fb3c1-107">An Excel custom function</span></span>

<span data-ttu-id="fb3c1-108">既定では、各パーツは独自の独立した JavaScript ランタイムで実行され、独自のグローバルオブジェクトとグローバル変数を持ちます。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-108">By default, each part runs in its own separate JavaScript runtime, with its own global object and global variables.</span></span> 

<span data-ttu-id="fb3c1-109">2つ以上のパーツを含むアドインは、共通の JavaScript ランタイムを共有できます。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-109">It's possible for add-ins with two or more parts to share a common JavaScript runtime.</span></span> <span data-ttu-id="fb3c1-110">この共有ランタイム機能により、アドインの実行中に作業ウィンドウを非表示にしたり、再び開くことができる新しいプレビュー Api を有効にすることができます。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-110">This shared runtime feature enables new preview APIs that hide and reopen the task pane while the add-in runs.</span></span>

> [!INCLUDE [Information about using preview APIs](../includes/excel-shared-runtime-preview-note.md)]

## <a name="configure-an-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="fb3c1-111">共有ランタイムを使用するようにアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="fb3c1-111">Configure an add-in to use a shared runtime</span></span>

<span data-ttu-id="fb3c1-112">共有ランタイムを使用するようにアドインを構成するには、「[共有ランタイムを使用するように Office アドインを構成](configure-your-add-in-to-use-a-shared-runtime.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-112">To configure the add-in to use a shared runtime, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="show-and-hide-the-task-pane"></a><span data-ttu-id="fb3c1-113">作業ウィンドウを表示または非表示にする</span><span class="sxs-lookup"><span data-stu-id="fb3c1-113">Show and hide the task pane</span></span>

<span data-ttu-id="fb3c1-114">新しい Api は`Office.addin`プロパティにあります。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-114">The new APIs are in the `Office.addin` property.</span></span> <span data-ttu-id="fb3c1-115">作業ウィンドウを表示するには、コード`Office.addin.showAsTaskpane()`を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-115">To show the task pane, your code calls `Office.addin.showAsTaskpane()`.</span></span> <span data-ttu-id="fb3c1-116">Office は、作業ウィンドウのリソース ID (`resid`) に割り当てたページを作業ウィンドウに表示します。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-116">Office will display in a task pane the page that you assigned to the resource ID (`resid`) for the task pane.</span></span> <span data-ttu-id="fb3c1-117">これは、 `resid`マニフェスト`<Action xsi:type="ShowTaskpane">`内のにに`<SourceLocation>`割り当てられたです。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-117">This is the `resid` that you assigned to the `<SourceLocation>` of the `<Action xsi:type="ShowTaskpane">` in the manifest.</span></span> <span data-ttu-id="fb3c1-118">(「[共有ランタイムを使用するために Office アドインを構成する」を](configure-your-add-in-to-use-a-shared-runtime.md)参照してください)。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-118">(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).)</span></span>

<span data-ttu-id="fb3c1-119">これは非同期メソッドなので、完了するまで後続のコードが実行されないように、コードで待機する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-119">This is an asynchronous method, so your code should await it when the subsequent code should not run until it completes.</span></span> <span data-ttu-id="fb3c1-120">この完了は、使用して`await`いる JavaScript 構文`then()`に応じて、キーワードまたはメソッドのいずれかを使用して待機します。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-120">Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.</span></span> <span data-ttu-id="fb3c1-121">次の例では、 **CurrentQuarterSales**という名前の Excel ワークシートが存在することを前提としています。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-121">The following assumes that there is an Excel worksheet named **CurrentQuarterSales**.</span></span> <span data-ttu-id="fb3c1-122">このワークシートがアクティブになると、アドインによって作業ウィンドウが表示されるようになります。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-122">The add-in should make the task pane visible whenever this worksheet is activated.</span></span> <span data-ttu-id="fb3c1-123">このメソッド`onCurrentQuarter`は、ワークシートに登録されている、 [onactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated)イベントのハンドラーです。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-123">The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated) event which has been registered for the worksheet.</span></span>

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

<span data-ttu-id="fb3c1-124">作業ウィンドウを非表示にするには`Office.addin.hide()`、コードを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-124">To hide the task pane, your code calls `Office.addin.hide()`.</span></span> <span data-ttu-id="fb3c1-125">次の例は、 [onDeactivated アクティブ](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated)化イベントに登録されているハンドラーです。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-125">The following example is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated) event.</span></span>

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

### <a name="preservation-of-state-and-event-listeners"></a><span data-ttu-id="fb3c1-126">状態およびイベントリスナーの保持</span><span class="sxs-lookup"><span data-stu-id="fb3c1-126">Preservation of state and event listeners</span></span>

<span data-ttu-id="fb3c1-127">メソッド`hide()`と`showAsTaskpane()`メソッドは、作業ウィンドウの*表示状態*のみを変更します。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-127">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span></span> <span data-ttu-id="fb3c1-128">アンロードまたは再ロードしたり、その状態を再初期化したりすることはありません。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-128">They do not unload or reload it (or reinitialize its state).</span></span>

<span data-ttu-id="fb3c1-129">次のシナリオを考えてみます。作業ウィンドウは、タブで設計されています。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-129">Consider the following scenario: A task pane is designed with tabs.</span></span> <span data-ttu-id="fb3c1-130">[**ホーム**] タブは、アドインを最初に起動したときに開かれています。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-130">The **Home** tab is open when the add-in is first launched.</span></span> <span data-ttu-id="fb3c1-131">ユーザーが [**設定**] タブを開き、後で、あるイベントに応答し`hide()`て、作業ウィンドウの呼び出しのコードを開くとします。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-131">Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event.</span></span> <span data-ttu-id="fb3c1-132">他のイベントに`showAsTaskpane()`応答して、後でコードを呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-132">Still later code calls `showAsTaskpane()` in response to another event.</span></span> <span data-ttu-id="fb3c1-133">作業ウィンドウが再度表示され、[**設定**] タブが選択されたままになります。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-133">The task pane will reappear, and the **Settings** tab is still selected.</span></span>

![[ホーム]、[設定]、[お気に入り]、および [アカウント] というラベルの付いた4つのタブがある作業ウィンドウのスクリーンショット。](../images/TaskpaneWithTabs.png)

<span data-ttu-id="fb3c1-135">さらに、作業ウィンドウに登録されているイベントリスナーは、作業ウィンドウが非表示になっている場合でも、引き続き実行されます。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-135">In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.</span></span>

<span data-ttu-id="fb3c1-136">次のシナリオを考えます。この作業ウィンドウには、 **Sheet1**と`Worksheet.onActivated`いう`Worksheet.onDeactivated`シートの Excel およびイベントのハンドラーが登録されています。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-136">Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**.</span></span> <span data-ttu-id="fb3c1-137">アクティブ化されたハンドラーによって、作業ウィンドウに緑の点が表示されます。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-137">The activated handler causes a green dot to appear in the task pane.</span></span> <span data-ttu-id="fb3c1-138">非アクティブ化されたハンドラーは、ドット red (これは既定の状態) をオフにします。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-138">The deactivated handler turns the dot red (which is its default state).</span></span> <span data-ttu-id="fb3c1-139">Sheet1 がアクティブ化さ`hide()`れ\*\*\*\* ておらず、ドットが赤の場合は、コードが呼び出されるとします。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-139">Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red.</span></span> <span data-ttu-id="fb3c1-140">作業ウィンドウは非表示になっていますが、 **Sheet1**がアクティブになります。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-140">While the task pane is hidden, **Sheet1** is activated.</span></span> <span data-ttu-id="fb3c1-141">イベントに応答`showAsTaskpane()`して、後でコードを呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-141">Later code calls `showAsTaskpane()` in response to some event.</span></span> <span data-ttu-id="fb3c1-142">作業ウィンドウが開くと、その作業ウィンドウが非表示になっているにもかかわらず、イベントリスナーとハンドラーが実行されるため、ドットは緑になります。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-142">When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.</span></span>

### <a name="handle-visibility-changed-event"></a><span data-ttu-id="fb3c1-143">可視性の変更イベントを処理する</span><span class="sxs-lookup"><span data-stu-id="fb3c1-143">Handle visibility changed event</span></span>

<span data-ttu-id="fb3c1-144">コードによって作業ウィンドウの表示がまたは`showAsTaskpane()` `hide()`に変更されると`VisibilityModeChanged` 、Office によってイベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-144">When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event.</span></span> <span data-ttu-id="fb3c1-145">このイベントを処理すると便利な場合があります。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-145">It can be useful to handle this event.</span></span> <span data-ttu-id="fb3c1-146">たとえば、作業ウィンドウにブック内のすべてのシートの一覧が表示されているとします。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-146">For example, suppose the task pane displays a list of all the sheets in a workbook.</span></span> <span data-ttu-id="fb3c1-147">作業ウィンドウが非表示になっているときに新しいワークシートが追加されても、その作業ウィンドウが表示されないようにするには、リストに新しいワークシート名を追加します。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-147">If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list.</span></span> <span data-ttu-id="fb3c1-148">しかし、以下のコード例に`VisibilityModeChanged`示されているように、コードでイベントに応答して、 [Worksheet.name](/javascript/api/excel/excel.worksheet#name)コレクション内のすべてのワークシートのプロパティを再読み込みすることが[できます。](/javascript/api/excel/excel.workbook#worksheets)</span><span class="sxs-lookup"><span data-stu-id="fb3c1-148">But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.</span></span>

<span data-ttu-id="fb3c1-149">イベントのハンドラーを登録するには、ほとんどの Office JavaScript コンテキストでの "add handler" メソッドは使用しません。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-149">To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts.</span></span> <span data-ttu-id="fb3c1-150">代わりに、ハンドラーを渡すための特殊な関数[onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-)が用意されています。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-150">Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span></span> <span data-ttu-id="fb3c1-151">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-151">The following is an example.</span></span> <span data-ttu-id="fb3c1-152">このプロパティの`args.visibilityMode`型は[VisibilityMode](/javascript/api/office/office.visibilitymode)であることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-152">Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span></span>

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

<span data-ttu-id="fb3c1-153">この関数は、ハンドラーを*deregisters*する別の関数を返します。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-153">The function returns another function that *deregisters* the handler.</span></span> <span data-ttu-id="fb3c1-154">この例は、単純ですが、堅牢ではありません。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-154">Here is a simple, but not robust, example:</span></span>

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

<span data-ttu-id="fb3c1-155">この`onVisibilityModeChanged`メソッドは非同期です。つまり、コードから返される\*\* `onVisibilityModeChanged`登録解除ハンドラーを呼び出す場合は、 `onVisibilityModeChanged`登録解除ハンドラーを呼び出す前に、が完了していることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-155">The `onVisibilityModeChanged` method is asynchronous which means that if your code calls the *deregister* handler that `onVisibilityModeChanged` returns, you should ensure that `onVisibilityModeChanged` has completed before calling the deregister handler.</span></span> <span data-ttu-id="fb3c1-156">そのための1つの方法は、 `await`次の例のように、メソッド呼び出しでキーワードを使用することです。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-156">One way to do that is to use the `await` keyword on the method call as in the following example.</span></span>

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

<span data-ttu-id="fb3c1-157">ES2015 の JavaScript のみを使用する場合は、次の例に示すよう`then`に、コードでメソッドを使用して、返された Promise オブジェクトが解決されるまで待機し、返された関数をグローバル変数に代入することができます。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-157">If you want to use only pre-ES2015 JavaScript, your code can use the `then` method to wait until the returned Promise object has resolved and assign the returned function to a global variable as in the following example.</span></span>

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

<span data-ttu-id="fb3c1-158">登録解除関数は、それ自体が非同期です。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-158">The deregister function is itself asynchronous.</span></span> <span data-ttu-id="fb3c1-159">そのため、登録解除の完了後に実行してはならないコードがある場合は、次の例に示すように`await` 、登録解除関数`then`をキーワードまたはメソッドで待機する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-159">So, if you have code that should not run until after the deregistration is complete, then the deregister function should also be awaited with either the `await` keyword or with a `then` method as in the following examples.</span></span>

<span data-ttu-id="fb3c1-160">ハンドラーを登録解除するには、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="fb3c1-160">To deregister the handler:</span></span>

```javascript
await removeVisibilityModeHandler();
// subsequent code here

// or use pre-ES2015 syntax:
removeVisibilityModeHandler().then(function () {
        // subsequent code here
    })
```
