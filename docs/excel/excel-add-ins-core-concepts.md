---
title: Excel JavaScript API の中心概念
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 37d652d2ad2f323d0f94583e530e91e775e06ddf
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925410"
---
# <a name="excel-javascript-api-core-concepts"></a><span data-ttu-id="e7dbb-102">Excel JavaScript API の中心概念</span><span class="sxs-lookup"><span data-stu-id="e7dbb-102">Excel JavaScript API core concepts</span></span>
 
<span data-ttu-id="e7dbb-103">この記事では、[Excel JavaScript API を使用して Excel 2016](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) のアドインをビルドする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-103">This article describes how to use the [Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) to build add-ins for Excel 2016.</span></span> <span data-ttu-id="e7dbb-104">ここでは API の使用の基本となる中心概念について説明し、広い範囲に対する読み取り、書き込み、一定範囲内すべてのセルの更新など、特定のタスクを実行するためのガイダンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-104">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="e7dbb-105">Excel API の非同期性</span><span class="sxs-lookup"><span data-stu-id="e7dbb-105">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="e7dbb-106">Web ベースの Excel アドインは、Office for Windows などのデスクトップ ベースのプラットフォーム上にある Office アプリケーションに組み込まれ、Office Online の HTML iFrame 内で実行されるブラウザー コンテナー内で実行されます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-106">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office for Windows and runs inside an HTML iFrame in Office Online.</span></span> <span data-ttu-id="e7dbb-107">サポートされているすべてのプラットフォームで Office.js API が Excel ホストと同期的に対話することは、パフォーマンスの観点からうまくいきません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-107">Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations.</span></span> <span data-ttu-id="e7dbb-108">このため、Office.js 内の **sync()** API の呼び出しにより [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) が返され、それは Excel アプリケーションが要求された読み取りまたは書き込み操作を完了したときに解決されます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-108">Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions.</span></span> <span data-ttu-id="e7dbb-109">また、操作ごとに別個の要求として送信する代わりに、プロパティの設定やメソッドの起動など、複数の操作をキューに登録し、**sync()** の 1 回の呼び出しでコマンドのバッチとしてそれらを実行することもできます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-109">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action.</span></span> <span data-ttu-id="e7dbb-110">次のセクションでは、**Excel.run()** と **sync()** API を使用してこれを実行する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-110">The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>
 
## <a name="excelrun"></a><span data-ttu-id="e7dbb-111">Excel.run</span><span class="sxs-lookup"><span data-stu-id="e7dbb-111">Excel.run</span></span>
 
<span data-ttu-id="e7dbb-112">**Excel.run** は Excel オブジェクト モデルに対して実行する操作を指定した関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-112">**Excel.run** executes a function where you specify the actions to perform against the Excel object model.</span></span> <span data-ttu-id="e7dbb-113">**Excel.run** は Excel オブジェクトと対話するために使用できる要求コンテキストを自動的に作成します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-113">**Excel.run** automatically creates a request context that you can use to interact with Excel objects.</span></span> <span data-ttu-id="e7dbb-114">**Excel.run**が完了すると、Promose が解決され、実行時に割り当てられたすべてのオブジェクトが自動的に解放されます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-114">When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>
 
<span data-ttu-id="e7dbb-115">次の例は、**Excel.run** の使用方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-115">The following example shows how to use **Excel.run**.</span></span> <span data-ttu-id="e7dbb-116">Catch ステートメントは **Excel.run** 内で発生するエラーをキャッチし、ログに記録します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-116">The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>
 
```js
Excel.run(function (context) {
  // You can use the Excel JavaScript API here in the batch function
  // to execute actions on the Excel object model.
  console.log('Your code goes here.');
}).catch(function (error) {
  console.log('error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```

## <a name="request-context"></a><span data-ttu-id="e7dbb-117">要求コンテキスト</span><span class="sxs-lookup"><span data-stu-id="e7dbb-117">Request context</span></span>
 
<span data-ttu-id="e7dbb-p105">Excel とユーザーのアドインは、2 つの異なるプロセスで実行されます。それらは異なるランタイム環境を使用するため、Excel アドインでは、ワークシート、範囲、グラフ、表など、Excel のオブジェクトにユーザーのアドインを接続するために **RequestContext** オブジェクトが必要です。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-p105">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>
 
## <a name="proxy-objects"></a><span data-ttu-id="e7dbb-120">プロキシ オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e7dbb-120">Proxy objects</span></span>
 
<span data-ttu-id="e7dbb-121">アドインで宣言し、使用する Excel JavaScript オブジェクトはプロキシ オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-121">The Excel JavaScript objects that you declare and use in an add-in are proxy objects.</span></span> <span data-ttu-id="e7dbb-122">起動するメソッドや、プロキシ オブジェクトに設定または読み込まれるプロパティは、保留中のコマンドのキューに単純に追加されます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-122">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="e7dbb-123">など、要求コンテキスト上で **sync()** メソッドを呼び出すと、キューに入れられたコマンドは Excel にディスパッチされて実行されます。`context.sync()`</span><span class="sxs-lookup"><span data-stu-id="e7dbb-123">When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run.</span></span> <span data-ttu-id="e7dbb-124">Excel の JavaScript API では、基本的にバッチを中心としています。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-124">The Excel JavaScript API is fundamentally batch-centric.</span></span> <span data-ttu-id="e7dbb-125">要求コンテキストに必要なだけ変更内容をキューに登録し、**sync()** メソッドを呼び出して、キューに入れられたコマンドをバッチで実行することができます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-125">You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>
 
<span data-ttu-id="e7dbb-126">たとえば、次のコード スニペットでは、ローカル JavaScript オブジェクト **selectedRange** が Excel ドキュメント内の選択範囲を参照することを宣言し、そのオブジェクトでいくつかのプロパティを設定します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-126">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object.</span></span> <span data-ttu-id="e7dbb-127">**selectedRange** オブジェクトはプロキシ オブジェクトであるため、設定されたプロパティと、そのオブジェクトに対して呼び出されたメソッドは、ユーザーのアドインが **context.sync()** を呼び出すまで Excel ドキュメントには反映されません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-127">The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a><span data-ttu-id="e7dbb-128">sync()</span><span class="sxs-lookup"><span data-stu-id="e7dbb-128">sync()</span></span>
 
<span data-ttu-id="e7dbb-129">要求コンテキストで **sync()** メソッドを呼び出すと、プロキシ オブジェクトと Excel ドキュメント内のオブジェクトの状態が同期されます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-129">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document.</span></span> <span data-ttu-id="e7dbb-130">**sync()** メソッドは、要求コンテキストのキューに登録されたすべてのコマンドを実行し、プロキシ オブジェクトに読み込まれるプロパティの値を取得します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-130">The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="e7dbb-131">**sync()** メソッドは非同期で実行され [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。これは、**sync()** メソッドが完了すると解決されます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-131">The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>
 
<span data-ttu-id="e7dbb-132">次の例は、ローカル JavaScript proxy オブジェクト (**selectedRange**) を定義し、そのオブジェクトのプロパティを読み込み、JavaScript の Promises パターンを使用して **context.sync()** を呼び出し、プロキシ オブジェクトと Excel ドキュメント内のオブジェクトの状態を同期するバッチ関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-132">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>
 
```js
Excel.run(function (context) {
  const selectedRange = context.workbook.getSelectedRange();
  selectedRange.load('address');
  return context.sync()
    .then(function () {
      console.log('The selected range is: ' + selectedRange.address);
  });
}).catch(function (error) {
  console.log('error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
<span data-ttu-id="e7dbb-133">前の例では、**selectedRange** が設定され、**context.sync()** が呼び出されるとその **address** プロパティが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-133">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>
 
<span data-ttu-id="e7dbb-134">**sync()** は Promise を返す非同期の操作であるため、常に Promise を (JavaScript で) **返す**必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-134">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript).</span></span> <span data-ttu-id="e7dbb-135">このようにして、スクリプトの実行を継続する前に、**sync()** 操作を完了します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-135">Doing so ensures that the **sync()** operation completes before the script continues to run.</span></span> <span data-ttu-id="e7dbb-136">**sync()** を用いたパフォーマンスの最適化の詳細については、「 [Excel JavaScript API のパフォーマンス最適化](https://docs.microsoft.com/office/dev/add-ins/excel/performance)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-136">For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span></span>
 
### <a name="load"></a><span data-ttu-id="e7dbb-137">load()</span><span class="sxs-lookup"><span data-stu-id="e7dbb-137">load()</span></span>
 
<span data-ttu-id="e7dbb-138">プロキシ オブジェクトのプロパティを読み取るには、まず Excel ドキュメントからプロキシ オブジェクトとデータを入力するプロパティを明示的に読み込み、それから **context.sync()** を呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-138">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**.</span></span> <span data-ttu-id="e7dbb-139">たとえば、選択範囲を参照するプロキシ オブジェクトを作成した後、選択範囲の **address** プロパティを読み取る必要がある場合、読み取る前に **address** プロパティを読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-139">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it.</span></span> <span data-ttu-id="e7dbb-140">プロキシ オブジェクトのプロパティを読み込むよう要求するには、オブジェクトに対して **load()** メソッドを呼び出し、読み込むプロパティを指定します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-140">To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="e7dbb-141">プロキシ オブジェクト上でメソッドを呼び出す、またはプロパティを設定するだけの場合は、**load()** メソッドを呼び出す必要はありません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-141">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method.</span></span> <span data-ttu-id="e7dbb-142">**load()** メソッドは、プロキシ オブジェクト上でプロパティを読み取る場合のみ必要です。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-142">The **load()** method is only required when you want to read properties on a proxy object.</span></span>
 
<span data-ttu-id="e7dbb-p112">プロキシ オブジェクトに対してプロパティを設定、またはメソッドを呼び出す要求と同じように、プロキシ オブジェクトに対してプロパティを読み込む要求も、要求コンテキストで保留中のコマンドのキューに追加され、次回 **sync()** メソッドを呼び出すときに実行されます。**load()** の呼び出しは、必要なだけ要求コンテキストのキューに入れることができます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-p112">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>
 
<span data-ttu-id="e7dbb-145">次の例では、範囲の特定のプロパティのみが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-145">In the following example, only specific properties of the range are loaded.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:B2';
  const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
 
  myRange.load(['address', 'format/*', 'format/fill', 'entireRow' ]);
 
  return context.sync()
    .then(function () {
      console.log (myRange.address);              // ok
      console.log (myRange.format.wrapText);      // ok
      console.log (myRange.format.fill.color);    // ok
      //console.log (myRange.format.font.color);  // not ok as it was not loaded
  });
}).then(function () {
  console.log('done');
}).catch(function (error) {
  console.log('Error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
<span data-ttu-id="e7dbb-146">前の例では、`format/font` が **myRange.load()** の呼び出しで指定されていないため、`format.font.color` プロパティは読み取れませんでした。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-146">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="e7dbb-147">パフォーマンスを最適化するには、「 **Excel JavaScript API のパフォーマンス最適化**」にあるように、[load()](performance.md) メソッドをオブジェクトに使用する際、プロパティとリレーションシップを明示的に指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-147">To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object.</span></span> <span data-ttu-id="e7dbb-148">**load()** メソッドの詳細は、「[Excel JavaScript API の高度な概念](excel-add-ins-advanced-concepts.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-148">For more information about the **load()** method, see [Excel JavaScript API advanced concepts](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="e7dbb-149">null または空白のプロパティ値</span><span class="sxs-lookup"><span data-stu-id="e7dbb-149">null or blank property values</span></span>
 
### <a name="null-input-in-2-d-array"></a><span data-ttu-id="e7dbb-150">2 次元配列での null の入力</span><span class="sxs-lookup"><span data-stu-id="e7dbb-150">null input in 2-D Array</span></span>
 
<span data-ttu-id="e7dbb-151">Excel では、範囲は 2 次元配列で表され、最初のディメンションは行、2 番目のディメンションは列を示します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-151">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns.</span></span> <span data-ttu-id="e7dbb-152">範囲内の特定のセルだけに値、数値書式、または数式を設定するには、2 次元配列内のそのセルに値、数値書式、または数式を指定し、2 次元配列内のその他のすべてのセルに `null` を指定します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-152">To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>
 
<span data-ttu-id="e7dbb-153">たとえば、範囲内の 1 つのセルの数値書式を更新し、範囲内の他のセルすべての既存の数値書式を保持する場合、更新するセルに新しい数値書式を指定し、他のセルすべてに `null` を指定します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-153">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells.</span></span> <span data-ttu-id="e7dbb-154">次のコード スニペットでは、範囲内の 4 番目のセルに新しい数値書式を設定し、その前の 3 つのセルについては数値書式を変更せずに保持します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-154">The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a><span data-ttu-id="e7dbb-155">プロパティに対する null の入力</span><span class="sxs-lookup"><span data-stu-id="e7dbb-155">null input for a property</span></span>
 
<span data-ttu-id="e7dbb-p116">`null` は単一プロパティに有効な入力ではありません。たとえば、次のコード スニペットは、範囲の **values** プロパティを `null` に設定できないため無効です。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-p116">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>
 
```js
range.values = null;
```
 
<span data-ttu-id="e7dbb-158">同様に、次のコード スニペットは、`null` が **color** プロパティで有効ではないため無効です。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-158">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a><span data-ttu-id="e7dbb-159">応答内の null プロパティ値</span><span class="sxs-lookup"><span data-stu-id="e7dbb-159">null property values in the response</span></span>
 
<span data-ttu-id="e7dbb-160">指定の範囲に複数の値がある場合、`size` および `color` などの書式設定プロパティでは、応答に `null` 値が含まれます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-160">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range.</span></span> <span data-ttu-id="e7dbb-161">たとえば、範囲を取得してその `format.font.color` プロパティを読み込む場合:</span><span class="sxs-lookup"><span data-stu-id="e7dbb-161">For example, if you retrieve a range and load its `format.font.color` property:</span></span>
 
* <span data-ttu-id="e7dbb-162">範囲内のすべてのセルのフォントの色が同じ場合、`range.format.font.color` がその色を指定します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-162">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="e7dbb-163">範囲内に複数のフォントの色がある場合、`range.format.font.color` は `null` です。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-163">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>
 
### <a name="blank-input-for-a-property"></a><span data-ttu-id="e7dbb-164">プロパティに対する空白の入力</span><span class="sxs-lookup"><span data-stu-id="e7dbb-164">Blank input for a property</span></span>
 
<span data-ttu-id="e7dbb-p118">プロパティに空白の値 (`''` の間にスペースのない 2 つの引用符) を指定すると、プロパティをクリアまたはリセットする指示として解釈されます。例:</span><span class="sxs-lookup"><span data-stu-id="e7dbb-p118">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>
 
* <span data-ttu-id="e7dbb-167">範囲の `values` プロパティに空白の値を指定すると、範囲のコンテンツはクリアされます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-167">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
 
* <span data-ttu-id="e7dbb-168">プロパティに空白の値を指定すると、数値書式は `General` にリセットされます。`numberFormat`</span><span class="sxs-lookup"><span data-stu-id="e7dbb-168">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
 
* <span data-ttu-id="e7dbb-169">プロパティと `formulaLocale` プロパティに空白の値を指定すると、数式の値はクリアされます。`formula`</span><span class="sxs-lookup"><span data-stu-id="e7dbb-169">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>
 
### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="e7dbb-170">応答内の空白のプロパティ値</span><span class="sxs-lookup"><span data-stu-id="e7dbb-170">Blank property values in the response</span></span>
 
<span data-ttu-id="e7dbb-171">読み取り操作では、応答内の空白のプロパティ値 (`''` の間にスペースのない、2 つの引用符) は、セルにデータまたは値がないことを示します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-171">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value.</span></span> <span data-ttu-id="e7dbb-172">次の 1 番目の例では、範囲内の最初と最後のセルにデータがありません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-172">In the first example below, the first and last cell in the range contain no data.</span></span> <span data-ttu-id="e7dbb-173">2 番目の例では、範囲内の最初の 2 つのセルに数式がありません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-173">In the second example, the first two cells in the range do not contain a formula.</span></span>
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="e7dbb-174">無制限の範囲への読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="e7dbb-174">Read or write to an unbounded range</span></span>
 
### <a name="read-an-unbounded-range"></a><span data-ttu-id="e7dbb-175">無制限の範囲の読み取り</span><span class="sxs-lookup"><span data-stu-id="e7dbb-175">Read an unbounded range</span></span>
 
<span data-ttu-id="e7dbb-p120">無制限の範囲のアドレスとは、列全体または行全体を指定する範囲のアドレスです。例:</span><span class="sxs-lookup"><span data-stu-id="e7dbb-p120">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>
 
* <span data-ttu-id="e7dbb-178">範囲のアドレスは、列全体で構成されます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-178">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
* <span data-ttu-id="e7dbb-179">範囲のアドレスは、行全体で構成されます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-179">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>
 
<span data-ttu-id="e7dbb-180">API が無制限の範囲を取得する要求を行う場合 (`getRange('C:C')` など)、返される応答では、`values`、`text`、`numberFormat`、または `formula` などのセル レベルのプロパティに `null` 値が含まれます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-180">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`.</span></span> <span data-ttu-id="e7dbb-181">または `cellCount` など、範囲のその他のプロパティには、無制限の範囲に有効な値が含まれます。`address`</span><span class="sxs-lookup"><span data-stu-id="e7dbb-181">Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>
 
### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="e7dbb-182">無制限の範囲への書き込み</span><span class="sxs-lookup"><span data-stu-id="e7dbb-182">Write to an unbounded range</span></span>
 
<span data-ttu-id="e7dbb-183">無制限の範囲では、入力要求が大きすぎるため、`values`、`numberFormat`、`formula` などのセル レベルのプロパティは設定できません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-183">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large.</span></span> <span data-ttu-id="e7dbb-184">たとえば、次のコード スニペットは、無制限の範囲に対して `values` を指定しようとしているため無効です。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-184">For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="e7dbb-185">無制限の範囲にセル レベルのプロパティを設定しようとすると、API からエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-185">The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="e7dbb-186">広い範囲に対する読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="e7dbb-186">Read or write to a large range</span></span>
 
<span data-ttu-id="e7dbb-187">範囲に多数のセル、値、数値書式、数式などが含まれる場合、その範囲では API 操作を実行できない場合があります。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-187">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range.</span></span> <span data-ttu-id="e7dbb-188">API は常に範囲に要求された操作 (特定のデータを取得または書き込む) を実行しようとしますが、広い範囲に対する読み取りや書き込みの操作は、過剰なリソース使用によるエラーになる場合があります。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-188">The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization.</span></span> <span data-ttu-id="e7dbb-189">このようなエラーを避けるため、広い範囲に対して読み取りや書き取り操作を 1 回で実行するのではなく、その範囲の小さいサブセットに対して個別に読み取りまたは書き込み操作を実行することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-189">To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>
 
## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="e7dbb-190">範囲内のすべてのセルの更新</span><span class="sxs-lookup"><span data-stu-id="e7dbb-190">Update all cells in a range</span></span>
 
<span data-ttu-id="e7dbb-191">範囲内のすべてのセルに同じ更新 (すべてのセルに同じ値を入力する、同じ数値書式を設定する、同じ数式ですべてのセルにデータを入力するなど) を適用するには、**range** オブジェクトの該当するプロパティを必要な 1 つの値に設定します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-191">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>
 
<span data-ttu-id="e7dbb-192">次の例では、20 個のセルを含む範囲を取得し、数値書式を設定してその範囲のすべてのセルに **3/11/2015** という値を設定します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-192">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:A20';
  const worksheet = context.workbook.worksheets.getItem(sheetName);
 
  const range = worksheet.getRange(rangeAddress);
  range.numberFormat = 'm/d/yyyy';
  range.values = '3/11/2015';
  range.load('text');
 
  return context.sync()
    .then(function () {
      console.log(range.text);
  });
}).catch(function (error) {
  console.log('Error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
## <a name="error-messages"></a><span data-ttu-id="e7dbb-193">エラー メッセージ</span><span class="sxs-lookup"><span data-stu-id="e7dbb-193">Error messages</span></span>
 
<span data-ttu-id="e7dbb-194">API エラーが発生すると、API ではコードとメッセージを含む **error** オブジェクトが返されます。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-194">When an API error occurs, the API will return an **error** object that contains a code and a message.</span></span> <span data-ttu-id="e7dbb-195">次の表は、API から返されるエラー一覧の定義を示します。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-195">The following table defines a list of errors that the API may return.</span></span>
 
|<span data-ttu-id="e7dbb-196">error.code</span><span class="sxs-lookup"><span data-stu-id="e7dbb-196">error.code</span></span> | <span data-ttu-id="e7dbb-197">error.message</span><span class="sxs-lookup"><span data-stu-id="e7dbb-197">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="e7dbb-198">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="e7dbb-198">InvalidArgument</span></span> |<span data-ttu-id="e7dbb-199">引数が無効であるか、存在しません。または形式が正しくありません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-199">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="e7dbb-200">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="e7dbb-200">InvalidRequest</span></span>  |<span data-ttu-id="e7dbb-201">要求を処理できません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-201">Cannot process the request.</span></span>|
|<span data-ttu-id="e7dbb-202">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="e7dbb-202">InvalidReference</span></span>|<span data-ttu-id="e7dbb-203">この参照は、現在の操作に対して無効です。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-203">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="e7dbb-204">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="e7dbb-204">InvalidBinding</span></span>  |<span data-ttu-id="e7dbb-205">このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-205">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="e7dbb-206">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="e7dbb-206">InvalidSelection</span></span>|<span data-ttu-id="e7dbb-207">現在の選択内容は、この操作では無効です。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-207">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="e7dbb-208">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="e7dbb-208">Unauthenticated</span></span> |<span data-ttu-id="e7dbb-209">必要な認証情報が見つからないか、無効です。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-209">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="e7dbb-210">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="e7dbb-210">AccessDenied</span></span> |<span data-ttu-id="e7dbb-211">要求された操作を実行できません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-211">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="e7dbb-212">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="e7dbb-212">ItemNotFound</span></span> |<span data-ttu-id="e7dbb-213">要求されたリソースは存在しません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-213">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="e7dbb-214">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="e7dbb-214">ActivityLimitReached</span></span>|<span data-ttu-id="e7dbb-215">アクティビティの制限に達しました。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-215">Activity limit has been reached.</span></span>|
|<span data-ttu-id="e7dbb-216">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e7dbb-216">GeneralException</span></span>|<span data-ttu-id="e7dbb-217">要求の処理中に内部エラーが発生しました。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-217">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="e7dbb-218">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="e7dbb-218">NotImplemented</span></span>  |<span data-ttu-id="e7dbb-219">要求された機能は実装されていません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-219">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="e7dbb-220">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="e7dbb-220">ServiceNotAvailable</span></span>|<span data-ttu-id="e7dbb-221">サービスを利用できません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-221">The service is unavailable.</span></span>|
|<span data-ttu-id="e7dbb-222">Conflict</span><span class="sxs-lookup"><span data-stu-id="e7dbb-222">Conflict</span></span>              |<span data-ttu-id="e7dbb-223">競合のため、要求を処理できませんでした。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-223">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="e7dbb-224">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="e7dbb-224">ItemAlreadyExists</span></span>|<span data-ttu-id="e7dbb-225">作成中のリソースはすでに存在しています。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-225">The resource being created already exists.</span></span>|
|<span data-ttu-id="e7dbb-226">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="e7dbb-226">UnsupportedOperation</span></span>|<span data-ttu-id="e7dbb-227">試行中の操作はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-227">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="e7dbb-228">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="e7dbb-228">RequestAborted</span></span>|<span data-ttu-id="e7dbb-229">実行時に要求が中止されました。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-229">The request was aborted during run time.</span></span>|
|<span data-ttu-id="e7dbb-230">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="e7dbb-230">ApiNotAvailable</span></span>|<span data-ttu-id="e7dbb-231">要求された API は使用できません。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-231">The requested API is not available.</span></span>|
|<span data-ttu-id="e7dbb-232">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="e7dbb-232">InsertDeleteConflict</span></span>|<span data-ttu-id="e7dbb-233">試行された挿入操作または削除操作で競合が発生しました。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-233">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="e7dbb-234">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="e7dbb-234">InvalidOperation</span></span>|<span data-ttu-id="e7dbb-235">試行された操作は、このオブジェクトでは無効です。</span><span class="sxs-lookup"><span data-stu-id="e7dbb-235">The operation attempted is invalid on the object.</span></span>|
 
## <a name="see-also"></a><span data-ttu-id="e7dbb-236">関連項目</span><span class="sxs-lookup"><span data-stu-id="e7dbb-236">See also</span></span>
 
* [<span data-ttu-id="e7dbb-237">Excel アドインを使う</span><span class="sxs-lookup"><span data-stu-id="e7dbb-237">Get started with Excel add-ins</span></span>](excel-add-ins-get-started-overview.md)
* [<span data-ttu-id="e7dbb-238">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="e7dbb-238">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
* [<span data-ttu-id="e7dbb-239">Excel JavaScript API パフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="e7dbb-239">Excel JavaScript API performance optimization</span></span>](https://docs.microsoft.com/office/dev/add-ins/excel/performance)
* [<span data-ttu-id="e7dbb-240">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="e7dbb-240">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
