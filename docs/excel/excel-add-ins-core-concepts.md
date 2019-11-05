---
title: Excel JavaScript API を使用した基本的なプログラミングの概念
description: Excel JavaScript API を使用して、Excel 用アドインをビルドします。
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: bd346764c3faba0cf3be7612c8b29dd5e0d4c28b
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924802"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="fdf83-103">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="fdf83-103">Fundamental programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="fdf83-104">この記事では、[Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) を使用して Excel 2016 以降のアドインをビルドする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-104">This article describes how to use the [Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="fdf83-105">ここでは API の使用の基本となる中心概念について説明し、広い範囲に対する読み取り、書き込み、一定範囲内すべてのセルの更新など、特定のタスクを実行するためのガイダンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-105">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="fdf83-106">Excel API の非同期性</span><span class="sxs-lookup"><span data-stu-id="fdf83-106">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="fdf83-p102">Web ベースの Excel アドインは、Windows 上の Office など、デスクトップ ベースのプラットフォーム上にある Office アプリケーションに組み込まれ、Office on the web の HTML iFrame 内で実行されるブラウザー コンテナー内で実行されます。サポートされているすべてのプラットフォームで Office.js API が Excel ホストと同期的に対話することは、パフォーマンスの観点からうまくいきません。このため、Office.js 内の **sync()** API の呼び出しにより [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) が返され、それは Excel アプリケーションが要求された読み取りまたは書き込み操作を完了したときに解決されます。また、操作ごとに別個の要求として送信する代わりに、プロパティの設定やメソッドの呼び出しなど、複数の操作をキューに登録し、**sync()** の 1 回の呼び出しでコマンドのバッチとしてそれらを実行することもできます。次のセクションでは、**Excel.run()** と **sync()** API を使用してこれを実行する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p102">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office on Windows and runs inside an HTML iFrame in Office on the web. Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations. Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions. Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action. The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>

## <a name="excelrun"></a><span data-ttu-id="fdf83-112">Excel.run</span><span class="sxs-lookup"><span data-stu-id="fdf83-112">Excel.run</span></span>

<span data-ttu-id="fdf83-p103">**Excel.run** は Excel オブジェクト モデルに対して実行する操作を指定した関数を実行します。 **Excel.run** は Excel オブジェクトと対話するために使用できる要求コンテキストを自動的に作成します。 **Excel.run**が完了すると、Promose が解決され、実行時に割り当てられたすべてのオブジェクトが自動的に解放されます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p103">**Excel.run** executes a function where you specify the actions to perform against the Excel object model. **Excel.run** automatically creates a request context that you can use to interact with Excel objects. When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="fdf83-p104">次の例は、**Excel.run** の使用方法を示しています。 Catch ステートメントは **Excel.run** 内で発生するエラーをキャッチし、ログに記録します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p104">The following example shows how to use **Excel.run**. The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>

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

### <a name="run-options"></a><span data-ttu-id="fdf83-118">実行オプション</span><span class="sxs-lookup"><span data-stu-id="fdf83-118">Run options</span></span>

<span data-ttu-id="fdf83-119">**Excel.run** には、[RunOptions](/javascript/api/excel/excel.runoptions) オブジェクトを使用するオーバーロードがあります。</span><span class="sxs-lookup"><span data-stu-id="fdf83-119">**Excel.run** has an overload that takes in a [RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="fdf83-120">これには、関数の実行時にプラットフォームの動作に影響を与えるプロパティのセットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="fdf83-120">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="fdf83-121">次のプロパティが現在サポートされています。</span><span class="sxs-lookup"><span data-stu-id="fdf83-121">The following property is currently supported:</span></span>

- <span data-ttu-id="fdf83-122">`delayForCellEdit`: ユーザーがセル編集モードを終了するまでバッチ要求を延期するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-122">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="fdf83-123">**true** の場合、バッチ要求は延期され、ユーザーがセル編集モードを終了した時点で実行されます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-123">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="fdf83-124">**false** の場合、バッチ要求は、ユーザーがセル編集モードにある場合、自動的に失敗します (ユーザーにエラーが表示されます)。</span><span class="sxs-lookup"><span data-stu-id="fdf83-124">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="fdf83-125">`delayForCellEdit` プロパティが指定されていない場合の既定の動作は、このプロパティが **false** の場合と同じ動作となります。</span><span class="sxs-lookup"><span data-stu-id="fdf83-125">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="request-context"></a><span data-ttu-id="fdf83-126">要求コンテキスト</span><span class="sxs-lookup"><span data-stu-id="fdf83-126">Request context</span></span>

<span data-ttu-id="fdf83-p107">Excel とユーザーのアドインは、2 つの異なるプロセスで実行されます。それらは異なるランタイム環境を使用するため、Excel アドインでは、ワークシート、範囲、グラフ、表など、Excel のオブジェクトにユーザーのアドインを接続するために **RequestContext** オブジェクトが必要です。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p107">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="fdf83-129">プロキシ オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fdf83-129">Proxy objects</span></span>

<span data-ttu-id="fdf83-p108">アドインで宣言し、使用する Excel JavaScript オブジェクトはプロキシ オブジェクトです。 起動するメソッドや、プロキシ オブジェクトに設定または読み込まれるプロパティは、保留中のコマンドのキューに単純に追加されます。 \*\*\*\* など、要求コンテキスト上で `context.sync()` メソッドを呼び出すと、キューに入れられたコマンドは Excel にディスパッチされて実行されます。 Excel の JavaScript API では、基本的にバッチを中心としています。 要求コンテキストに必要なだけ変更内容をキューに登録し、**sync()** メソッドを呼び出して、キューに入れられたコマンドをバッチで実行することができます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p108">The Excel JavaScript objects that you declare and use in an add-in are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run. The Excel JavaScript API is fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>

<span data-ttu-id="fdf83-p109">たとえば、次のコード スニペットでは、ローカル JavaScript オブジェクト **selectedRange** が Excel ドキュメント内の選択範囲を参照することを宣言し、そのオブジェクトでいくつかのプロパティを設定します。 **selectedRange** オブジェクトはプロキシ オブジェクトであるため、設定されたプロパティと、そのオブジェクトに対して呼び出されたメソッドは、ユーザーのアドインが **context.sync()** を呼び出すまで Excel ドキュメントには反映されません。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p109">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object. The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="sync"></a><span data-ttu-id="fdf83-137">sync()</span><span class="sxs-lookup"><span data-stu-id="fdf83-137">sync()</span></span>

<span data-ttu-id="fdf83-p110">要求コンテキストで **sync()** メソッドを呼び出すと、プロキシ オブジェクトと Excel ドキュメント内のオブジェクトの状態が同期されます。 **sync()** メソッドは、要求コンテキストのキューに登録されたすべてのコマンドを実行し、プロキシ オブジェクトに読み込まれるプロパティの値を取得します。 **sync()** メソッドは非同期で実行されて [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。これは、**sync()** メソッドが完了すると解決されます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p110">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document. The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects. The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>

<span data-ttu-id="fdf83-141">次の例は、ローカル JavaScript proxy オブジェクト (**selectedRange**) を定義し、そのオブジェクトのプロパティを読み込み、JavaScript の Promises パターンを使用して **context.sync()** を呼び出し、プロキシ オブジェクトと Excel ドキュメント内のオブジェクトの状態を同期するバッチ関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="fdf83-141">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>

```js
Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
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

<span data-ttu-id="fdf83-142">前の例では、**selectedRange** が設定され、**context.sync()** が呼び出されるとその **address** プロパティが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-142">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>

<span data-ttu-id="fdf83-143">**sync()** は Promise を返す非同期の操作であるため、常に Promise を (JavaScript で) **返す**必要があります。</span><span class="sxs-lookup"><span data-stu-id="fdf83-143">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript).</span></span> <span data-ttu-id="fdf83-144">このようにして、スクリプトの実行を継続する前に、**sync()** 操作を完了します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-144">Doing so ensures that the **sync()** operation completes before the script continues to run.</span></span> <span data-ttu-id="fdf83-145">**sync()** を用いたパフォーマンスの最適化の詳細については、「[Excel の JavaScript API を使用した、パフォーマンスの最適化](/office/dev/add-ins/excel/performance)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="fdf83-145">For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](/office/dev/add-ins/excel/performance).</span></span>

### <a name="load"></a><span data-ttu-id="fdf83-146">load()</span><span class="sxs-lookup"><span data-stu-id="fdf83-146">load()</span></span>

<span data-ttu-id="fdf83-p112">プロキシ オブジェクトのプロパティを読み取るには、まず Excel ドキュメントからプロキシ オブジェクトとデータを入力するプロパティを明示的に読み込み、それから **context.sync()** を呼び出す必要があります。 たとえば、選択範囲を参照するプロキシ オブジェクトを作成した後、選択範囲の **address** プロパティを読み取る必要がある場合、読み取る前に **address** プロパティを読み込む必要があります。 プロキシ オブジェクトのプロパティを読み込むよう要求するには、オブジェクトに対して **load()** メソッドを呼び出し、読み込むプロパティを指定します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p112">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it. To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="fdf83-p113">プロキシ オブジェクト上でメソッドを呼び出す、またはプロパティを設定するだけの場合は、**load()** メソッドを呼び出す必要はありません。 **load()** メソッドは、プロキシ オブジェクト上でプロパティを読み取る場合のみ必要です。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p113">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method. The **load()** method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="fdf83-p114">プロキシ オブジェクトに対してプロパティを設定、またはメソッドを呼び出す要求と同じように、プロキシ オブジェクトに対してプロパティを読み込む要求も、要求コンテキストで保留中のコマンドのキューに追加され、次回 **sync()** メソッドを呼び出すときに実行されます。**load()** の呼び出しは、必要なだけ要求コンテキストのキューに入れることができます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p114">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>

<span data-ttu-id="fdf83-154">次の例では、範囲の特定のプロパティのみが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-154">In the following example, only specific properties of the range are loaded.</span></span>

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

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

<span data-ttu-id="fdf83-155">前の例では、`format/font` が **myRange.load()** の呼び出しで指定されていないため、`format.font.color` プロパティは読み取れませんでした。</span><span class="sxs-lookup"><span data-stu-id="fdf83-155">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="fdf83-156">「[Excel の JavaScript API を使用した、パフォーマンスの最適化](performance.md)」の説明にあるとおり、パフォーマンスを最適化するため、オブジェクトに対して **load()** メソッドを使用するときに読み込むプロパティを明示的に指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fdf83-156">To optimize performance, you should explicitly specify the properties to load when using the **load()** method on an object, as covered in [Excel JavaScript API performance optimizations](performance.md).</span></span> <span data-ttu-id="fdf83-157">**load()** メソッドの詳細については、「[Excel JavaScript API を使用した高度なプログラミングの概念](excel-add-ins-advanced-concepts.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fdf83-157">For more information about the **load()** method, see [Advanced programming concepts with the Excel JavaScript API](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="fdf83-158">null または空白のプロパティ値</span><span class="sxs-lookup"><span data-stu-id="fdf83-158">null or blank property values</span></span>

### <a name="null-input-in-2-d-array"></a><span data-ttu-id="fdf83-159">2 次元配列での null の入力</span><span class="sxs-lookup"><span data-stu-id="fdf83-159">null input in 2-D Array</span></span>

<span data-ttu-id="fdf83-p116">Excel では、範囲は 2 次元配列で表され、最初のディメンションは行、2 番目のディメンションは列を示します。 範囲内の特定のセルだけに値、数値書式、または数式を設定するには、2 次元配列内のそのセルに値、数値書式、または数式を指定し、2 次元配列内のその他のすべてのセルに `null` を指定します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p116">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="fdf83-p117">たとえば、範囲内の 1 つのセルの数値書式を更新し、範囲内の他のセルすべての既存の数値書式を保持する場合、更新するセルに新しい数値書式を指定し、他のセルすべてに `null` を指定します。 次のコード スニペットでは、範囲内の 4 番目のセルに新しい数値書式を設定し、その前の 3 つのセルについては数値書式を変更せずに保持します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p117">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a><span data-ttu-id="fdf83-164">プロパティに対する null の入力</span><span class="sxs-lookup"><span data-stu-id="fdf83-164">null input for a property</span></span>

<span data-ttu-id="fdf83-p118">`null` は単一プロパティに有効な入力ではありません。たとえば、次のコード スニペットは、範囲の **values** プロパティを `null` に設定できないため無効です。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p118">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>

```js
range.values = null;
```

<span data-ttu-id="fdf83-167">同様に、次のコード スニペットは、`null` が **color** プロパティで有効ではないため無効です。</span><span class="sxs-lookup"><span data-stu-id="fdf83-167">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a><span data-ttu-id="fdf83-168">応答内の null プロパティ値</span><span class="sxs-lookup"><span data-stu-id="fdf83-168">null property values in the response</span></span>

<span data-ttu-id="fdf83-p119">指定の範囲に複数の値がある場合、`size` および `color` などの書式設定プロパティでは、応答に `null` 値が含まれます。 たとえば、範囲を取得してその `format.font.color` プロパティを読み込む場合:</span><span class="sxs-lookup"><span data-stu-id="fdf83-p119">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

- <span data-ttu-id="fdf83-171">範囲内のすべてのセルのフォントの色が同じ場合、`range.format.font.color` がその色を指定します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-171">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
- <span data-ttu-id="fdf83-172">範囲内に複数のフォントの色がある場合、`range.format.font.color` は `null` です。</span><span class="sxs-lookup"><span data-stu-id="fdf83-172">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

### <a name="blank-input-for-a-property"></a><span data-ttu-id="fdf83-173">プロパティに対する空白の入力</span><span class="sxs-lookup"><span data-stu-id="fdf83-173">Blank input for a property</span></span>

<span data-ttu-id="fdf83-p120">プロパティに空白の値 (`''` の間にスペースのない 2 つの引用符) を指定すると、プロパティをクリアまたはリセットする指示として解釈されます。例:</span><span class="sxs-lookup"><span data-stu-id="fdf83-p120">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

- <span data-ttu-id="fdf83-176">範囲の `values` プロパティに空白の値を指定すると、範囲のコンテンツはクリアされます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-176">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>

- <span data-ttu-id="fdf83-177">`numberFormat` プロパティに空白の値を指定すると、数値書式は `General` にリセットされます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-177">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>

- <span data-ttu-id="fdf83-178">`formula` プロパティと `formulaLocale` プロパティに空白の値を指定すると、数式の値はクリアされます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-178">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="fdf83-179">応答内の空白のプロパティ値</span><span class="sxs-lookup"><span data-stu-id="fdf83-179">Blank property values in the response</span></span>

<span data-ttu-id="fdf83-p121">読み取り操作では、応答内の空白のプロパティ値 (`''` の間にスペースのない、2 つの引用符) は、セルにデータまたは値がないことを示します。 次の 1 番目の例では、範囲内の最初と最後のセルにデータがありません。 2 番目の例では、範囲内の最初の 2 つのセルに数式がありません。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p121">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="fdf83-183">無制限の範囲への読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="fdf83-183">Read or write to an unbounded range</span></span>

### <a name="read-an-unbounded-range"></a><span data-ttu-id="fdf83-184">無制限の範囲の読み取り</span><span class="sxs-lookup"><span data-stu-id="fdf83-184">Read an unbounded range</span></span>

<span data-ttu-id="fdf83-p122">無制限の範囲のアドレスとは、列全体または行全体を指定する範囲のアドレスです。例:</span><span class="sxs-lookup"><span data-stu-id="fdf83-p122">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>

- <span data-ttu-id="fdf83-187">範囲のアドレスは、列全体で構成されます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-187">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
- <span data-ttu-id="fdf83-188">範囲のアドレスは、行全体で構成されます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-188">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>

<span data-ttu-id="fdf83-p123">API が無制限の範囲を取得する要求を行う場合 (`getRange('C:C')` など)、返される応答では、`null`、`values`、`text`、または `numberFormat` などのセル レベルのプロパティに `formula` 値が含まれます。 `address` または `cellCount` など、範囲のその他のプロパティには、無制限の範囲に有効な値が含まれます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p123">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>

### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="fdf83-191">無制限の範囲への書き込み</span><span class="sxs-lookup"><span data-stu-id="fdf83-191">Write to an unbounded range</span></span>

<span data-ttu-id="fdf83-p124">無制限の範囲では、入力要求が大きすぎるため、`values`、`numberFormat`、`formula` などのセル レベルのプロパティは設定できません。 たとえば、次のコード スニペットは、無制限の範囲に対して `values` を指定しようとしているため無効です。 無制限の範囲にセル レベルのプロパティを設定しようとすると、API からエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p124">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large. For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range. The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>

```js
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="fdf83-195">広い範囲に対する読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="fdf83-195">Read or write to a large range</span></span>

<span data-ttu-id="fdf83-p125">範囲に多数のセル、値、数値書式、数式などが含まれる場合、その範囲では API 操作を実行できない場合があります。 API は常に範囲に要求された操作 (特定のデータを取得または書き込む) を実行しようとしますが、広い範囲に対する読み取りや書き込みの操作は、過剰なリソース使用によるエラーになる場合があります。 このようなエラーを避けるため、広い範囲に対して読み取りや書き取り操作を 1 回で実行するのではなく、その範囲の小さいサブセットに対して個別に読み取りまたは書き込み操作を実行することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="fdf83-p125">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>

<span data-ttu-id="fdf83-199">システムの制限の詳細については、「[Excel の範囲の制限](../develop/common-coding-issues.md#excel-range-limits)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fdf83-199">For details on the system limitations, see [Excel Range limits](../develop/common-coding-issues.md#excel-range-limits).</span></span>

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="fdf83-200">範囲内のすべてのセルの更新</span><span class="sxs-lookup"><span data-stu-id="fdf83-200">Update all cells in a range</span></span>

<span data-ttu-id="fdf83-201">範囲内のすべてのセルに同じ更新 (すべてのセルに同じ値を入力する、同じ数値書式を設定する、同じ数式ですべてのセルにデータを入力するなど) を適用するには、**range** オブジェクトの該当するプロパティを必要な 1 つの値に設定します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-201">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>

<span data-ttu-id="fdf83-202">次の例では、20 個のセルを含む範囲を取得し、数値書式を設定してその範囲のすべてのセルに **3/11/2015** という値を設定します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-202">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = context.workbook.worksheets.getItem(sheetName);

    var range = worksheet.getRange(rangeAddress);
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

## <a name="handle-errors"></a><span data-ttu-id="fdf83-203">エラーを処理する</span><span class="sxs-lookup"><span data-stu-id="fdf83-203">Handle errors</span></span>

<span data-ttu-id="fdf83-204">API エラーが発生すると、API はコードとメッセージを含む **error** オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="fdf83-204">When an API error occurs, the API returns an **error** object that contains a code and a message.</span></span> <span data-ttu-id="fdf83-205">エラーの処理に関する詳細と、API エラーの一覧については、「[エラー処理](excel-add-ins-error-handling.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fdf83-205">For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="fdf83-206">関連項目</span><span class="sxs-lookup"><span data-stu-id="fdf83-206">See also</span></span>

- [<span data-ttu-id="fdf83-207">最初の Excel アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="fdf83-207">Build your first Excel add-in</span></span>](../quickstarts/excel-quickstart-jquery.md)
- [<span data-ttu-id="fdf83-208">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="fdf83-208">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [<span data-ttu-id="fdf83-209">Excel JavaScript API を使用した高度なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="fdf83-209">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="fdf83-210">Excel の JavaScript API を使用した、パフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="fdf83-210">Excel JavaScript API performance optimization</span></span>](/office/dev/add-ins/excel/performance)
- [<span data-ttu-id="fdf83-211">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="fdf83-211">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- <span data-ttu-id="fdf83-212">[一般的なコーディングの問題と、予期しないプラットフォームの動作](../develop/common-coding-issues.md)。</span><span class="sxs-lookup"><span data-stu-id="fdf83-212">[Common coding issues and unexpected platform behaviors](../develop/common-coding-issues.md).</span></span>
