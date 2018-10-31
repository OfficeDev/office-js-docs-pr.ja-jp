---
title: Excel JavaScript API を使用した基本的なプログラミングの概念
description: Excel JavaScript APIを使用して、Excel 用アドインを構築します。
ms.date: 10/03/2018
ms.openlocfilehash: f93ec7b5e34f90f2d61f29d861b7e0c19f66f6e3
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505987"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="d83d9-103">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="d83d9-103">Fundamental programming concepts with the Excel JavaScript API</span></span>
 
<span data-ttu-id="d83d9-p101">この資料では、 [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) を使用してアドイン を、Excel 2016 またはそれ以降にビルドする方法について説明します。API を使用する上で基本となる、読み取りまたは書き込み、広い範囲、範囲、およびその他のすべてのセルを更新するなどの特定のタスクを実行するためのガイダンスを提供する主要な概念が導入されています。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p101">This article describes how to use the [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) to build add-ins for Excel 2016 or later. It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="d83d9-106">Excel API の非同期性</span><span class="sxs-lookup"><span data-stu-id="d83d9-106">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="d83d9-p102">Office for Windows などのデスクトップ ベースのプラットフォーム上の Office アプリケーション内で埋め込まれ、Office オンラインの HTML iFrame 内で動作するブラウザーのコンテナー内で実行して、web を使用した Excel のアドインを使用します。Office.js の API がサポートされているすべてのプラットフォーム間で、Excel ホストと同期的に対話を有効にすることは、パフォーマンスに関する考慮事項のため不可能です。したがって、Office.jsの **sync()** API 呼び出しは、Excel アプリケーションが要求された読み取りまたは書き込みアクションを完了したときに解決される [約束](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。また、操作ごとに別の要求を送信するのではなく、プロパティを設定またはメソッドを呼び出すなど、複数の操作のキューして、 **sync()** への 1 回の呼び出しで指定されたコマンドのバッチとして実行できます。次のセクションでは、 **Excel.run()** と **sync()** API を使用してこれを実現する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p102">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office for Windows and runs inside an HTML iFrame in Office Online. Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations. Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions. Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action. The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>
 
## <a name="excelrun"></a><span data-ttu-id="d83d9-112">Excel.run</span><span class="sxs-lookup"><span data-stu-id="d83d9-112">Excel.run</span></span>
 
<span data-ttu-id="d83d9-p103">**Excel.run** では、Excel オブジェクト モデルに対して実行するアクションを指定する関数を実行します。 **Excel.run** は、自動的に Excel のオブジェクトと対話するために使用できる要求のコンテキストを作成します。 **Excel.run** が完了して、約束を解決すると、実行時に割り当てられたすべてのオブジェクトが自動的に解放されます。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p103">**Excel.run** executes a function where you specify the actions to perform against the Excel object model. **Excel.run** automatically creates a request context that you can use to interact with Excel objects. When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>
 
<span data-ttu-id="d83d9-p104">**Excel.run**を使用する例を次に示します。Catch ステートメントでは、 **Excel.run**で発生したエラーキャッチし、ログに記録します。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p104">The following example shows how to use **Excel.run**. The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>
 
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

## <a name="request-context"></a><span data-ttu-id="d83d9-118">要求コンテキスト</span><span class="sxs-lookup"><span data-stu-id="d83d9-118">Request context</span></span>
 
<span data-ttu-id="d83d9-p105">Excel とユーザーのアドインは、2 つの異なるプロセスで実行されます。それらは異なるランタイム環境を使用するため、Excel アドインでは、ワークシート、範囲、グラフ、表など、Excel のオブジェクトにユーザーのアドインを接続するために **RequestContext** オブジェクトが必要です。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p105">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>
 
## <a name="proxy-objects"></a><span data-ttu-id="d83d9-121">プロキシ オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d83d9-121">Proxy objects</span></span>
 
<span data-ttu-id="d83d9-p106">宣言して、アドインで使用する Excel の JavaScript オブジェクトは、プロキシ オブジェクトです。起動するメソッドまたはプロパティを設定するか、プロキシ オブジェクトにロードするだけで、保留中のコマンドのキューに追加されます。要求のコンテキストに **sync()** メソッドを呼び出すとき (たとえば、 `context.sync()`)、キュー内のコマンドを Excel にディスパッチし、実行します。Excel JavaScript API では、基本的にバッチを中心としました。要求のコンテキストにメソッドを呼び出して、 **sync()** をキューに登録されたコマンドのバッチを実行すると、多くの変更をキューにできます。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p106">The Excel JavaScript objects that you declare and use in an add-in are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run. The Excel JavaScript API is fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>
 
<span data-ttu-id="d83d9-p107">たとえば、次のコード スニペットでは、ローカル JavaScript オブジェクト **selectedRange** が Excel ドキュメント内の選択範囲を参照することを宣言し、そのオブジェクトでいくつかのプロパティを設定します。**selectedRange** オブジェクトはプロキシ オブジェクトであるため、設定されたプロパティと、そのオブジェクトに対して呼び出されたメソッドは、ユーザーのアドインが**context.sync()** を呼び出すまで Excel ドキュメントには反映されません。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p107">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object. The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a><span data-ttu-id="d83d9-129">sync()</span><span class="sxs-lookup"><span data-stu-id="d83d9-129">sync()</span></span>
 
<span data-ttu-id="d83d9-p108">要求コンテキストで **sync()** メソッドを呼び出すと、プロキシ オブジェクトと Excel ドキュメント内のオブジェクトの状態が同期されます。\*\* sync()\*\* メソッドは、要求コンテキストのキューに登録されたすべてのコマンドを実行し、プロキシ オブジェクトに読み込まれるプロパティの値を取得します。 **Sync()** メソッドでは、非同期的に実行し、**sync()** メソッドの完了時に解決する[約束](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p108">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document. The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects. The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>
 
<span data-ttu-id="d83d9-133">次の例は、ローカル JavaScript proxy オブジェクト (**selectedRange**) を定義し、そのオブジェクトのプロパティを読み込み、JavaScript の Promises パターンを使用して **context.sync()** を呼び出し、プロキシ オブジェクトと Excel ドキュメント内のオブジェクトの状態を同期するバッチ関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="d83d9-133">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>
 
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
 
<span data-ttu-id="d83d9-134">前の例では、**selectedRange** が設定され、**context.sync()** が呼び出されるとその **address** プロパティが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="d83d9-134">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>
 
<span data-ttu-id="d83d9-p109">**sync()** は 約束 を返す非同期の操作であるため、常に 約束を (JavaScript で) **返す**必要があります。これにより、スクリプトは実行を継続する前に **sync()** 操作が完了しています。\*\* sync()\*\* を用いたパフォーマンスの最適化の詳細については、「[   Excel JavaScript API のパフォーマンス最適化](https://docs.microsoft.com/office/dev/add-ins/excel/performance)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p109">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript). Doing so ensures that the **sync()** operation completes before the script continues to run. For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span></span>
 
### <a name="load"></a><span data-ttu-id="d83d9-138">load()</span><span class="sxs-lookup"><span data-stu-id="d83d9-138">load()</span></span>
 
<span data-ttu-id="d83d9-p110">プロキシ オブジェクトのプロパティを読み取るには、まず Excel ドキュメントからプロキシ オブジェクトとデータを入力するプロパティを明示的に読み込み、それから **context.sync()** を呼び出す必要があります。たとえば、選択範囲を参照するプロキシ オブジェクトを作成した後、選択範囲の\*\*  address\*\* プロパティを読み取る必要がある場合、読み取る前に\*\* address\*\* プロパティを読み込む必要があります。読み込むプロキシ オブジェクトのプロパティを要求するには、オブジェクトの **load()** メソッドを呼び出し、ロードするプロパティを指定します。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p110">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it. To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="d83d9-p111">プロキシ オブジェクト上でメソッドを呼び出す、またはプロパティを設定するだけの場合は、**load()** メソッドを呼び出す必要はありません。**load()** メソッドは、プロキシ オブジェクト上でプロパティを読み取る場合のみ必要です。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p111">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method. The **load()** method is only required when you want to read properties on a proxy object.</span></span>
 
<span data-ttu-id="d83d9-p112">プロキシ オブジェクトに対してプロパティを設定、またはメソッドを呼び出す要求と同じように、プロキシ オブジェクトに対してプロパティを読み込む要求も、要求コンテキストで保留中のコマンドのキューに追加され、次回 **sync()** メソッドを呼び出すときに実行されます。**load()** の呼び出しは、必要なだけ要求コンテキストのキューに入れることができます。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p112">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>
 
<span data-ttu-id="d83d9-146">次の例では、範囲の特定のプロパティのみが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="d83d9-146">In the following example, only specific properties of the range are loaded.</span></span>
 
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
 
<span data-ttu-id="d83d9-147">前の例では、`format/font` が **myRange.load()** の呼び出しで指定されていないため、`format.font.color` プロパティは読み取れませんでした。</span><span class="sxs-lookup"><span data-stu-id="d83d9-147">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="d83d9-p113">パフォーマンスを最適化するにはプロパティと [Excel JavaScript API のパフォーマンスの最適化](performance.md) で説明したように、オブジェクトの **load()** メソッドを使用する場合の読み込みに関係を明示的に指定する必要があります。 **Load()** メソッドの詳細については、 [Excel JavaScript API を使用して高度なプログラミングの概念](excel-add-ins-advanced-concepts.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p113">To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object, as covered in [Excel JavaScript API performance optimizations](performance.md). For more information about the **load()** method, see [Excel JavaScript API advanced concepts](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="d83d9-150">null または空白のプロパティ値</span><span class="sxs-lookup"><span data-stu-id="d83d9-150">null or blank property values</span></span>
 
### <a name="null-input-in-2-d-array"></a><span data-ttu-id="d83d9-151">2 次元配列での null の入力</span><span class="sxs-lookup"><span data-stu-id="d83d9-151">null input in 2-D Array</span></span>
 
<span data-ttu-id="d83d9-p114">Excel では、範囲は、最初の次元が行と2 番目の次元が列である、2 次元配列で表されます。値、数値の書式、または範囲内で特定のセルの数式を設定するに、2 次元配列の値、数値形式、またはそれらのセルの数式を指定して、 `null` 2 次元配列内のすべてのセルにします。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p114">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>
 
<span data-ttu-id="d83d9-p115">たとえば、範囲内の 1 つのセルの数値書式を更新し、範囲内の他のセルすべての既存の数値書式を保持する場合、更新するセルに新しい数値書式を指定し、他のセルすべてに `null` を指定します。次のコード スニペットでは、範囲内の 4 番目のセルに新しい数値書式を設定し、その前の 3 つのセルについては数値書式を変更せずに保持します。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p115">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a><span data-ttu-id="d83d9-156">プロパティに対する null の入力</span><span class="sxs-lookup"><span data-stu-id="d83d9-156">null input for a property</span></span>
 
<span data-ttu-id="d83d9-p116">`null` を、単独プロパティに対する有効な入力として指定することはできません。たとえば、次のコード スニペットは、範囲の **values** プロパティを `null` に設定できないため無効です。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p116">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>
 
```js
range.values = null;
```
 
<span data-ttu-id="d83d9-159">同様に、次のコード スニペットは、`null` が **color** プロパティで有効ではないため無効です。</span><span class="sxs-lookup"><span data-stu-id="d83d9-159">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a><span data-ttu-id="d83d9-160">応答内の null プロパティ値</span><span class="sxs-lookup"><span data-stu-id="d83d9-160">null property values in the response</span></span>
 
<span data-ttu-id="d83d9-p117">指定の範囲に複数の値がある場合、`size` および `color` などの書式設定プロパティでは、応答に `null` 値が含まれます。たとえば、範囲を取得してその`format.font.color`   プロパティを読み込む場合:</span><span class="sxs-lookup"><span data-stu-id="d83d9-p117">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>
 
* <span data-ttu-id="d83d9-163">範囲内のすべてのセルのフォントの色が同じ場合、`range.format.font.color` がその色を指定します。</span><span class="sxs-lookup"><span data-stu-id="d83d9-163">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="d83d9-164">範囲内に複数のフォントの色がある場合、`range.format.font.color` は `null` です。</span><span class="sxs-lookup"><span data-stu-id="d83d9-164">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>
 
### <a name="blank-input-for-a-property"></a><span data-ttu-id="d83d9-165">プロパティに対する空白の入力</span><span class="sxs-lookup"><span data-stu-id="d83d9-165">Blank input for a property</span></span>
 
<span data-ttu-id="d83d9-p118">プロパティに空白の値 (`''` の間にスペースのない 2 つの引用符) を指定すると、プロパティをクリアまたはリセットする指示として解釈されます。例:</span><span class="sxs-lookup"><span data-stu-id="d83d9-p118">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>
 
* <span data-ttu-id="d83d9-168">範囲の `values` プロパティに空白の値を指定すると、範囲のコンテンツはクリアされます。</span><span class="sxs-lookup"><span data-stu-id="d83d9-168">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
 
* <span data-ttu-id="d83d9-169">`numberFormat` プロパティに空白の値を指定すると、数値書式は `General` にリセットされます。</span><span class="sxs-lookup"><span data-stu-id="d83d9-169">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
 
* <span data-ttu-id="d83d9-170">`formula` プロパティと `formulaLocale` プロパティに空白の値を指定すると、数式の値はクリアされます。</span><span class="sxs-lookup"><span data-stu-id="d83d9-170">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>
 
### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="d83d9-171">応答内の空白のプロパティ値</span><span class="sxs-lookup"><span data-stu-id="d83d9-171">Blank property values in the response</span></span>
 
<span data-ttu-id="d83d9-p119">応答での空白のプロパティ値の読み取り操作は、(つまり、スペースを入れない `''`の間の2 つの引用符) そのセルが含まれていないデータまたは値を示します。最初次の例で、最初と最後のセル範囲のデータは含まれません。2 番目の例では、範囲内の最初の 2 つのセルには、数式が入力されません。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p119">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="d83d9-175">無制限の範囲への読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="d83d9-175">Read or write to an unbounded range</span></span>
 
### <a name="read-an-unbounded-range"></a><span data-ttu-id="d83d9-176">無制限の範囲の読み取り</span><span class="sxs-lookup"><span data-stu-id="d83d9-176">Read an unbounded range</span></span>
 
<span data-ttu-id="d83d9-p120">無制限の範囲のアドレスとは、列全体または行全体を指定する範囲のアドレスです。例:</span><span class="sxs-lookup"><span data-stu-id="d83d9-p120">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>
 
* <span data-ttu-id="d83d9-179">範囲のアドレスは、列全体で構成されます。</span><span class="sxs-lookup"><span data-stu-id="d83d9-179">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
* <span data-ttu-id="d83d9-180">範囲のアドレスは、行全体で構成されます。</span><span class="sxs-lookup"><span data-stu-id="d83d9-180">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>
 
<span data-ttu-id="d83d9-p121">API が無制限の範囲を取得する要求を行う場合 (`getRange('C:C')`)、返される応答には、`values`、`text`、`numberFormat`、または `formula` などのセル レベルのプロパティに `null` が含まれます。`address`、または `cellCount` などのその他の範囲プロパティは、無制限の範囲を反映します。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p121">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>
 
### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="d83d9-183">無制限の範囲への書き込み</span><span class="sxs-lookup"><span data-stu-id="d83d9-183">Write to an unbounded range</span></span>
 
<span data-ttu-id="d83d9-p122">セル レベルのプロパティを非制限の範囲で次のように設定することはできません `values`、 `numberFormat`、および `formula` 。これは入力の要求が大きすぎるためです。たとえば、次のコード スニペットは、無限の範囲の `values` を指定しようとするので、無効です、無限の範囲のセル レベルのプロパティを設定しようとした場合、API はエラーを返します。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p122">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large. For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range. The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="d83d9-187">広い範囲に対する読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="d83d9-187">Read or write to a large range</span></span>
 
<span data-ttu-id="d83d9-p123">範囲に多数のセル、値、数値の書式、数式が含まれている場合は、その範囲に API の操作を実行することはできません。範囲上で要求された操作を実行するため、API の最適な試行を必ず確認 (つまり、指定したデータの取得または書き込み)しますが、広い範囲での 読み取りまたは書き込み操作の実行を試みると、過剰なリソース使用率による API エラーが発生します。このようなエラーを避けるためには、別の読み取りを実行するか、1 つを実行しようとしてではなく、広い範囲の小さなサブセットを操作の読み取りまたは書き込み操作に大きな範囲での書き込みをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p123">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>
 
## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="d83d9-191">範囲内のすべてのセルの更新</span><span class="sxs-lookup"><span data-stu-id="d83d9-191">Update all cells in a range</span></span>
 
<span data-ttu-id="d83d9-192">範囲内のすべてのセルに同じ更新 (すべてのセルに同じ値を入力する、同じ数値書式を設定する、同じ数式ですべてのセルにデータを入力するなど) を適用するには、**range** オブジェクトの該当するプロパティを必要な 1 つの値に設定します。</span><span class="sxs-lookup"><span data-stu-id="d83d9-192">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>
 
<span data-ttu-id="d83d9-193">次の例では、20 個のセルを含む範囲を取得し、数値書式を設定してその範囲のすべてのセルに **3/11/2015** という値を設定します。</span><span class="sxs-lookup"><span data-stu-id="d83d9-193">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>
 
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
 
## <a name="error-messages"></a><span data-ttu-id="d83d9-194">エラー メッセージ</span><span class="sxs-lookup"><span data-stu-id="d83d9-194">Error messages</span></span>
 
<span data-ttu-id="d83d9-p124">API エラーが発生すると、API ではコードとメッセージを含む **error** オブジェクトが返されます。次の表は、API から返されるエラー一覧の定義を示します。</span><span class="sxs-lookup"><span data-stu-id="d83d9-p124">When an API error occurs, the API will return an **error** object that contains a code and a message. The following table defines a list of errors that the API may return.</span></span>
 
|<span data-ttu-id="d83d9-197">error.code</span><span class="sxs-lookup"><span data-stu-id="d83d9-197">error.code</span></span> | <span data-ttu-id="d83d9-198">error.message</span><span class="sxs-lookup"><span data-stu-id="d83d9-198">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="d83d9-199">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="d83d9-199">InvalidArgument</span></span> |<span data-ttu-id="d83d9-200">引数が無効であるか、存在しません。または形式が正しくありません。</span><span class="sxs-lookup"><span data-stu-id="d83d9-200">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="d83d9-201">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="d83d9-201">InvalidRequest</span></span>  |<span data-ttu-id="d83d9-202">要求を処理できません。</span><span class="sxs-lookup"><span data-stu-id="d83d9-202">Cannot process the request.</span></span>|
|<span data-ttu-id="d83d9-203">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="d83d9-203">InvalidReference</span></span>|<span data-ttu-id="d83d9-204">この参照は、現在の操作に対して無効です。</span><span class="sxs-lookup"><span data-stu-id="d83d9-204">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="d83d9-205">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="d83d9-205">InvalidBinding</span></span>  |<span data-ttu-id="d83d9-206">このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。</span><span class="sxs-lookup"><span data-stu-id="d83d9-206">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="d83d9-207">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="d83d9-207">InvalidSelection</span></span>|<span data-ttu-id="d83d9-208">現在の選択内容は、この操作では無効です。</span><span class="sxs-lookup"><span data-stu-id="d83d9-208">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="d83d9-209">認証されていません</span><span class="sxs-lookup"><span data-stu-id="d83d9-209">Unauthenticated</span></span> |<span data-ttu-id="d83d9-210">必要な認証情報が見つからないか、無効です。</span><span class="sxs-lookup"><span data-stu-id="d83d9-210">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="d83d9-211">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="d83d9-211">AccessDenied</span></span> |<span data-ttu-id="d83d9-212">要求された操作を実行できません。</span><span class="sxs-lookup"><span data-stu-id="d83d9-212">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="d83d9-213">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="d83d9-213">ItemNotFound</span></span> |<span data-ttu-id="d83d9-214">要求されたリソースは存在しません。</span><span class="sxs-lookup"><span data-stu-id="d83d9-214">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="d83d9-215">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="d83d9-215">ActivityLimitReached</span></span>|<span data-ttu-id="d83d9-216">アクティビティの制限に達しました。</span><span class="sxs-lookup"><span data-stu-id="d83d9-216">Activity limit has been reached.</span></span>|
|<span data-ttu-id="d83d9-217">GeneralException</span><span class="sxs-lookup"><span data-stu-id="d83d9-217">GeneralException</span></span>|<span data-ttu-id="d83d9-218">リクエストの処理中に内部エラーが発生しました。</span><span class="sxs-lookup"><span data-stu-id="d83d9-218">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="d83d9-219">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="d83d9-219">NotImplemented</span></span>  |<span data-ttu-id="d83d9-220">リクエストされた機能は実装されていません。</span><span class="sxs-lookup"><span data-stu-id="d83d9-220">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="d83d9-221">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="d83d9-221">ServiceNotAvailable</span></span>|<span data-ttu-id="d83d9-222">サービスを利用できません。</span><span class="sxs-lookup"><span data-stu-id="d83d9-222">The service is unavailable.</span></span>|
|<span data-ttu-id="d83d9-223">一致しません</span><span class="sxs-lookup"><span data-stu-id="d83d9-223">Conflict</span></span>              |<span data-ttu-id="d83d9-224">競合のため、要求を処理できませんでした。</span><span class="sxs-lookup"><span data-stu-id="d83d9-224">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="d83d9-225">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="d83d9-225">ItemAlreadyExists</span></span>|<span data-ttu-id="d83d9-226">作成中のリソースはすでに存在しています。</span><span class="sxs-lookup"><span data-stu-id="d83d9-226">The resource being created already exists.</span></span>|
|<span data-ttu-id="d83d9-227">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="d83d9-227">UnsupportedOperation</span></span>|<span data-ttu-id="d83d9-228">試行中の操作はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d83d9-228">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="d83d9-229">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="d83d9-229">RequestAborted</span></span>|<span data-ttu-id="d83d9-230">実行時に要求が中止されました。</span><span class="sxs-lookup"><span data-stu-id="d83d9-230">The request was aborted during run time.</span></span>|
|<span data-ttu-id="d83d9-231">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="d83d9-231">ApiNotAvailable</span></span>|<span data-ttu-id="d83d9-232">要求された API は使用できません。</span><span class="sxs-lookup"><span data-stu-id="d83d9-232">The requested API is not available.</span></span>|
|<span data-ttu-id="d83d9-233">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="d83d9-233">InsertDeleteConflict</span></span>|<span data-ttu-id="d83d9-234">試行された挿入操作または削除操作で競合が発生しました。</span><span class="sxs-lookup"><span data-stu-id="d83d9-234">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="d83d9-235">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="d83d9-235">InvalidOperation</span></span>|<span data-ttu-id="d83d9-236">試行された操作は、このオブジェクトでは無効です。</span><span class="sxs-lookup"><span data-stu-id="d83d9-236">The operation attempted is invalid on the object.</span></span>|
 
## <a name="see-also"></a><span data-ttu-id="d83d9-237">関連項目</span><span class="sxs-lookup"><span data-stu-id="d83d9-237">See also</span></span>
 
* [<span data-ttu-id="d83d9-238">Excel アドインを使う</span><span class="sxs-lookup"><span data-stu-id="d83d9-238">Get started with Excel add-ins</span></span>](excel-add-ins-get-started-overview.md)
* [<span data-ttu-id="d83d9-239">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="d83d9-239">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
* [<span data-ttu-id="d83d9-240">Excel JavaScript API を使用した高度なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="d83d9-240">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
* [<span data-ttu-id="d83d9-241">Excel JavaScript API パフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="d83d9-241">Excel JavaScript API performance optimization</span></span>](https://docs.microsoft.com/office/dev/add-ins/excel/performance)
* [<span data-ttu-id="d83d9-242">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="d83d9-242">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
