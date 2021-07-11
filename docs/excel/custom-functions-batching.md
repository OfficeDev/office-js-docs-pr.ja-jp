---
ms.date: 07/10/2019
description: バッチ処理カスタム関数を組み合わせてリモート サービスへのネットワーク呼び出しを減らします。
title: リモート サービスのためのバッチ処理カスタム関数の呼び出し
localization_priority: Normal
ms.openlocfilehash: 0729e06df5f6e26f9726e1de0dcdaac0f101b18d
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349652"
---
# <a name="batching-custom-function-calls-for-a-remote-service"></a><span data-ttu-id="9a378-103">リモート サービスのためのバッチ処理カスタム関数の呼び出し</span><span class="sxs-lookup"><span data-stu-id="9a378-103">Batching custom function calls for a remote service</span></span>

<span data-ttu-id="9a378-104">カスタム関数がリモート サービスを呼び出す場合は、リモート サービスへのネットワークの呼び出し数を減らすバッチ処理のパターンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="9a378-104">If your custom functions call a remote service you can use a batching pattern to reduce the number of network calls to the remote service.</span></span> <span data-ttu-id="9a378-105">バッチ処理をしたネットワーク ラウンド トリップのウェブ サービスへのすべての呼び出しを、1 回に減らします。</span><span class="sxs-lookup"><span data-stu-id="9a378-105">To reduce network round trips you batch all the calls into a single call to the web service.</span></span> <span data-ttu-id="9a378-106">これは、ワークシートが再計算するときに最適な方法です。</span><span class="sxs-lookup"><span data-stu-id="9a378-106">This is ideal when the spreadsheet is recalculated.</span></span>

<span data-ttu-id="9a378-107">たとえば、別のユーザーがスプレッドシートの 100 セル内でカスタム関数を使用し、スプレッドシートを再計算した場合、カスタム関数は 100 回実行され、100 回ネットワークの呼び出しを行います。</span><span class="sxs-lookup"><span data-stu-id="9a378-107">For example, if someone used your custom function in 100 cells in a spreadsheet, and then recalculated the spreadsheet, your custom function would run 100 times and make 100 network calls.</span></span> <span data-ttu-id="9a378-108">バッチ処理のパターンを使用すると、1 つのネットワークの呼び出しで 100 の計算すべてを結合することができます。</span><span class="sxs-lookup"><span data-stu-id="9a378-108">By using a batching pattern, the calls can be combined to make all 100 calculations in a single network call.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a><span data-ttu-id="9a378-109">完成したサンプルを表示する</span><span class="sxs-lookup"><span data-stu-id="9a378-109">View the completed sample</span></span>

<span data-ttu-id="9a378-110">この記事を参考にして、自分のプロジェクトにコードの例を貼り付けることができます。</span><span class="sxs-lookup"><span data-stu-id="9a378-110">You can follow this article and paste the code examples into your own project.</span></span> <span data-ttu-id="9a378-111">たとえば、[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して TypeScript 用の新しいカスタム関数プロジェクトを作成し、この記事のすべてのコードをそのプロジェクトに追加することができます。</span><span class="sxs-lookup"><span data-stu-id="9a378-111">For example, you can use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create a new custom function project for TypeScript, then add all the code from this article to the project.</span></span> <span data-ttu-id="9a378-112">その後、コードを実行して試してください。</span><span class="sxs-lookup"><span data-stu-id="9a378-112">You can then run the code and try it out.</span></span>

<span data-ttu-id="9a378-113">[カスタム関数のバッチ処理パターン](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching)で完全なサンプル プロジェクトをダウンロードまたは表示することができます。</span><span class="sxs-lookup"><span data-stu-id="9a378-113">Also, you can download or view the complete sample project at [Custom function batching pattern](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching).</span></span> <span data-ttu-id="9a378-114">読み進める前に全体のコードを表示したい場合、 [スクリプト ファイル](https://github.com/OfficeDev/PnP-OfficeAddins/blob/main/Excel-custom-functions/Batching/src/functions/functions.js)をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="9a378-114">If you want to view the code in whole before reading any further, take a look at the [script file](https://github.com/OfficeDev/PnP-OfficeAddins/blob/main/Excel-custom-functions/Batching/src/functions/functions.js).</span></span>

## <a name="create-the-batching-pattern-in-this-article"></a><span data-ttu-id="9a378-115">この記事内でバッチ処理パターンを作成する</span><span class="sxs-lookup"><span data-stu-id="9a378-115">Create the batching pattern in this article</span></span>

<span data-ttu-id="9a378-116">カスタム関数にバッチ処理を設定するには、次の 3 つの主要なセクションのコードを記述する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9a378-116">To set up batching for your custom functions you'll need to write three main sections of code.</span></span>

1. <span data-ttu-id="9a378-117">バッチに新しい操作を追加するプッシュ操作の呼び出しのたびに、Excel はカスタム関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="9a378-117">A push operation to add a new operation to the batch of calls each time Excel calls your custom function.</span></span>
2. <span data-ttu-id="9a378-118">バッチの準備ができたときのリモート要求を行う関数です。</span><span class="sxs-lookup"><span data-stu-id="9a378-118">A function to make the remote request when the batch is ready.</span></span>
3. <span data-ttu-id="9a378-119">バッチ要求に応答するサーバー コードは、すべての操作の結果を計算して値を返します。</span><span class="sxs-lookup"><span data-stu-id="9a378-119">Server code to respond to the batch request, calculate all of the operation results, and return the values.</span></span>

<span data-ttu-id="9a378-120">次のセクションでは、一度に 1 つのコード例を構築する方法が表示されます。</span><span class="sxs-lookup"><span data-stu-id="9a378-120">In the following sections you will be shown how to construct the code one example at a time.</span></span> <span data-ttu-id="9a378-121">**functions.ts** ファイルにそれぞれのコード例を追加します。</span><span class="sxs-lookup"><span data-stu-id="9a378-121">You'll add each code example to your **functions.ts** file.</span></span> <span data-ttu-id="9a378-122">Yo Office ジェネレーター 使用して、新しいカスタム関数のプロジェクトを作成することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="9a378-122">It's recommended you create a brand new custom functions project using the Yo Office generator.</span></span> <span data-ttu-id="9a378-123">新しいプロジェクトを作成するには、 [Excel のカスタム関数の開発を開始する](../quickstarts/excel-custom-functions-quickstart.md)を参照し、JavaScript ではなく TypeScript を使用してください。</span><span class="sxs-lookup"><span data-stu-id="9a378-123">To create a new project see [Get started developing Excel custom functions](../quickstarts/excel-custom-functions-quickstart.md) and use TypeScript instead of JavaScript.</span></span>

## <a name="batch-each-call-to-your-custom-function"></a><span data-ttu-id="9a378-124">カスタム関数の各呼び出しにバッチ処理をする</span><span class="sxs-lookup"><span data-stu-id="9a378-124">Batch each call to your custom function</span></span>

<span data-ttu-id="9a378-125">操作を実行するリモート サービスの呼び出し機能を使ってカスタム関数の演算を実行し、必要な結果を計算します。</span><span class="sxs-lookup"><span data-stu-id="9a378-125">Your custom functions work by calling a remote service to perform the operation and calculate the result they need.</span></span> <span data-ttu-id="9a378-126">要求された各操作をバッチ内に保存する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="9a378-126">This provides a way for them to store each requested operation into a batch.</span></span> <span data-ttu-id="9a378-127">後で、その操作にバッチ処理をする `_pushOperation`関数を作成する方法が表示されます。</span><span class="sxs-lookup"><span data-stu-id="9a378-127">Later you'll see how to create a `_pushOperation` function to batch the operations.</span></span> <span data-ttu-id="9a378-128">最初に、カスタム関数から`_pushOperation`を呼び出す方法については、次のコード例をみてください。</span><span class="sxs-lookup"><span data-stu-id="9a378-128">First, take a look at the following code example to see how to call `_pushOperation` from your custom function.</span></span>

<span data-ttu-id="9a378-129">次のコードでは、カスタム関数は除算を実行しますが、実際の計算を実行するにはリモート サービスに依存しています。</span><span class="sxs-lookup"><span data-stu-id="9a378-129">In the following code, the custom function performs division but relies on a remote service to do the actual calculation.</span></span> <span data-ttu-id="9a378-130">リモート サービスにその操作と別の操作を一緒にバッチ処理し、`_pushOperation`を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="9a378-130">It calls `_pushOperation` to batch the operation along with other operations to the remote service.</span></span> <span data-ttu-id="9a378-131">その名称は **div2** 操作といいます。</span><span class="sxs-lookup"><span data-stu-id="9a378-131">It names the operation **div2**.</span></span> <span data-ttu-id="9a378-132">リモート サービスが同じスキーム (詳細については、この後のリモート サービスで) を使用する限り、任意の名前付けスキームを操作に使用することができます。</span><span class="sxs-lookup"><span data-stu-id="9a378-132">You can use any naming scheme you want for operations as long as the remote service is also using the same scheme (more on the remote service later).</span></span> <span data-ttu-id="9a378-133">また、操作を実行する必要があるリモートサービスの引数が渡されます。</span><span class="sxs-lookup"><span data-stu-id="9a378-133">Also, the arguments the remote service will need to run the operation are passed.</span></span>

### <a name="add-the-div2-custom-function-to-functionsts"></a><span data-ttu-id="9a378-134">functions.ts に div2 カスタム関数を追加する</span><span class="sxs-lookup"><span data-stu-id="9a378-134">Add the div2 custom function to functions.ts</span></span>

```typescript
/**
 * @CustomFunction
 * Divides two numbers using batching
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend: number, divisor: number) {
  return _pushOperation(
    "div2",
    [dividend, divisor]
  );
}
```

<span data-ttu-id="9a378-135">次に、1 つのネットワークの呼び出しに渡されるすべての操作が格納されるバッチの配列を定義します。</span><span class="sxs-lookup"><span data-stu-id="9a378-135">Next, you will define the batch array which will store all operations to be passed in one network call.</span></span> <span data-ttu-id="9a378-136">次のコードでは、配列内で各バッチのエントリを記述するインターフェイスを定義する方法を表示します。</span><span class="sxs-lookup"><span data-stu-id="9a378-136">The following code shows how to define an interface describing each batch entry in the array.</span></span> <span data-ttu-id="9a378-137">どの文字列名のどの操作を実行するのか、インターフェイスが操作を定義します。</span><span class="sxs-lookup"><span data-stu-id="9a378-137">The interface defines an operation, which is a string name of which operation to run.</span></span> <span data-ttu-id="9a378-138">たとえば、 `multiply` と `divide`という名前の 2 つのカスタム関数がある場合、バッチのエントリ内で操作名として再利用できます。</span><span class="sxs-lookup"><span data-stu-id="9a378-138">For example, if you had two custom functions named `multiply` and `divide`, you could reuse those as the operation names in your batch entries.</span></span> <span data-ttu-id="9a378-139">`args` は、Excel からカスタム関数に渡された引数が保持されます。</span><span class="sxs-lookup"><span data-stu-id="9a378-139">`args` will hold the arguments that were passed to your custom function from Excel.</span></span> <span data-ttu-id="9a378-140">最後に、`resolve` または `reject`はリモート サービスが返した情報を保持している promise を格納します。</span><span class="sxs-lookup"><span data-stu-id="9a378-140">And finally, `resolve` or `reject` will store a promise holding the information the remote service returns.</span></span>

```typescript
interface IBatchEntry {
  operation: string;
  args: any[];
  resolve: (data: any) => void;
  reject: (error: Error) => void;
}
```

<span data-ttu-id="9a378-141">次に、前のインターフェイスを使用するバッチの配列を作成します。</span><span class="sxs-lookup"><span data-stu-id="9a378-141">Next, create the batch array that uses the previous interface.</span></span> <span data-ttu-id="9a378-142">バッチが予定されているかどうかを追跡するため、`_isBatchedRequestSchedule` 変数を作成します。</span><span class="sxs-lookup"><span data-stu-id="9a378-142">To track if a batch is scheduled or not, create an `_isBatchedRequestSchedule` variable.</span></span> <span data-ttu-id="9a378-143">リモート サービスへのバッチの呼び出しのタイミングは、後で重要になります。</span><span class="sxs-lookup"><span data-stu-id="9a378-143">This will be important later for timing batch calls to the remote service.</span></span>

```typescript
const _batch: IBatchEntry[] = [];
let _isBatchedRequestScheduled = false;
```

<span data-ttu-id="9a378-144">最後に、Excel がカスタム関数を呼び出すと、バッチ配列への操作をプッシュする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9a378-144">Finally when Excel calls your custom function, you need to push the operation into the batch array.</span></span> <span data-ttu-id="9a378-145">次のコードでは、カスタム関数から新しい操作を追加する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="9a378-145">The following code shows how to add a new operation from a custom function.</span></span> <span data-ttu-id="9a378-146">新しいバッチ エントリを作成し、処理を解決または拒否するための新しい promise を作成し、そしてバッチ配列にエントリをプッシュします。</span><span class="sxs-lookup"><span data-stu-id="9a378-146">It creates a new batch entry, creates a new promise to resolve or reject the operation, and pushes the entry into the batch array.</span></span>

<span data-ttu-id="9a378-147">このコードは、バッチがスケジュールされているかどうかも確認します。</span><span class="sxs-lookup"><span data-stu-id="9a378-147">This code also checks to see if a batch is scheduled.</span></span> <span data-ttu-id="9a378-148">この例では、それぞれのバッチはすべて100 ミリ秒ごとに実行するようスケジュールされています。</span><span class="sxs-lookup"><span data-stu-id="9a378-148">In this example, each batch is scheduled to run every 100ms.</span></span> <span data-ttu-id="9a378-149">必要に応じて、この値を調整することができます。</span><span class="sxs-lookup"><span data-stu-id="9a378-149">You can adjust this value as needed.</span></span> <span data-ttu-id="9a378-150">高い値は、リモート サービスに送信される大きなバッチで発生し、ユーザーが結果を確認するまでの応答時間が長くなります。</span><span class="sxs-lookup"><span data-stu-id="9a378-150">Higher values result in bigger batches being sent to the remote service, and a longer wait time for the user to see results.</span></span> <span data-ttu-id="9a378-151">小さい値は、より多くのバッチがリモート サービスに送信されますが、ユーザーの応答時間は短くなる傾向があります。</span><span class="sxs-lookup"><span data-stu-id="9a378-151">Lower values tend to send more batches to the remote service, but with a quick response time for users.</span></span>

### <a name="add-the-_pushoperation-function-to-functionsts"></a><span data-ttu-id="9a378-152">functions.ts に `_pushOperation` 関数を追加する</span><span class="sxs-lookup"><span data-stu-id="9a378-152">Add the `_pushOperation` function to functions.ts</span></span>

```typescript
function _pushOperation(op: string, args: any[]) {
  // Create an entry for your custom function.
  const invocationEntry: IBatchEntry = {
    operation: op, // e.g. sum
    args: args,
    resolve: undefined,
    reject: undefined,
  };

  // Create a unique promise for this invocation,
  // and save its resolve and reject functions into the invocation entry.
  const promise = new Promise((resolve, reject) => {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch.
  _batch.push(invocationEntry);

  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g. 100 ms.
  if (!_isBatchedRequestScheduled) {
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a><span data-ttu-id="9a378-153">リモートの要求を行う</span><span class="sxs-lookup"><span data-stu-id="9a378-153">Make the remote request</span></span>

<span data-ttu-id="9a378-154">`_makeRemoteRequest`関数の目的は、操作のバッチをリモート サービスに渡し、それから各カスタム関数に結果を返します。</span><span class="sxs-lookup"><span data-stu-id="9a378-154">The purpose of the `_makeRemoteRequest` function is to pass the batch of operations to the remote service, and then return the results to each custom function.</span></span> <span data-ttu-id="9a378-155">まず、バッチ配列のコピーを作成します。</span><span class="sxs-lookup"><span data-stu-id="9a378-155">It first creates a copy of the batch array.</span></span> <span data-ttu-id="9a378-156">これにより、concurrent カスタム関数は、Excel からすぐに新しい配列にバッチ処理を呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="9a378-156">This allows concurrent custom function calls from Excel to immediately begin batching in a new array.</span></span> <span data-ttu-id="9a378-157">そのコピーは、それから promise 情報が含まれていない単純な配列になります。</span><span class="sxs-lookup"><span data-stu-id="9a378-157">The copy is then turned into a simpler array that does not contain the promise information.</span></span> <span data-ttu-id="9a378-158">機能しない場合は、リモート サービスにその promise を渡しても意味をなしません。</span><span class="sxs-lookup"><span data-stu-id="9a378-158">It wouldn't make sense to pass the promises to a remote service since they would not work.</span></span> <span data-ttu-id="9a378-159">リモート サービスが何を返すかによって、`_makeRemoteRequest` は拒否するか、またはそれぞれの promise を解決します。</span><span class="sxs-lookup"><span data-stu-id="9a378-159">The `_makeRemoteRequest` will either reject or resolve each promise based on what the remote service returns.</span></span>

### <a name="add-the-following-_makeremoterequest-method-to-functionsts"></a><span data-ttu-id="9a378-160">次の`_makeRemoteRequest`メソッドを functions.ts に追加します。</span><span class="sxs-lookup"><span data-stu-id="9a378-160">Add the following `_makeRemoteRequest` method to functions.ts</span></span>

```typescript
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });

  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
        } else {
          console.log(response);
          batchCopy[index].resolve(response.result);
        }
      });
    });
}
```

### <a name="modify-_makeremoterequest-for-your-own-solution"></a><span data-ttu-id="9a378-161">独自のソリューションに`_makeRemoteRequest`を変更します。</span><span class="sxs-lookup"><span data-stu-id="9a378-161">Modify `_makeRemoteRequest` for your own solution</span></span>

<span data-ttu-id="9a378-162">`_makeRemoteRequest`関数は、あとで表示されますが、リモート サービスを表すモックの`_fetchFromRemoteService`を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="9a378-162">The `_makeRemoteRequest` function calls `_fetchFromRemoteService` which, as you'll see later, is just a mock representing the remote service.</span></span> <span data-ttu-id="9a378-163">これにより、簡単に学習でき、この記事でコードを実行することができます。</span><span class="sxs-lookup"><span data-stu-id="9a378-163">This makes it easier to study and run the code in this article.</span></span> <span data-ttu-id="9a378-164">ただし、実際のリモート サービスでこのコードを使用する場合は、次の変更を行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="9a378-164">But when you want to use this code for an actual remote service you should make the following changes.</span></span>

- <span data-ttu-id="9a378-165">ネットワーク経由でバッチ処理をシリアル化する方法を決定します。</span><span class="sxs-lookup"><span data-stu-id="9a378-165">Decide how to serialize the batch operations over the network.</span></span> <span data-ttu-id="9a378-166">たとえば、JSON の本文に、配列を配置することがあります。</span><span class="sxs-lookup"><span data-stu-id="9a378-166">For example, you may want to put the array into a JSON body.</span></span>
- <span data-ttu-id="9a378-167">`_fetchFromRemoteService`を呼び出す代わりに、バッチ処理を渡すリモート サービスに実際にネットワークの呼び出しをする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9a378-167">Instead of calling `_fetchFromRemoteService` you need to make the actual network call to the remote service passing the batch of operations.</span></span>

## <a name="process-the-batch-call-on-the-remote-service"></a><span data-ttu-id="9a378-168">リモート サービスでバッチの呼び出しを処理します。</span><span class="sxs-lookup"><span data-stu-id="9a378-168">Process the batch call on the remote service</span></span>

<span data-ttu-id="9a378-169">最後の手順では、リモート サービスでバッチの呼び出しを処理をします。</span><span class="sxs-lookup"><span data-stu-id="9a378-169">The last step is to handle the batch call in the remote service.</span></span> <span data-ttu-id="9a378-170">つぎのコード サンプルは、`_fetchFromRemoteService`関数を表しています。</span><span class="sxs-lookup"><span data-stu-id="9a378-170">The following code sample shows the `_fetchFromRemoteService` function.</span></span> <span data-ttu-id="9a378-171">この関数は、それぞれの操作を展開せずに指定した操作を実行し、それから結果を返します。</span><span class="sxs-lookup"><span data-stu-id="9a378-171">This function unpacks each operation, performs the specified operation, and returns the results.</span></span> <span data-ttu-id="9a378-172">この記事の学習の目的は、 `_fetchFromRemoteService`関数がリモート サービスを web アドインで実行し、リモート サービスをモックするように設計されています。</span><span class="sxs-lookup"><span data-stu-id="9a378-172">For learning purposes in this article, the `_fetchFromRemoteService` function is designed to run in your web add-in and mock a remote service.</span></span> <span data-ttu-id="9a378-173">**functions.ts** ファイルにこのコードを追加することができ、実際のリモート サービスを設定しなくても、この記事内のすべてのコードを学習し実行することができます。</span><span class="sxs-lookup"><span data-stu-id="9a378-173">You can add this code to your **functions.ts** file so that you can study and run all the code in this article without having to set up an actual remote service.</span></span>

### <a name="add-the-following-_fetchfromremoteservice-function-to-functionsts"></a><span data-ttu-id="9a378-174">次の `_fetchFromRemoteService` 関数を functions.ts に追加します。</span><span class="sxs-lookup"><span data-stu-id="9a378-174">Add the following `_fetchFromRemoteService` function to functions.ts</span></span>

```typescript
async function _fetchFromRemoteService(
  requestBatch: Array<{ operation: string, args: any[] }>
): Promise<IServerResponse[]> {
  // Simulate a slow network request to the server;
  await pause(1000);

  return requestBatch.map((request): IServerResponse => {
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myresult = args[0] * args[1];
        console.log(myresult);
        return {
          result: myresult
        };
      } else {
        return {
          error: `Operation not supported: ${operation}`
        };
      }
    } catch (error) {
      return {
        error: `Operation failed: ${operation}`
      };
    }
  });
}

function pause(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-_fetchfromremoteservice-for-your-live-remote-service"></a><span data-ttu-id="9a378-175">`_fetchFromRemoteService`をライブ リモート サービスに変更する</span><span class="sxs-lookup"><span data-stu-id="9a378-175">Modify `_fetchFromRemoteService` for your live remote service</span></span>

<span data-ttu-id="9a378-176">ライブ リモート サービス `_fetchFromRemoteService` で実行する関数を変更するには、次の変更を行います。</span><span class="sxs-lookup"><span data-stu-id="9a378-176">To modify the `_fetchFromRemoteService` function to run in your live remote service, make the following changes.</span></span>

- <span data-ttu-id="9a378-177">サーバー プラットフォーム (Node.js またはその他) のマップによっては、クライアント ネットワークがこの関数を呼び出します。 </span><span class="sxs-lookup"><span data-stu-id="9a378-177">Depending on your server platform (Node.js or others) map the client network call to this function.</span></span>
- <span data-ttu-id="9a378-178">モックの一部としてネットワークの遅延をシミュレートする`pause`関数を削除する。</span><span class="sxs-lookup"><span data-stu-id="9a378-178">Remove the `pause` function which simulates network latency as part of the mock.</span></span>
- <span data-ttu-id="9a378-179">パラメーターがネットワーク用に変更された場合、渡されたパラメーターで動作する関数の宣言を変更します。</span><span class="sxs-lookup"><span data-stu-id="9a378-179">Modify the function declaration to work with the parameter passed if the parameter is changed for network purposes.</span></span> <span data-ttu-id="9a378-180">たとえば、配列の代わりに、JSON 本体のバッチ処理で処理をします。</span><span class="sxs-lookup"><span data-stu-id="9a378-180">For example, instead of an array, it may be a JSON body of batched operations to process.</span></span>
- <span data-ttu-id="9a378-181">操作を実行する関数を変更する (または、操作を実行する関数を呼び出す)。</span><span class="sxs-lookup"><span data-stu-id="9a378-181">Modify the function to perform the operations (or call functions that do the operations).</span></span>
- <span data-ttu-id="9a378-182">適切な認証機構を適用する。</span><span class="sxs-lookup"><span data-stu-id="9a378-182">Apply an appropriate authentication mechanism.</span></span> <span data-ttu-id="9a378-183">適切な呼び出し元のみが関数にアクセスできることを確認します。</span><span class="sxs-lookup"><span data-stu-id="9a378-183">Ensure that only the correct callers can access the function.</span></span>
- <span data-ttu-id="9a378-184">リモート サービスで、コードを配置します。</span><span class="sxs-lookup"><span data-stu-id="9a378-184">Place the code in the remote service.</span></span>

## <a name="next-steps"></a><span data-ttu-id="9a378-185">次の手順</span><span class="sxs-lookup"><span data-stu-id="9a378-185">Next steps</span></span>
<span data-ttu-id="9a378-186">カスタム関数で使用できる[さまざまなパラメーター](custom-functions-parameter-options.md)について確認してください。</span><span class="sxs-lookup"><span data-stu-id="9a378-186">Learn about [the various parameters](custom-functions-parameter-options.md) you can use in your custom functions.</span></span> <span data-ttu-id="9a378-187">または、[カスタム関数で Web 通話](custom-functions-web-reqs.md)を発信する際の基本事項を確認してください。</span><span class="sxs-lookup"><span data-stu-id="9a378-187">Or review the basics behind making [a web call through a custom function](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="9a378-188">関連項目</span><span class="sxs-lookup"><span data-stu-id="9a378-188">See also</span></span>

* [<span data-ttu-id="9a378-189">関数の揮発性の値</span><span class="sxs-lookup"><span data-stu-id="9a378-189">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="9a378-190">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="9a378-190">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="9a378-191">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="9a378-191">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
