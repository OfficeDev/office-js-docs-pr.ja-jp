---
ms.date: 05/03/2019
description: Excel のカスタム関数に関する一般的な問題をトラブルシューティングします。
title: カスタム関数のトラブルシューティング
localization_priority: Priority
ms.openlocfilehash: 04da6d58c2610130961a1b89d2b9a1101b54bcb2
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628012"
---
# <a name="troubleshoot-custom-functions"></a><span data-ttu-id="cbcba-103">カスタム関数のトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="cbcba-103">Troubleshoot custom functions</span></span>

<span data-ttu-id="cbcba-104">カスタム関数を作成してテストするとき、製品でエラーが発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="cbcba-104">When developing custom functions, you may encounter errors in the product while creating and testing your functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="cbcba-105">問題を解決するには、[ランタイム ログを有効にしてエラーをキャプチャ](#enable-runtime-logging)し、[Excel のネイティブ エラー メッセージ](#check-for-excel-error-messages)を参照します。</span><span class="sxs-lookup"><span data-stu-id="cbcba-105">To resolve issues, you can [enable runtime logging to capture errors](#enable-runtime-logging) and refer to [Excel's native error messages](#check-for-excel-error-messages).</span></span> <span data-ttu-id="cbcba-106">また、[予約を未解決のままにしたり](#ensure-promises-return)、[機能の関連付け](#my-functions-wont-load-associate-functions)を忘れてしまうといったよくある間違いを確認します。</span><span class="sxs-lookup"><span data-stu-id="cbcba-106">Also, check for common mistakes such as not [verifying ssl certificates](#ensure-promises-return) properly, [leaving promises unresolved](#my-functions-wont-load-associate-functions), and forgetting to associate your functions.</span></span>

## <a name="enable-runtime-logging"></a><span data-ttu-id="cbcba-107">ランタイム ログを有効にする</span><span class="sxs-lookup"><span data-stu-id="cbcba-107">Enable runtime logging</span></span>

<span data-ttu-id="cbcba-108">Windows 上の Office でアドインをテストする場合は、[ランタイム ログを有効にする](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbcba-108">If you are testing your add-in in Office on Windows, you should [enable runtime logging](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span> <span data-ttu-id="cbcba-109">ランタイム ログでは、問題解明用に別に作成したログ ファイルに `console.log` ステートメントが配信されます。</span><span class="sxs-lookup"><span data-stu-id="cbcba-109">Runtime logging delivers `console.log` statements to a separate log file you create to help you uncover issues.</span></span> <span data-ttu-id="cbcba-110">ステートメントでは、アドインの XML マニフェスト ファイルに関するエラー、実行時の条件、カスタム関数のインストールなど、さまざまなエラーがカバーされます。</span><span class="sxs-lookup"><span data-stu-id="cbcba-110">The statements cover a variety of errors, including errors pertaining to your add-in's XML manifest file, runtime conditions, or installation of your custom functions.</span></span>  <span data-ttu-id="cbcba-111">ランタイム ログの詳細については、「[アドインのデバッグにランタイム ログを使用する](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="cbcba-111">For more information about runtime logging, see [Use runtime logging to debug your add-in](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span>  

### <a name="check-for-excel-error-messages"></a><span data-ttu-id="cbcba-112">Excel のエラー メッセージを確認する</span><span class="sxs-lookup"><span data-stu-id="cbcba-112">Check for Excel error messages</span></span>

<span data-ttu-id="cbcba-113">Excel には多くの組み込みエラー メッセージがあり、計算エラーが発生するとセルに返されます。</span><span class="sxs-lookup"><span data-stu-id="cbcba-113">Excel has a number of built-in error messages which are returned to a cell if there is calculation error.</span></span> <span data-ttu-id="cbcba-114">カスタム関数では、`#NULL!`、`#DIV/0!`、`#VALUE!`、`#REF!`、`#NAME?`、`#NUM!`、`#N/A`、`#BUSY!` の各エラー メッセージのみが使用されます。</span><span class="sxs-lookup"><span data-stu-id="cbcba-114">Custom functions only use the following error messages: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, and `#BUSY!`.</span></span>

<span data-ttu-id="cbcba-115">通常、これらのエラーは、あなたがExcelで既によく見たことがあるかもしれないエラーと対応関係があります。</span><span class="sxs-lookup"><span data-stu-id="cbcba-115">Generally, these errors correspond to the errors you might already be familiar with in Excel.</span></span> <span data-ttu-id="cbcba-116">カスタム関数に固有の例外はわずかにあります。以下に記載します。</span><span class="sxs-lookup"><span data-stu-id="cbcba-116">The are only a few exceptions specific to custom functions, listed here:</span></span>

- <span data-ttu-id="cbcba-117">`#NAME`エラーは通常、関数の登録に問題があることを意味します。</span><span class="sxs-lookup"><span data-stu-id="cbcba-117">A `#NAME` error generally means there has been an issue registering your functions.</span></span>
- <span data-ttu-id="cbcba-118">`#VALUE`エラーは通常、関数のスクリプトファイル内のエラーを示します。</span><span class="sxs-lookup"><span data-stu-id="cbcba-118">A `#VALUE` error typically indicates an error in the functions' script file.</span></span>
- <span data-ttu-id="cbcba-119">`#N/A`エラーは、登録されている間にその機能を実行できなかったということを示す可能性もあります。</span><span class="sxs-lookup"><span data-stu-id="cbcba-119">A `#N/A` error is also maybe a sign that that function while registered could not be run.</span></span> <span data-ttu-id="cbcba-120">この多くは、`CustomFunctions.associate`コマンドが欠落していることが原因です。</span><span class="sxs-lookup"><span data-stu-id="cbcba-120">This is typically due to a missing `CustomFunctions.associate` command.</span></span>
- <span data-ttu-id="cbcba-121">`#REF!`エラーは、関数名がアドイン内に既に存在するの関数名と同じであることを示している可能性があります。</span><span class="sxs-lookup"><span data-stu-id="cbcba-121">A `#REF!` error may indicate that your function name is the same as a function name in an add-in that already exists.</span></span>

## <a name="clear-the-office-cache"></a><span data-ttu-id="cbcba-122">Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="cbcba-122">Clear the Office cache</span></span>

<span data-ttu-id="cbcba-123">カスタム関数に関する情報はOfficeによってキャッシュされます。</span><span class="sxs-lookup"><span data-stu-id="cbcba-123">Information about custom functions is cached by Office.</span></span> <span data-ttu-id="cbcba-124">開発中、またカスタム関数を使用して繰り返しリロードしている間は、変更が反映されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="cbcba-124">Sometimes while developing and repeatedly reloading an add-in with custom functions your changes may not appear.</span></span> <span data-ttu-id="cbcba-125">Officeのキャッシュをクリアすることでこれを修正できます。</span><span class="sxs-lookup"><span data-stu-id="cbcba-125">You can fix this by clearing the Office cache.</span></span> <span data-ttu-id="cbcba-126">詳細については、記事[マニフェストの問題を検証、問題解決する](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache)内「Officeキャッシュをクリアする」の部分を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cbcba-126">For more information, see the "Clear the Office cache" section in the article [Validate and troubleshoot issues with your manifest](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache)</span></span>

## <a name="common-issues"></a><span data-ttu-id="cbcba-127">一般的な問題</span><span class="sxs-lookup"><span data-stu-id="cbcba-127">Common issues</span></span>

### <a name="my-functions-wont-load-associate-functions"></a><span data-ttu-id="cbcba-128">関数が読み込まれない: 関数を関連付ける</span><span class="sxs-lookup"><span data-stu-id="cbcba-128">My functions won't load: associate functions</span></span>

<span data-ttu-id="cbcba-129">カスタム関数のスクリプト ファイルで、各カスタム関数を、[JSON メタデータ ファイル](custom-functions-json.md)で指定されている ID と関連付ける必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbcba-129">In your custom functions' script file, you need to associate each custom function with its ID specified in the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="cbcba-130">これを行うには、`CustomFunctions.associate()` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="cbcba-130">This is done by using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="cbcba-131">通常、このメソッドの呼び出しは、各関数の後またはスクリプト ファイルの最後に行います。</span><span class="sxs-lookup"><span data-stu-id="cbcba-131">Typically this method call is made after each function or at the end of the script file.</span></span> <span data-ttu-id="cbcba-132">カスタム関数を関連付けないと、カスタム関数は機能しません。</span><span class="sxs-lookup"><span data-stu-id="cbcba-132">If a custom function is not associated, it will not work.</span></span>

<span data-ttu-id="cbcba-133">次の例では、add 関数の後で、関数の名前 `add` と対応する JSON ID `ADD` を関連付けています。</span><span class="sxs-lookup"><span data-stu-id="cbcba-133">The following example shows an add function, followed by the function's name `add` being associated with the corresponding JSON id `ADD`.</span></span>

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="cbcba-134">このプロセスの詳細については、「[関数名を JSON メタデータに関連付ける](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="cbcba-134">For more information on this process, see [Associating function names with json metadata](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).</span></span>

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a><span data-ttu-id="cbcba-135">localhostからアドインを開くことができません：ローカルループバック例外を使用してください</span><span class="sxs-lookup"><span data-stu-id="cbcba-135">Can't open add-in from localhost: use a local loopback exception</span></span>

<span data-ttu-id="cbcba-136">"We can't open this add-in from localhost"というエラーが表示された場合は、ローカルループバック例外を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbcba-136">If you see the error "We can't open this add-in from localhost," you will need to enable a local loopback exception.</span></span> <span data-ttu-id="cbcba-137">方法の詳細については、[このMicrosoft サポート記事](https://support.microsoft.com/ja-JP/help/4490419/local-loopback-exemption-does-not-work)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cbcba-137">For details on how to do this, see [this Microsoft support article](https://support.microsoft.com/ja-JP/help/4490419/local-loopback-exemption-does-not-work).</span></span>

### <a name="ensure-promises-return"></a><span data-ttu-id="cbcba-138">promise の戻り値を確認する</span><span class="sxs-lookup"><span data-stu-id="cbcba-138">Ensure promises return</span></span>

<span data-ttu-id="cbcba-139">Excelがカスタム関数の完了を待っている間、＃BUSY！と表示されます</span><span class="sxs-lookup"><span data-stu-id="cbcba-139">When Excel is waiting for a custom function to complete, it displays #BUSY!</span></span> <span data-ttu-id="cbcba-140">セル内に。</span><span class="sxs-lookup"><span data-stu-id="cbcba-140">in the cell.</span></span> <span data-ttu-id="cbcba-141">カスタム関数のコードで promise が返されているのに、promise で結果が返されない場合、Excel は #BUSY! を表示し続けます。</span><span class="sxs-lookup"><span data-stu-id="cbcba-141">If your custom function code returns a promise, but the promise does not return a result, Excel will continue showing #BUSY!.</span></span> <span data-ttu-id="cbcba-142">すべての promise でセルに結果が正しく返されていることを、関数で確認します。</span><span class="sxs-lookup"><span data-stu-id="cbcba-142">Check your functions to make sure that any promises are properly returning a result to a cell.</span></span>

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a><span data-ttu-id="cbcba-143">エラー：開発サーバーはすでにポート3000で実行されています。</span><span class="sxs-lookup"><span data-stu-id="cbcba-143">Error: The dev server is already running on port 3000</span></span>

<span data-ttu-id="cbcba-144">`npm start`を実行しているときに、開発サーバーが既にポート3000（またはアドインが使用しているポート）で実行されているというエラーが表示されることがあります。</span><span class="sxs-lookup"><span data-stu-id="cbcba-144">Sometimes when running `npm start` you may see an error that the dev server is already running on port 3000 (or whichever port your add-in uses).</span></span> <span data-ttu-id="cbcba-145">`npm stop`を実行するか、Node.jsウィンドウを閉じることによって、開発サーバーを停止できます。</span><span class="sxs-lookup"><span data-stu-id="cbcba-145">You can stop the dev server by running `npm stop` or by closing the Node.js window.</span></span> <span data-ttu-id="cbcba-146">しかし場合によっては、開発サーバーが実際に実行を停止するのに数分かかることがあります。</span><span class="sxs-lookup"><span data-stu-id="cbcba-146">But in some cases in can take a few minutes for the dev server to actually stop running.</span></span>

## <a name="reporting-feedback"></a><span data-ttu-id="cbcba-147">フィードバックの報告</span><span class="sxs-lookup"><span data-stu-id="cbcba-147">Reporting Feedback</span></span>

<span data-ttu-id="cbcba-148">ここに記載されていない問題が発生している場合は、お知らせください。</span><span class="sxs-lookup"><span data-stu-id="cbcba-148">If you are encountering issues that aren't documented here, let us know.</span></span> <span data-ttu-id="cbcba-149">問題を報告するには 2 つの方法があります。</span><span class="sxs-lookup"><span data-stu-id="cbcba-149">There are two ways to report issues.</span></span>

### <a name="in-excel-on-windows-or-mac"></a><span data-ttu-id="cbcba-150">Windows または Mac の Excel で</span><span class="sxs-lookup"><span data-stu-id="cbcba-150">In Excel on Windows or Mac</span></span>

<span data-ttu-id="cbcba-151">Windows 用または Mac 用の Excel を使用している場合は、Excel から Office の機能拡張チームにフィードバックを直接報告できます。</span><span class="sxs-lookup"><span data-stu-id="cbcba-151">If using Excel for Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="cbcba-152">これを行うには、**[ファイル]、[フィードバック]、[問題点、改善点の報告]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="cbcba-152">To do this, select **File -> Feedback -> Send a Frown**.</span></span> <span data-ttu-id="cbcba-153">問題点や改善点の報告では、発生した問題を理解するために必要なログが提供されます。</span><span class="sxs-lookup"><span data-stu-id="cbcba-153">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

### <a name="in-github"></a><span data-ttu-id="cbcba-154">GitHub で</span><span class="sxs-lookup"><span data-stu-id="cbcba-154">In Github</span></span>

<span data-ttu-id="cbcba-155">ドキュメント ページの下部にある "コンテンツ フィードバック" 機能を使用するか、[カスタム関数リポジトリに直接新しい問題を記入](https://github.com/OfficeDev/Excel-Custom-Functions/issues)して、発生した問題をお気軽に送信してください。</span><span class="sxs-lookup"><span data-stu-id="cbcba-155">Feel free to submit an issue you encounter either through the "Content feedback" feature at the bottom of any documentation page, or by [filing a new issue directly to the custom functions repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="cbcba-156">次の手順</span><span class="sxs-lookup"><span data-stu-id="cbcba-156">Next steps</span></span>
<span data-ttu-id="cbcba-157">[カスタム関数をデバッグする](custom-functions-debugging.md)手順をご参照ください。</span><span class="sxs-lookup"><span data-stu-id="cbcba-157">Learn how to [debug your custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="cbcba-158">関連項目</span><span class="sxs-lookup"><span data-stu-id="cbcba-158">See also</span></span>

* [<span data-ttu-id="cbcba-159">カスタム関数メタデータ自動生成</span><span class="sxs-lookup"><span data-stu-id="cbcba-159">Custom functions metadata</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="cbcba-160">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="cbcba-160">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="cbcba-161">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="cbcba-161">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* <span data-ttu-id="cbcba-162">[カスタム関数をXLLユーザー定義関数と互換性のあるものにします](make-custom-functions-compatible-with-xll-udf.md)。</span><span class="sxs-lookup"><span data-stu-id="cbcba-162">[Make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md)</span></span>
* [<span data-ttu-id="cbcba-163">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="cbcba-163">Create custom functions in Excel</span></span>](custom-functions-overview.md)
