---
ms.date: 12/31/2019
description: Excel カスタム関数に関する一般的な問題のトラブルシューティングを行います。
title: カスタム関数のトラブルシューティング
localization_priority: Normal
ms.openlocfilehash: bc8a450b1436b487f2c2a77e191182c540f55923
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719610"
---
# <a name="troubleshoot-custom-functions"></a><span data-ttu-id="cc3cf-103">カスタム関数のトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="cc3cf-103">Troubleshoot custom functions</span></span>

<span data-ttu-id="cc3cf-104">カスタム関数を作成してテストするとき、製品でエラーが発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-104">When developing custom functions, you may encounter errors in the product while creating and testing your functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="cc3cf-105">問題を解決するには、[ランタイム ログを有効にしてエラーをキャプチャ](#enable-runtime-logging)し、[Excel のネイティブ エラー メッセージ](#check-for-excel-error-messages)を参照します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-105">To resolve issues, you can [enable runtime logging to capture errors](#enable-runtime-logging) and refer to [Excel's native error messages](#check-for-excel-error-messages).</span></span> <span data-ttu-id="cc3cf-106">また、[promise を未解決のままにしておく](#ensure-promises-return)など、よくある間違いがないか確認します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-106">Also, check for common mistakes such as [leaving promises unresolved](#ensure-promises-return).</span></span>

## <a name="enable-runtime-logging"></a><span data-ttu-id="cc3cf-107">ランタイム ログを有効にする</span><span class="sxs-lookup"><span data-stu-id="cc3cf-107">Enable runtime logging</span></span>

<span data-ttu-id="cc3cf-108">Windows 上の Office でアドインをテストする場合は、[ランタイム ログを有効にする](../testing/runtime-logging.md)必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-108">If you're testing your add-in in Office on Windows, you should [enable runtime logging](../testing/runtime-logging.md).</span></span> <span data-ttu-id="cc3cf-109">ランタイム ログでは、問題解明用に別に作成したログ ファイルに `console.log` ステートメントが配信されます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-109">Runtime logging delivers `console.log` statements to a separate log file you create to help you uncover issues.</span></span> <span data-ttu-id="cc3cf-110">ステートメントでは、アドインの XML マニフェスト ファイルに関するエラー、実行時の条件、カスタム関数のインストールなど、さまざまなエラーがカバーされます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-110">The statements cover a variety of errors, including errors pertaining to your add-in's XML manifest file, runtime conditions, or installation of your custom functions.</span></span> <span data-ttu-id="cc3cf-111">ランタイム ログの詳細については、「[ランタイム ログを使用してアドインをデバッグする](../testing/runtime-logging.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-111">For more information about runtime logging, see [Debug your add-in with runtime logging](../testing/runtime-logging.md).</span></span>

### <a name="check-for-excel-error-messages"></a><span data-ttu-id="cc3cf-112">Excel のエラー メッセージを確認する</span><span class="sxs-lookup"><span data-stu-id="cc3cf-112">Check for Excel error messages</span></span>

<span data-ttu-id="cc3cf-113">Excel には多くの組み込みエラー メッセージがあり、計算エラーが発生するとセルに返されます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-113">Excel has a number of built-in error messages which are returned to a cell if there is calculation error.</span></span> <span data-ttu-id="cc3cf-114">カスタム関数では、`#NULL!`、`#DIV/0!`、`#VALUE!`、`#REF!`、`#NAME?`、`#NUM!`、`#N/A`、`#BUSY!` の各エラー メッセージのみが使用されます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-114">Custom functions only use the following error messages: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, and `#BUSY!`.</span></span>

<span data-ttu-id="cc3cf-115">通常、これらのエラーは、あなたがExcelで既によく見たことがあるかもしれないエラーと対応関係があります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-115">Generally, these errors correspond to the errors you might already be familiar with in Excel.</span></span> <span data-ttu-id="cc3cf-116">カスタム関数に固有の例外はわずかにあります。以下に記載します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-116">The are only a few exceptions specific to custom functions, listed here:</span></span>

- <span data-ttu-id="cc3cf-117">`#NAME`エラーは通常、関数の登録に問題があることを意味します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-117">A `#NAME` error generally means there has been an issue registering your functions.</span></span>
- <span data-ttu-id="cc3cf-118">`#N/A`エラーは、登録されている間にその機能を実行できなかったということを示す可能性もあります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-118">A `#N/A` error is also maybe a sign that that function while registered could not be run.</span></span> <span data-ttu-id="cc3cf-119">この多くは、`CustomFunctions.associate`コマンドが欠落していることが原因です。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-119">This is typically due to a missing `CustomFunctions.associate` command.</span></span>
- <span data-ttu-id="cc3cf-120">`#VALUE`エラーは通常、関数のスクリプトファイル内のエラーを示します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-120">A `#VALUE` error typically indicates an error in the functions' script file.</span></span>
- <span data-ttu-id="cc3cf-121">`#REF!`エラーは、関数名がアドイン内に既に存在するの関数名と同じであることを示している可能性があります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-121">A `#REF!` error may indicate that your function name is the same as a function name in an add-in that already exists.</span></span>

## <a name="clear-the-office-cache"></a><span data-ttu-id="cc3cf-122">Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="cc3cf-122">Clear the Office cache</span></span>

<span data-ttu-id="cc3cf-123">カスタム関数に関する情報はOfficeによってキャッシュされます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-123">Information about custom functions is cached by Office.</span></span> <span data-ttu-id="cc3cf-124">開発中、またカスタム関数を使用して繰り返しリロードしている間は、変更が反映されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-124">Sometimes while developing and repeatedly reloading an add-in with custom functions your changes may not appear.</span></span> <span data-ttu-id="cc3cf-125">Officeのキャッシュをクリアすることでこれを修正できます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-125">You can fix this by clearing the Office cache.</span></span> <span data-ttu-id="cc3cf-126">詳細については、「[Office のキャッシュをクリアする](../testing/clear-cache.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-126">For more information, see [Clear the Office cache](../testing/clear-cache.md).</span></span>

## <a name="common-issues"></a><span data-ttu-id="cc3cf-127">一般的な問題</span><span class="sxs-lookup"><span data-stu-id="cc3cf-127">Common issues</span></span>

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a><span data-ttu-id="cc3cf-128">localhostからアドインを開くことができません：ローカルループバック例外を使用してください</span><span class="sxs-lookup"><span data-stu-id="cc3cf-128">Can't open add-in from localhost: use a local loopback exception</span></span>

<span data-ttu-id="cc3cf-129">"We can't open this add-in from localhost"というエラーが表示された場合は、ローカルループバック例外を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-129">If you see the error "We can't open this add-in from localhost," you will need to enable a local loopback exception.</span></span> <span data-ttu-id="cc3cf-130">方法の詳細については、[このMicrosoft サポート記事](https://support.microsoft.com/help/4490419/local-loopback-exemption-does-not-work)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-130">For details on how to do this, see [this Microsoft support article](https://support.microsoft.com/help/4490419/local-loopback-exemption-does-not-work).</span></span>

### <a name="runtime-logging-reports-typeerror-network-request-failed-on-excel-on-windows"></a><span data-ttu-id="cc3cf-131">Windows 上の Excel でランタイム ログが「TypeError: Network request failed」と報告する</span><span class="sxs-lookup"><span data-stu-id="cc3cf-131">Runtime logging reports "TypeError: Network request failed" on Excel on Windows</span></span>

<span data-ttu-id="cc3cf-132">localhost サーバーへの呼び出し中に[ランタイム ログ](custom-functions-troubleshooting.md#enable-runtime-logging)に「TypeError: Network request failed」というエラーが表示された場合は、ローカル ループバック例外を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-132">If you see the error "TypeError: Network request failed" in your [runtime log](custom-functions-troubleshooting.md#enable-runtime-logging) while making calls to your localhost server, you'll need to enable a local loopback exception.</span></span> <span data-ttu-id="cc3cf-133">方法の詳細については、[このMicrosoft サポート記事](https://support.microsoft.com/help/4490419/local-loopback-exemption-does-not-work)の*オプション 2* を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-133">For details on how to do this, see *Option #2* in [this Microsoft support article](https://support.microsoft.com/help/4490419/local-loopback-exemption-does-not-work).</span></span>

### <a name="ensure-promises-return"></a><span data-ttu-id="cc3cf-134">promise の戻り値を確認する</span><span class="sxs-lookup"><span data-stu-id="cc3cf-134">Ensure promises return</span></span>

<span data-ttu-id="cc3cf-135">Excelがカスタム関数の完了を待っている間、＃BUSY！と表示されます</span><span class="sxs-lookup"><span data-stu-id="cc3cf-135">When Excel is waiting for a custom function to complete, it displays #BUSY!</span></span> <span data-ttu-id="cc3cf-136">セル内に。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-136">in the cell.</span></span> <span data-ttu-id="cc3cf-137">カスタム関数のコードで promise が返されているのに、promise で結果が返されない場合、Excel は `#BUSY!` を表示し続けます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-137">If your custom function code returns a promise, but the promise does not return a result, Excel will continue showing `#BUSY!`.</span></span> <span data-ttu-id="cc3cf-138">すべての promise でセルに結果が正しく返されていることを、関数で確認します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-138">Check your functions to make sure that any promises are properly returning a result to a cell.</span></span>

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a><span data-ttu-id="cc3cf-139">エラー：開発サーバーはすでにポート3000で実行されています。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-139">Error: The dev server is already running on port 3000</span></span>

<span data-ttu-id="cc3cf-140">`npm start`を実行しているときに、開発サーバーが既にポート3000（またはアドインが使用しているポート）で実行されているというエラーが表示されることがあります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-140">Sometimes when running `npm start` you may see an error that the dev server is already running on port 3000 (or whichever port your add-in uses).</span></span> <span data-ttu-id="cc3cf-141">`npm stop`を実行するか、Node.jsウィンドウを閉じることによって、開発サーバーを停止できます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-141">You can stop the dev server by running `npm stop` or by closing the Node.js window.</span></span> <span data-ttu-id="cc3cf-142">しかし場合によっては、開発サーバーが実際に実行を停止するのに数分かかることがあります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-142">But in some cases in can take a few minutes for the dev server to actually stop running.</span></span>

### <a name="my-functions-wont-load-associate-functions"></a><span data-ttu-id="cc3cf-143">関数が読み込まれない: 関数を関連付ける</span><span class="sxs-lookup"><span data-stu-id="cc3cf-143">My functions won't load: associate functions</span></span>

<span data-ttu-id="cc3cf-144">JSON が登録されておらず、独自の JSON メタデータを作成した場合、`#VALUE!` エラーが表示されるか、アドインを読み込めないという通知が表示されます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-144">In cases where your JSON has not been registered and you have authored your own JSON metadata, you may see a `#VALUE!` error or receive a notification that your add-in cannot be loaded.</span></span> <span data-ttu-id="cc3cf-145">これは通常、各カスタム関数を [JSON メタデータ ファイル](custom-functions-json.md)で指定されている `id` プロパティと関連付ける必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-145">This usually means you need to associate each custom function with its `id` property specified in the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="cc3cf-146">これを行うには、`CustomFunctions.associate()` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-146">This is done by using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="cc3cf-147">通常、このメソッドの呼び出しは、各関数の後またはスクリプト ファイルの最後に行います。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-147">Typically this method call is made after each function or at the end of the script file.</span></span> <span data-ttu-id="cc3cf-148">カスタム関数を関連付けないと、カスタム関数は機能しません。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-148">If a custom function is not associated, it will not work.</span></span>

<span data-ttu-id="cc3cf-149">次の例では、add 関数の後で、関数の名前 `add` と対応する JSON ID `ADD` を関連付けています。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-149">The following example shows an add function, followed by the function's name `add` being associated with the corresponding JSON id `ADD`.</span></span>

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

<span data-ttu-id="cc3cf-150">このプロセスの詳細については、「[関数名を JSON メタデータに関連付ける](../excel/custom-functions-json.md#associating-function-names-with-json-metadata)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-150">For more information on this process, see [Associating function names with json metadata](../excel/custom-functions-json.md#associating-function-names-with-json-metadata).</span></span>

## <a name="reporting-feedback"></a><span data-ttu-id="cc3cf-151">フィードバックの報告</span><span class="sxs-lookup"><span data-stu-id="cc3cf-151">Reporting feedback</span></span>

<span data-ttu-id="cc3cf-152">ここに記載されていない問題が発生している場合は、お知らせください。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-152">If you are encountering issues that aren't documented here, let us know.</span></span> <span data-ttu-id="cc3cf-153">問題を報告するには 2 つの方法があります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-153">There are two ways to report issues.</span></span>

### <a name="in-excel-on-windows-or-mac"></a><span data-ttu-id="cc3cf-154">Windows または Mac の Excel で</span><span class="sxs-lookup"><span data-stu-id="cc3cf-154">In Excel on Windows or Mac</span></span>

<span data-ttu-id="cc3cf-155">Windows または Mac で Excel を使用している場合は、Excel から Office の機能拡張チームにフィードバックを直接報告できます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-155">If using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="cc3cf-156">これを行うには、**[ファイル]、[フィードバック]、[問題点、改善点の報告]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-156">To do this, select **File -> Feedback -> Send a Frown**.</span></span> <span data-ttu-id="cc3cf-157">問題点や改善点の報告では、発生した問題を理解するために必要なログが提供されます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-157">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

### <a name="in-github"></a><span data-ttu-id="cc3cf-158">GitHub で</span><span class="sxs-lookup"><span data-stu-id="cc3cf-158">In Github</span></span>

<span data-ttu-id="cc3cf-159">ドキュメント ページの下部にある "コンテンツ フィードバック" 機能を使用するか、[カスタム関数リポジトリに直接新しい問題を記入](https://github.com/OfficeDev/Excel-Custom-Functions/issues)して、発生した問題をお気軽に送信してください。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-159">Feel free to submit an issue you encounter either through the "Content feedback" feature at the bottom of any documentation page, or by [filing a new issue directly to the custom functions repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="cc3cf-160">次の手順</span><span class="sxs-lookup"><span data-stu-id="cc3cf-160">Next steps</span></span>
<span data-ttu-id="cc3cf-161">[カスタム関数をデバッグする](custom-functions-debugging.md)手順をご参照ください。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-161">Learn how to [debug your custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="cc3cf-162">関連項目</span><span class="sxs-lookup"><span data-stu-id="cc3cf-162">See also</span></span>

* [<span data-ttu-id="cc3cf-163">カスタム関数メタデータ自動生成</span><span class="sxs-lookup"><span data-stu-id="cc3cf-163">Custom functions metadata autogeneration</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="cc3cf-164">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="cc3cf-164">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="cc3cf-165">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-165">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="cc3cf-166">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="cc3cf-166">Create custom functions in Excel</span></span>](custom-functions-overview.md)
