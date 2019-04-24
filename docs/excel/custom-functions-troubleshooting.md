---
ms.date: 04/18/2019
description: Excel のカスタム関数で一般的な問題をトラブルシューティングします。
title: カスタム関数のトラブルシューティング (プレビュー)
localization_priority: Priority
ms.openlocfilehash: cf54aa3b719b7893799df5d1c5206c6fb904be69
ms.sourcegitcommit: 44c61926d35809152cbd48f7b97feb694c7fa3de
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/22/2019
ms.locfileid: "31959105"
---
# <a name="troubleshoot-custom-functions"></a><span data-ttu-id="3e124-103">カスタム関数のトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="3e124-103">Troubleshoot custom functions</span></span>

<span data-ttu-id="3e124-104">カスタム関数を作成してテストするとき、製品でエラーが発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="3e124-104">When developing custom functions, you may encounter errors in the product while creating and testing your functions.</span></span>

<span data-ttu-id="3e124-105">問題を解決するには、[ランタイム ログを有効にしてエラーをキャプチャ](#enable-runtime-logging)し、[Excel のネイティブ エラー メッセージ](#check-for-excel-error-messages)を参照します。</span><span class="sxs-lookup"><span data-stu-id="3e124-105">To resolve issues, you can [enable runtime logging to capture errors](#enable-runtime-logging) and refer to [Excel's native error messages](#check-for-excel-error-messages).</span></span> <span data-ttu-id="3e124-106">また、[SSL 証明書の検証](#my-add-in-wont-load-verify-certificates)を正しく行っていない、[promises を未解決のままにしている](#ensure-promises-return)、[関数の関連付け](#my-functions-wont-load-associate-functions)を忘れる、などの一般的な誤りを確認します。</span><span class="sxs-lookup"><span data-stu-id="3e124-106">Also, check for common mistakes such as not [verifying ssl certificates](#my-add-in-wont-load-verify-certificates) properly, [leaving promises unresolved](#ensure-promises-return), and forgetting to [associate your functions](#my-functions-wont-load-associate-functions).</span></span>

## <a name="enable-runtime-logging"></a><span data-ttu-id="3e124-107">ランタイム ログを有効にする</span><span class="sxs-lookup"><span data-stu-id="3e124-107">Enable runtime logging</span></span>

<span data-ttu-id="3e124-108">Windows 上の Office でアドインをテストする場合は、[ランタイム ログを有効にする](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)必要があります。</span><span class="sxs-lookup"><span data-stu-id="3e124-108">If you are testing your add-in in Office on Windows, you should [enable runtime logging](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span> <span data-ttu-id="3e124-109">ランタイム ログでは、問題解明用に別に作成したログ ファイルに `console.log` ステートメントが配信されます。</span><span class="sxs-lookup"><span data-stu-id="3e124-109">Runtime logging delivers `console.log` statements to a separate log file you create to help you uncover issues.</span></span> <span data-ttu-id="3e124-110">ステートメントでは、アドインの XML マニフェスト ファイルに関するエラー、実行時の条件、カスタム関数のインストールなど、さまざまなエラーがカバーされます。</span><span class="sxs-lookup"><span data-stu-id="3e124-110">The statements cover a variety of errors, including errors pertaining to your add-in's XML manifest file, runtime conditions, or installation of your custom functions.</span></span>  <span data-ttu-id="3e124-111">ランタイム ログの詳細については、「[アドインのデバッグにランタイム ログを使用する](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="3e124-111">For more information about runtime logging, see [Use runtime logging to debug your add-in](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span>  

### <a name="check-for-excel-error-messages"></a><span data-ttu-id="3e124-112">Excel のエラー メッセージを確認する</span><span class="sxs-lookup"><span data-stu-id="3e124-112">Check for Excel error messages</span></span>

<span data-ttu-id="3e124-113">Excel には多くの組み込みエラー メッセージがあり、計算エラーが発生するとセルに返されます。</span><span class="sxs-lookup"><span data-stu-id="3e124-113">Excel has a number of built-in error messages which are returned to a cell if there is calculation error.</span></span> <span data-ttu-id="3e124-114">カスタム関数では、`#NULL!`、`#DIV/0!`、`#VALUE!`、`#REF!`、`#NAME?`、`#NUM!`、`#N/A`、`#BUSY!` の各エラー メッセージのみが使用されます。</span><span class="sxs-lookup"><span data-stu-id="3e124-114">Custom functions only use the following error messages: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, and `#BUSY!`.</span></span>

## <a name="common-issues"></a><span data-ttu-id="3e124-115">一般的な問題</span><span class="sxs-lookup"><span data-stu-id="3e124-115">Common issues</span></span>

### <a name="my-add-in-wont-load-verify-certificates"></a><span data-ttu-id="3e124-116">アドインが読み込まれない: 証明書を確認する</span><span class="sxs-lookup"><span data-stu-id="3e124-116">My add-in won't load: verify certificates</span></span>

<span data-ttu-id="3e124-117">アドインのインストールが失敗する場合は、アドインをホストしている Web サーバーに対して SSL 証明書が正しく構成されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="3e124-117">If your add-in fails to install, verify that the SSL certificates are configured correctly for the web server that's hosting your add-in.</span></span> <span data-ttu-id="3e124-118">通常、SSL 証明書に問題がある場合は、アドインを正しくインストールできなかったことを警告する Excel のエラー メッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="3e124-118">Typically if there is a problem with SSL certificates, you will see an error message in Excel warning you that your add-in could not be installed properly.</span></span> <span data-ttu-id="3e124-119">詳細については、「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="3e124-119">For more information, see [Adding self-signed certificates as trusted root certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>

### <a name="my-functions-wont-load-associate-functions"></a><span data-ttu-id="3e124-120">関数が読み込まれない: 関数を関連付ける</span><span class="sxs-lookup"><span data-stu-id="3e124-120">My functions won't load: associate functions</span></span>

<span data-ttu-id="3e124-121">カスタム関数のスクリプト ファイルで、各カスタム関数を、[JSON メタデータ ファイル](custom-functions-json.md)で指定されている ID と関連付ける必要があります。</span><span class="sxs-lookup"><span data-stu-id="3e124-121">In your custom functions' script file, you need to associate each custom function with its ID specified in the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="3e124-122">これを行うには、`CustomFunctions.associate()` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="3e124-122">This is done by using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="3e124-123">通常、このメソッドの呼び出しは、各関数の後またはスクリプト ファイルの最後に行います。</span><span class="sxs-lookup"><span data-stu-id="3e124-123">Typically this method call is made after each function or at the end of the script file.</span></span> <span data-ttu-id="3e124-124">カスタム関数を関連付けないと、カスタム関数は機能しません。</span><span class="sxs-lookup"><span data-stu-id="3e124-124">If a custom function is not associated, it will not work.</span></span>

<span data-ttu-id="3e124-125">次の例では、add 関数の後で、関数の名前 `add` と対応する JSON ID `ADD` を関連付けています。</span><span class="sxs-lookup"><span data-stu-id="3e124-125">The following example shows an add function, followed by the function's name `add` being associated with the corresponding JSON id `ADD`.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="3e124-126">このプロセスの詳細については、「[関数名を JSON メタデータに関連付ける](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="3e124-126">For more information on this process, see [Associating function names with json metadata](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).</span></span>

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a><span data-ttu-id="3e124-127">localhostからアドインを開くことができません：ローカルループバック例外を使用してください</span><span class="sxs-lookup"><span data-stu-id="3e124-127">Can't open add-in from localhost: use a local loopback exception</span></span>

<span data-ttu-id="3e124-128">"We can't open this add-in from localhost"というエラーが表示された場合は、ローカルループバック例外を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="3e124-128">If you see the error "We can't open this add-in from localhost," you will need to enable a local loopback exception.</span></span> <span data-ttu-id="3e124-129">方法の詳細については、[このMicrosoft サポート記事](https://support.microsoft.com/ja-JP/help/4490419/local-loopback-exemption-does-not-work)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3e124-129">For details on how to do this, see [this Microsoft support article](https://support.microsoft.com/ja-JP/help/4490419/local-loopback-exemption-does-not-work).</span></span>

### <a name="ensure-promises-return"></a><span data-ttu-id="3e124-130">promise の戻り値を確認する</span><span class="sxs-lookup"><span data-stu-id="3e124-130">Ensure promises return</span></span>

<span data-ttu-id="3e124-131">Excelがカスタム関数の完了を待っている間、＃BUSY！と表示されます</span><span class="sxs-lookup"><span data-stu-id="3e124-131">When Excel is waiting for a custom function to complete, it displays #BUSY!</span></span> <span data-ttu-id="3e124-132">セル内に。</span><span class="sxs-lookup"><span data-stu-id="3e124-132">in the cell.</span></span> <span data-ttu-id="3e124-133">カスタム関数のコードで promise が返されているのに、promise で結果が返されない場合、Excel は #BUSY! を表示し続けます。</span><span class="sxs-lookup"><span data-stu-id="3e124-133">If your custom function code returns a promise, but the promise does not return a result, Excel will continue showing #BUSY!.</span></span> <span data-ttu-id="3e124-134">すべての promise でセルに結果が正しく返されていることを、関数で確認します。</span><span class="sxs-lookup"><span data-stu-id="3e124-134">Check your functions to make sure that any promises are properly returning a result to a cell.</span></span>

## <a name="reporting-feedback"></a><span data-ttu-id="3e124-135">フィードバックの報告</span><span class="sxs-lookup"><span data-stu-id="3e124-135">Reporting Feedback</span></span>

<span data-ttu-id="3e124-136">ここに記載されていない問題が発生している場合は、お知らせください。</span><span class="sxs-lookup"><span data-stu-id="3e124-136">If you are encountering issues that aren't documented here, let us know.</span></span> <span data-ttu-id="3e124-137">問題を報告するには 2 つの方法があります。</span><span class="sxs-lookup"><span data-stu-id="3e124-137">There are two ways to report issues.</span></span>

### <a name="in-excel-on-windows-or-mac"></a><span data-ttu-id="3e124-138">Windows または Mac の Excel で</span><span class="sxs-lookup"><span data-stu-id="3e124-138">In Excel on Windows or Mac</span></span>

<span data-ttu-id="3e124-139">Windows 用または Mac 用の Excel を使用している場合は、Excel から Office の機能拡張チームにフィードバックを直接報告できます。</span><span class="sxs-lookup"><span data-stu-id="3e124-139">If using Excel for Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="3e124-140">これを行うには、**[ファイル]、[フィードバック]、[問題点、改善点の報告]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="3e124-140">To do this, select **File -> Feedback -> Send a Frown**.</span></span> <span data-ttu-id="3e124-141">問題点や改善点の報告では、発生した問題を理解するために必要なログが提供されます。</span><span class="sxs-lookup"><span data-stu-id="3e124-141">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

### <a name="in-github"></a><span data-ttu-id="3e124-142">GitHub で</span><span class="sxs-lookup"><span data-stu-id="3e124-142">In Github</span></span>

<span data-ttu-id="3e124-143">ドキュメント ページの下部にある "コンテンツ フィードバック" 機能を使用するか、[カスタム関数リポジトリに直接新しい問題を記入](https://github.com/OfficeDev/Excel-Custom-Functions/issues)して、発生した問題をお気軽に送信してください。</span><span class="sxs-lookup"><span data-stu-id="3e124-143">Feel free to submit an issue you encounter either through the "Content feedback" feature at the bottom of any documentation page, or by [filing a new issue directly to the custom functions repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="see-also"></a><span data-ttu-id="3e124-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="3e124-144">See also</span></span>

* [<span data-ttu-id="3e124-145">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="3e124-145">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="3e124-146">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="3e124-146">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="3e124-147">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="3e124-147">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="3e124-148">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="3e124-148">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="3e124-149">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="3e124-149">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
