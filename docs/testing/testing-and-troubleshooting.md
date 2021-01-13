---
title: Office アドインでのユーザー エラーのトラブルシューティング
description: アドインのユーザー エラーをトラブルシューティングするOffice説明します。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: e1cb5e0bb8005f04425a5ad9c7e807d10f054e35
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49840098"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="76900-103">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="76900-103">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="76900-p101">時折、ユーザーは開発した Office アドインの問題に遭遇することがあります。たとえば、アドインが読み込みに失敗したり、アクセスできないなどです。この記事の情報は、ユーザーが Office アドインを使用する際に遭遇する一般的な問題を解決するために用いることができます。</span><span class="sxs-lookup"><span data-stu-id="76900-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="76900-107">また、[Fiddler](https://www.telerik.com/fiddler) を使用して、アドインの問題を特定してデバッグすることもできます。</span><span class="sxs-lookup"><span data-stu-id="76900-107">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="76900-108">一般的なエラーとトラブルシューティングの手順</span><span class="sxs-lookup"><span data-stu-id="76900-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="76900-109">次の表は、ユーザーが遭遇する可能性がある一般的なエラー メッセージとエラーを解決するためにユーザーが実行できる手順を示しています。</span><span class="sxs-lookup"><span data-stu-id="76900-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="76900-110">**エラー メッセージ**</span><span class="sxs-lookup"><span data-stu-id="76900-110">**Error message**</span></span>|<span data-ttu-id="76900-111">**解決策**</span><span class="sxs-lookup"><span data-stu-id="76900-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="76900-112">アプリのエラー: カタログに到達できませんでした</span><span class="sxs-lookup"><span data-stu-id="76900-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="76900-p102">ファイアウォールの設定を確認します。「カタログ」は、AppSource を指します。このメッセージは、ユーザーが AppSource にアクセスできないことを示しています。</span><span class="sxs-lookup"><span data-stu-id="76900-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="76900-p103">アプリのエラー: このアプリを起動できませんでした。このダイアログを閉じて問題を無視するか、[再起動] をクリックしてもう一度お試しください。</span><span class="sxs-lookup"><span data-stu-id="76900-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="76900-117">Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム ](https://support.microsoft.com/kb/2986156/)をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="76900-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="76900-118">エラー: オブジェクトがプロパティまたはメソッド 'defineProperty' をサポートしていません</span><span class="sxs-lookup"><span data-stu-id="76900-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="76900-119">Internet Explorerが互換モードで実行されていないことを確認します。</span><span class="sxs-lookup"><span data-stu-id="76900-119">Confirm that Internet Explorer is not running in Compatibility Mode.</span></span> <span data-ttu-id="76900-120">[ツール] > **[互換表示設定]** に移動します。</span><span class="sxs-lookup"><span data-stu-id="76900-120">Go to Tools > **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="76900-p105">ブラウザーのバージョンがサポートされていないため、アプリを読み込めませんでした。サポートされているブラウザーのバージョンの一覧についてはここをクリックしてください。</span><span class="sxs-lookup"><span data-stu-id="76900-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="76900-p106">ブラウザーが HTML5 のローカル ストレージをサポートしていることを確認するか、Internet Explorer の設定をリセットします。サポートされているブラウザーの詳細については、「[Office アドインを実行するための要件](../concepts/requirements-for-running-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="76900-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a><span data-ttu-id="76900-125">アドインをインストールすると、ステータス バーに "アドイン読み込み中のエラー" と表示される</span><span class="sxs-lookup"><span data-stu-id="76900-125">When installing an add-in, you see "Error loading add-in" in the status bar</span></span>

1. <span data-ttu-id="76900-126">Office を終了します。</span><span class="sxs-lookup"><span data-stu-id="76900-126">Close Office.</span></span>
2. <span data-ttu-id="76900-127">マニフェストが有効であることを確認する</span><span class="sxs-lookup"><span data-stu-id="76900-127">Verify that the manifest is valid</span></span>
3. <span data-ttu-id="76900-128">アドインを再起動する</span><span class="sxs-lookup"><span data-stu-id="76900-128">Restart the add-in</span></span>
4. <span data-ttu-id="76900-129">もう一度アドインをインストールする。</span><span class="sxs-lookup"><span data-stu-id="76900-129">Install the add-in again.</span></span>

<span data-ttu-id="76900-130">また、フィードバックを寄せることができます。Windows または Mac 用 Excel を使用している場合は、Excel から直接 Office の機能拡張チームにフィードバックを送信できます。</span><span class="sxs-lookup"><span data-stu-id="76900-130">You can also give us feedback: if using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="76900-131">これを行うには、[**ファイル**] | [**フィードバック**] | [**問題点、改善の報告**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="76900-131">To do this, select **File** | **Feedback** | **Send a Frown**.</span></span> <span data-ttu-id="76900-132">問題点、改善の報告により、問題を理解するために必要なログが提供されます。</span><span class="sxs-lookup"><span data-stu-id="76900-132">Sending a frown provides the necessary logs to understand the issue.</span></span>

## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="76900-133">Outlook アドインが正常に機能しない</span><span class="sxs-lookup"><span data-stu-id="76900-133">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="76900-134">Windows で実行され、[Internet Explorer を使用している](../concepts/browsers-used-by-office-web-add-ins.md) Outlook アドインが正常に機能しない場合は、Internet Explorer でスクリプトのデバッグを有効にしてみてください。</span><span class="sxs-lookup"><span data-stu-id="76900-134">If an Outlook add-in running on Windows and [using Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="76900-135">Go to Tools > **Internet Options**  >  **Advanced**.</span><span class="sxs-lookup"><span data-stu-id="76900-135">Go to Tools > **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="76900-136">**[参照]** で、**[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオフにします。</span><span class="sxs-lookup"><span data-stu-id="76900-136">Under **Browsing**, uncheck **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="76900-p108">これらの設定は、問題のトラブルシューティングを行う場合にのみチェックボックスをオフにすることをお勧めします。チェックボックスをオフにしたままにすると、参照時にメッセージが表示されます。問題が解決したら、**[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオンにしてください。</span><span class="sxs-lookup"><span data-stu-id="76900-p108">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="76900-140">Office 2013 でアドインがアクティブにならない</span><span class="sxs-lookup"><span data-stu-id="76900-140">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="76900-141">ユーザーが次の手順を実行したときに、アドインがアクティブにならない場合があります。</span><span class="sxs-lookup"><span data-stu-id="76900-141">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="76900-142">Office 2013 で自分の Microsoft アカウントでサインインする。</span><span class="sxs-lookup"><span data-stu-id="76900-142">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="76900-143">自分の Microsoft アカウントの 2 段階検証を有効にする。</span><span class="sxs-lookup"><span data-stu-id="76900-143">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="76900-144">アドインを挿入しようとする際に、メッセージに従って ID の確認を行う。</span><span class="sxs-lookup"><span data-stu-id="76900-144">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="76900-145">Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム](https://support.microsoft.com/kb/2986156/)をダウンロードしてください。</span><span class="sxs-lookup"><span data-stu-id="76900-145">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>

## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="76900-146">アドイン ダイアログ ボックスを表示できない</span><span class="sxs-lookup"><span data-stu-id="76900-146">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="76900-p109">Office アドインを使用するとき、ユーザーは、ダイアログ ボックスの表示を許可するよう求められます。ユーザーが **[許可]** を選択すると、次のエラー メッセージが発生します。</span><span class="sxs-lookup"><span data-stu-id="76900-p109">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="76900-p110">"ブラウザーのセキュリティ設定により、ダイアログ ボックスを作成できませんでした。別のブラウザーを試すか、アドレス バーに表示される [URL] とドメインが同じセキュリティ ゾーンに存在するようにブラウザーを構成してください。"</span><span class="sxs-lookup"><span data-stu-id="76900-p110">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![ダイアログ ボックスのエラー メッセージのスクリーンショット](../images/dialog-prevented.png)

|<span data-ttu-id="76900-152">**影響を受けるブラウザー**</span><span class="sxs-lookup"><span data-stu-id="76900-152">**Affected browsers**</span></span>|<span data-ttu-id="76900-153">**影響を受けるプラットフォーム**</span><span class="sxs-lookup"><span data-stu-id="76900-153">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="76900-154">Internet Explorer、Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="76900-154">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="76900-155">Office on the web</span><span class="sxs-lookup"><span data-stu-id="76900-155">Office on the web</span></span>|

<span data-ttu-id="76900-p111">この問題を解決するために、エンド ユーザーまたは管理者は、Internet Explore の信頼済みサイトのリストにアドインのドメインを追加することができます。Internet Explorer または Microsoft Edge ブラウザーのどちらを使用していても、同じ手順を使用します。</span><span class="sxs-lookup"><span data-stu-id="76900-p111">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="76900-158">アドインを信頼しない場合は、信頼済みサイトのリストにアドインの URL を追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="76900-158">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="76900-159">URL を信頼済みサイトのリストに追加する方法:</span><span class="sxs-lookup"><span data-stu-id="76900-159">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="76900-160">[**コントロール パネル**]で、[**インターネット オプション**]  >  [**セキュリティ**] に移動します。</span><span class="sxs-lookup"><span data-stu-id="76900-160">In **Control Panel**, go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="76900-161">**[信頼済みサイト]** ゾーンを選択して、**[サイト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="76900-161">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="76900-162">エラー メッセージに表示される URL を入力して、**[追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="76900-162">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="76900-p112">アドインの使用をもう一度お試しください。問題が続く場合は、他のセキュリティ ゾーンの設定を変えて、アドインのドメインが Office アプリケーションのアドレス バーに表示される URL と同じゾーンに存在するようにします。</span><span class="sxs-lookup"><span data-stu-id="76900-p112">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="76900-p113">この問題は、ポップアップ モードでダイアログ API が使用されているときに発生します。この問題を防ぐには、[displayInFrame](/javascript/api/office/office.ui) フラグを使います。そのために、ページが iframe 内の表示をサポートしている必要があります。次の例は、フラグの使用方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="76900-p113">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="see-also"></a><span data-ttu-id="76900-169">関連項目</span><span class="sxs-lookup"><span data-stu-id="76900-169">See also</span></span>

- [<span data-ttu-id="76900-170">アドインを使用したOfficeエラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="76900-170">Troubleshoot development errors with Office Add-ins</span></span>](troubleshoot-development-errors.md)
