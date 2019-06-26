---
title: Office アドインでのユーザー エラーのトラブルシューティング
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: a82dc05789b4c35a954337a64197d3ac1a190b96
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126905"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="eeccf-102">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="eeccf-102">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="eeccf-p101">時折、ユーザーは開発した Office アドインの問題に遭遇することがあります。たとえば、アドインが読み込みに失敗したり、アクセスできないなどです。この記事の情報は、ユーザーが Office アドインを使用する際に遭遇する一般的な問題を解決するために用いることができます。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="eeccf-106">また、[Fiddler](https://www.telerik.com/fiddler) を使用して、アドインの問題を特定してデバッグすることもできます。</span><span class="sxs-lookup"><span data-stu-id="eeccf-106">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

<span data-ttu-id="eeccf-107">ユーザーの問題を解決した後、[AppSource でカスタマー レビューに直接返信することができます](/office/dev/store/create-effective-office-store-listings)。</span><span class="sxs-lookup"><span data-stu-id="eeccf-107">After you resolve the user's issue, you can [respond directly to customer reviews in AppSource](/office/dev/store/create-effective-office-store-listings).</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="eeccf-108">一般的なエラーとトラブルシューティングの手順</span><span class="sxs-lookup"><span data-stu-id="eeccf-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="eeccf-109">次の表は、ユーザーが遭遇する可能性がある一般的なエラー メッセージとエラーを解決するためにユーザーが実行できる手順を示しています。</span><span class="sxs-lookup"><span data-stu-id="eeccf-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="eeccf-110">**エラー メッセージ**</span><span class="sxs-lookup"><span data-stu-id="eeccf-110">**Error message**</span></span>|<span data-ttu-id="eeccf-111">**解決策**</span><span class="sxs-lookup"><span data-stu-id="eeccf-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="eeccf-112">アプリのエラー: カタログに到達できませんでした</span><span class="sxs-lookup"><span data-stu-id="eeccf-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="eeccf-p102">ファイアウォールの設定を確認します。「カタログ」は、AppSource を指します。このメッセージは、ユーザーが AppSource にアクセスできないことを示しています。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="eeccf-p103">アプリのエラー: このアプリを起動できませんでした。このダイアログを閉じて問題を無視するか、[再起動] をクリックしてもう一度お試しください。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="eeccf-117">Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム ](https://support.microsoft.com/kb/2986156/)をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="eeccf-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="eeccf-118">エラー: オブジェクトがプロパティまたはメソッド 'defineProperty' をサポートしていません</span><span class="sxs-lookup"><span data-stu-id="eeccf-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="eeccf-p104">Internet Explorerが互換モードで実行されていないことを確認します。 [ツール] >  **[互換表示設定]** に移動します。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p104">Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools >  **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="eeccf-p105">ブラウザーのバージョンがサポートされていないため、アプリを読み込めませんでした。サポートされているブラウザーのバージョンの一覧についてはここをクリックしてください。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="eeccf-p106">ブラウザーが HTML5 のローカル ストレージをサポートしていることを確認するか、Internet Explorer の設定をリセットします。サポートされているブラウザーの詳細については、「[Office アドインを実行するための要件](../concepts/requirements-for-running-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|


## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="eeccf-125">Outlook アドインが正常に機能しない</span><span class="sxs-lookup"><span data-stu-id="eeccf-125">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="eeccf-126">Windows で実行している Outlook アドインが正常に機能しない場合は、Internet Explorer でスクリプトのデバッグを有効にしてみてください。</span><span class="sxs-lookup"><span data-stu-id="eeccf-126">If an Outlook add-in running on Windows is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="eeccf-127">[ツール] >  **[インターネット オプション]** > **[詳細]** に移動します。</span><span class="sxs-lookup"><span data-stu-id="eeccf-127">Go to Tools >  **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="eeccf-128">**[参照]** で、 **[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオフにします。</span><span class="sxs-lookup"><span data-stu-id="eeccf-128">Under  **Browsing**, uncheck  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="eeccf-p107">これらの設定は、問題のトラブルシューティングを行う場合にのみチェックボックスをオフにすることをお勧めします。チェックボックスをオフにしたままにすると、参照時にメッセージが表示されます。問題が解決したら、 **[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオンにしてください。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p107">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="eeccf-132">Office 2013 でアドインがアクティブにならない</span><span class="sxs-lookup"><span data-stu-id="eeccf-132">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="eeccf-133">ユーザーが次の手順を実行したときに、アドインがアクティブにならない場合があります。</span><span class="sxs-lookup"><span data-stu-id="eeccf-133">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="eeccf-134">Office 2013 で自分の Microsoft アカウントでサインインする。</span><span class="sxs-lookup"><span data-stu-id="eeccf-134">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="eeccf-135">自分の Microsoft アカウントの 2 段階検証を有効にする。</span><span class="sxs-lookup"><span data-stu-id="eeccf-135">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="eeccf-136">アドインを挿入しようとする際に、メッセージに従って ID の確認を行う。</span><span class="sxs-lookup"><span data-stu-id="eeccf-136">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="eeccf-137">Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム](https://support.microsoft.com/kb/2986156/)をダウンロードしてください。</span><span class="sxs-lookup"><span data-stu-id="eeccf-137">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="eeccf-138">アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題</span><span class="sxs-lookup"><span data-stu-id="eeccf-138">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="eeccf-139">「[マニフェストの問題を検証し、トラブルシューティングする](troubleshoot-manifest.md)」を参照して、アドインのマニフェストの問題をデバッグしてください。</span><span class="sxs-lookup"><span data-stu-id="eeccf-139">See [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="eeccf-140">アドイン ダイアログ ボックスを表示できない</span><span class="sxs-lookup"><span data-stu-id="eeccf-140">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="eeccf-p108">Office アドインを使用するとき、ユーザーは、ダイアログ ボックスの表示を許可するよう求められます。ユーザーが **[許可]** を選択すると、次のエラー メッセージが発生します。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p108">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="eeccf-p109">"ブラウザーのセキュリティ設定により、ダイアログ ボックスを作成できませんでした。別のブラウザーを試すか、アドレス バーに表示される [URL] とドメインが同じセキュリティ ゾーンに存在するようにブラウザーを構成してください。"</span><span class="sxs-lookup"><span data-stu-id="eeccf-p109">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![ダイアログ ボックスのエラー メッセージのスクリーン ショット](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="eeccf-146">**影響を受けるブラウザー**</span><span class="sxs-lookup"><span data-stu-id="eeccf-146">**Affected browsers**</span></span>|<span data-ttu-id="eeccf-147">**影響を受けるプラットフォーム**</span><span class="sxs-lookup"><span data-stu-id="eeccf-147">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="eeccf-148">Internet Explorer、Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="eeccf-148">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="eeccf-149">Office on the web</span><span class="sxs-lookup"><span data-stu-id="eeccf-149">Office on the web</span></span>|

<span data-ttu-id="eeccf-p110">この問題を解決するために、エンド ユーザーまたは管理者は、Internet Explore の信頼済みサイトのリストにアドインのドメインを追加することができます。Internet Explorer または Microsoft Edge ブラウザーのどちらを使用していても、同じ手順を使用します。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p110">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="eeccf-152">アドインを信頼しない場合は、信頼済みサイトのリストにアドインの URL を追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="eeccf-152">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="eeccf-153">URL を信頼済みサイトのリストに追加する方法:</span><span class="sxs-lookup"><span data-stu-id="eeccf-153">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="eeccf-154">Internet Explorer で [ツール] ボタンを選択し、**[インター ネット オプション]** > **[セキュリティ]** へ移動します。</span><span class="sxs-lookup"><span data-stu-id="eeccf-154">In Internet Explorer, choose the Tools button, and go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="eeccf-155">**[信頼済みサイト]** ゾーンを選択して、**[サイト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="eeccf-155">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="eeccf-156">エラー メッセージに表示される URL を入力して、**[追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="eeccf-156">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="eeccf-p111">アドインの使用をもう一度お試しください。問題が続く場合は、他のセキュリティ ゾーンの設定を変えて、アドインのドメインが Office アプリケーションのアドレス バーに表示される URL と同じゾーンに存在するようにします。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p111">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="eeccf-p112">この問題は、ポップアップ モードでダイアログ API が使用されているときに発生します。この問題を防ぐには、[displayInFrame](/javascript/api/office/office.ui) フラグを使います。そのために、ページが iframe 内の表示をサポートしている必要があります。次の例は、フラグの使用方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p112">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="eeccf-163">リボン ボタンとメニュー項目が含まれているアドイン コマンドへの変更が反映されない</span><span class="sxs-lookup"><span data-stu-id="eeccf-163">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="eeccf-164">リボン ボタンのアイコンのファイル名やメニュー アイテムのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。</span><span class="sxs-lookup"><span data-stu-id="eeccf-164">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="eeccf-165">Windows の場合:</span><span class="sxs-lookup"><span data-stu-id="eeccf-165">For Windows:</span></span>
<span data-ttu-id="eeccf-166">フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` の内容を削除する</span><span class="sxs-lookup"><span data-stu-id="eeccf-166">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="eeccf-167">Mac の場合: </span><span class="sxs-lookup"><span data-stu-id="eeccf-167">For Mac:</span></span>
<span data-ttu-id="eeccf-168">フォルダー `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` の内容を削除する</span><span class="sxs-lookup"><span data-stu-id="eeccf-168">Delete the content of the folder `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span> 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="eeccf-169">iOS の場合: </span><span class="sxs-lookup"><span data-stu-id="eeccf-169">For iOS:</span></span>
<span data-ttu-id="eeccf-p113">アドイン内の JavaScript から `window.location.reload(true)` を呼び出して強制的に再読み込みします。または、Office を再インストールしてください。</span><span class="sxs-lookup"><span data-stu-id="eeccf-p113">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="eeccf-172">関連項目</span><span class="sxs-lookup"><span data-stu-id="eeccf-172">See also</span></span>

- [<span data-ttu-id="eeccf-173">Office on the web でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="eeccf-173">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="eeccf-174">iPad または Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="eeccf-174">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="eeccf-175">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="eeccf-175">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="eeccf-176">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="eeccf-176">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
    
