---
title: Office アドインでのユーザー エラーのトラブルシューティング
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: fb1b15236214056e6e15b4581a3813d42e31dc54
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270776"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="43ea0-102">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="43ea0-102">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="43ea0-p101">時折、ユーザーは開発した Office アドインの問題に遭遇することがあります。たとえば、アドインが読み込みに失敗したり、アクセスできないなどです。この記事の情報は、ユーザーが Office アドインを使用する際に遭遇する一般的な問題を解決するために用いることができます。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="43ea0-106">また、[Fiddler](https://www.telerik.com/fiddler) を使用して、アドインの問題を特定してデバッグすることもできます。</span><span class="sxs-lookup"><span data-stu-id="43ea0-106">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

<span data-ttu-id="43ea0-107">ユーザーの問題を解決した後、[AppSource でカスタマー レビューに直接返信することができます](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)。</span><span class="sxs-lookup"><span data-stu-id="43ea0-107">After you resolve the user's issue, you can [respond directly to customer reviews in AppSource](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="43ea0-108">一般的なエラーとトラブルシューティングの手順</span><span class="sxs-lookup"><span data-stu-id="43ea0-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="43ea0-109">次の表は、ユーザーが遭遇する可能性がある一般的なエラー メッセージとエラーを解決するためにユーザーが実行できる手順を示しています。</span><span class="sxs-lookup"><span data-stu-id="43ea0-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="43ea0-110">**エラー メッセージ**</span><span class="sxs-lookup"><span data-stu-id="43ea0-110">**Error message**</span></span>|<span data-ttu-id="43ea0-111">**解決策**</span><span class="sxs-lookup"><span data-stu-id="43ea0-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="43ea0-112">アプリのエラー: カタログに到達できませんでした</span><span class="sxs-lookup"><span data-stu-id="43ea0-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="43ea0-p102">ファイアウォールの設定を確認します。「カタログ」は、AppSource を指します。このメッセージは、ユーザーが AppSource にアクセスできないことを示しています。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="43ea0-p103">アプリのエラー: このアプリを起動できませんでした。このダイアログを閉じて問題を無視するか、[再起動] をクリックしてもう一度お試しください。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="43ea0-117">Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム](https://support.microsoft.com/kb/2986156/)をダウンロードしてください。</span><span class="sxs-lookup"><span data-stu-id="43ea0-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="43ea0-118">エラー: オブジェクトがプロパティまたはメソッド 'defineProperty' をサポートしていません</span><span class="sxs-lookup"><span data-stu-id="43ea0-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="43ea0-p104">Internet Explorerが互換モードで実行されていないことを確認します。 [ツール] >  **[互換表示設定]** に移動します。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p104">Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools >  **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="43ea0-p105">ブラウザーのバージョンがサポートされていないため、アプリを読み込めませんでした。サポートされているブラウザーのバージョンの一覧についてはここをクリックしてください。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="43ea0-p106">ブラウザーが HTML5 のローカル ストレージをサポートしていることを確認するか、Internet Explorer の設定をリセットします。サポートされているブラウザーの詳細については、「[Office アドインを実行するための要件](../concepts/requirements-for-running-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|


## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="43ea0-125">Outlook アドインが正常に機能しない</span><span class="sxs-lookup"><span data-stu-id="43ea0-125">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="43ea0-126">Windows で実行している Outlook アドインが正常に機能しない場合は、Internet Explorer でスクリプトのデバッグを有効にしてみてください。</span><span class="sxs-lookup"><span data-stu-id="43ea0-126">If an Outlook add-in running on Windows is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="43ea0-127">[ツール] >  **[インターネット オプション]** > **[詳細]** に移動します。</span><span class="sxs-lookup"><span data-stu-id="43ea0-127">Go to Tools >  **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="43ea0-128">**[参照]** で、 **[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオフにします。</span><span class="sxs-lookup"><span data-stu-id="43ea0-128">Under  **Browsing**, uncheck  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="43ea0-p107">これらの設定は、問題のトラブルシューティングを行う場合にのみチェックボックスをオフにすることをお勧めします。チェックボックスをオフにしたままにすると、参照時にメッセージが表示されます。問題が解決したら、 **[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオンにしてください。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p107">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="43ea0-132">Office 2013 でアドインがアクティブにならない</span><span class="sxs-lookup"><span data-stu-id="43ea0-132">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="43ea0-133">ユーザーが次の手順を実行したときに、アドインがアクティブにならない場合があります。</span><span class="sxs-lookup"><span data-stu-id="43ea0-133">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="43ea0-134">Office 2013 で自分の Microsoft アカウントでサインインする。</span><span class="sxs-lookup"><span data-stu-id="43ea0-134">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="43ea0-135">自分の Microsoft アカウントの 2 段階検証を有効にする。</span><span class="sxs-lookup"><span data-stu-id="43ea0-135">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="43ea0-136">アドインを挿入しようとする際に、メッセージに従って ID の確認を行う。</span><span class="sxs-lookup"><span data-stu-id="43ea0-136">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="43ea0-137">Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム](https://support.microsoft.com/kb/2986156/)をダウンロードしてください。</span><span class="sxs-lookup"><span data-stu-id="43ea0-137">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="43ea0-138">アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題</span><span class="sxs-lookup"><span data-stu-id="43ea0-138">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="43ea0-139">「[マニフェストの問題を検証し、トラブルシューティングする](troubleshoot-manifest.md)」を参照して、アドインのマニフェストの問題をデバッグしてください。</span><span class="sxs-lookup"><span data-stu-id="43ea0-139">See [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="43ea0-140">アドイン ダイアログ ボックスを表示できない</span><span class="sxs-lookup"><span data-stu-id="43ea0-140">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="43ea0-p108">Office アドインを使用するとき、ユーザーは、ダイアログ ボックスの表示を許可するよう求められます。ユーザーが **[許可]** を選択すると、次のエラー メッセージが発生します。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p108">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="43ea0-p109">"ブラウザーのセキュリティ設定により、ダイアログ ボックスを作成できませんでした。別のブラウザーを試すか、アドレス バーに表示される [URL] とドメインが同じセキュリティ ゾーンに存在するようにブラウザーを構成してください。"</span><span class="sxs-lookup"><span data-stu-id="43ea0-p109">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![ダイアログ ボックスのエラー メッセージのスクリーン ショット](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="43ea0-146">**影響を受けるブラウザー**</span><span class="sxs-lookup"><span data-stu-id="43ea0-146">**Affected browsers**</span></span>|<span data-ttu-id="43ea0-147">**影響を受けるプラットフォーム**</span><span class="sxs-lookup"><span data-stu-id="43ea0-147">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="43ea0-148">Internet Explorer、Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="43ea0-148">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="43ea0-149">Office Online</span><span class="sxs-lookup"><span data-stu-id="43ea0-149">Office Online</span></span>|

<span data-ttu-id="43ea0-p110">この問題を解決するために、エンド ユーザーまたは管理者は、Internet Explore の信頼済みサイトのリストにアドインのドメインを追加することができます。Internet Explorer または Microsoft Edge ブラウザーのどちらを使用していても、同じ手順を使用します。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p110">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="43ea0-152">アドインを信頼しない場合は、信頼済みサイトのリストにアドインの URL を追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="43ea0-152">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="43ea0-153">URL を信頼済みサイトのリストに追加する方法:</span><span class="sxs-lookup"><span data-stu-id="43ea0-153">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="43ea0-154">Internet Explorer で [ツール] ボタンを選択し、**[インター ネット オプション]** > **[セキュリティ]** へ移動します。</span><span class="sxs-lookup"><span data-stu-id="43ea0-154">In Internet Explorer, choose the Tools button, and go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="43ea0-155">**[信頼済みサイト]** ゾーンを選択して、**[サイト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="43ea0-155">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="43ea0-156">エラー メッセージに表示される URL を入力して、**[追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="43ea0-156">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="43ea0-p111">アドインの使用をもう一度お試しください。問題が続く場合は、他のセキュリティ ゾーンの設定を変えて、アドインのドメインが Office アプリケーションのアドレス バーに表示される URL と同じゾーンに存在するようにします。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p111">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="43ea0-p112">この問題は、ポップアップ モードでダイアログ API が使用されているときに発生します。この問題を防ぐには、[displayInFrame](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) フラグを使います。そのために、ページが iframe 内の表示をサポートしている必要があります。次の例は、フラグの使用方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p112">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js

Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="43ea0-163">リボン ボタンとメニュー項目が含まれているアドイン コマンドへの変更が反映されない</span><span class="sxs-lookup"><span data-stu-id="43ea0-163">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>
<span data-ttu-id="43ea0-p113">アドイン コマンドにリボン ボタンのアイコンやメニュー項目のテキストなどの変更を加えても、その変更が反映されないことがあります。その場合は、以前のバージョンの Office のキャッシュをクリアしてください。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p113">Sometimes changes to add-in commands such as the icon for a ribbon button or the text of a menu item do not seem to take effect. Clear the Office cache of the old versions.</span></span>

#### <a name="for-windows"></a><span data-ttu-id="43ea0-166">Windows の場合:</span><span class="sxs-lookup"><span data-stu-id="43ea0-166">For Windows:</span></span>
<span data-ttu-id="43ea0-167">フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` の内容を削除します。</span><span class="sxs-lookup"><span data-stu-id="43ea0-167">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="43ea0-168">Mac の場合: </span><span class="sxs-lookup"><span data-stu-id="43ea0-168">For Mac:</span></span>
<span data-ttu-id="43ea0-169">フォルダー `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` の内容を削除します。</span><span class="sxs-lookup"><span data-stu-id="43ea0-169">Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

#### <a name="for-ios"></a><span data-ttu-id="43ea0-170">iOS の場合: </span><span class="sxs-lookup"><span data-stu-id="43ea0-170">For iOS:</span></span>
<span data-ttu-id="43ea0-p114">アドイン内の JavaScript から `window.location.reload(true)` を呼び出して強制的に再読み込みします。または、Office を再インストールしてください。</span><span class="sxs-lookup"><span data-stu-id="43ea0-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="43ea0-173">関連項目</span><span class="sxs-lookup"><span data-stu-id="43ea0-173">See also</span></span>

- [<span data-ttu-id="43ea0-174">Office Online でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="43ea0-174">Debug add-ins in Office Online</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="43ea0-175">iPad または Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="43ea0-175">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="43ea0-176">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="43ea0-176">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="43ea0-177">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="43ea0-177">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
    
