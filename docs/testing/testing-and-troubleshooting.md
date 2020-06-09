---
title: Office アドインでのユーザー エラーのトラブルシューティング
description: Office アドインでのユーザーエラーのトラブルシューティング方法について説明します。
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 859cf5019d40d05dbb3ad211d4d2934b309f3ccd
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612074"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="5609b-103">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="5609b-103">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="5609b-p101">時折、ユーザーは開発した Office アドインの問題に遭遇することがあります。たとえば、アドインが読み込みに失敗したり、アクセスできないなどです。この記事の情報は、ユーザーが Office アドインを使用する際に遭遇する一般的な問題を解決するために用いることができます。</span><span class="sxs-lookup"><span data-stu-id="5609b-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="5609b-107">また、[Fiddler](https://www.telerik.com/fiddler) を使用して、アドインの問題を特定してデバッグすることもできます。</span><span class="sxs-lookup"><span data-stu-id="5609b-107">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="5609b-108">一般的なエラーとトラブルシューティングの手順</span><span class="sxs-lookup"><span data-stu-id="5609b-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="5609b-109">次の表は、ユーザーが遭遇する可能性がある一般的なエラー メッセージとエラーを解決するためにユーザーが実行できる手順を示しています。</span><span class="sxs-lookup"><span data-stu-id="5609b-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="5609b-110">**エラー メッセージ**</span><span class="sxs-lookup"><span data-stu-id="5609b-110">**Error message**</span></span>|<span data-ttu-id="5609b-111">**解決策**</span><span class="sxs-lookup"><span data-stu-id="5609b-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="5609b-112">アプリのエラー: カタログに到達できませんでした</span><span class="sxs-lookup"><span data-stu-id="5609b-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="5609b-p102">ファイアウォールの設定を確認します。「カタログ」は、AppSource を指します。このメッセージは、ユーザーが AppSource にアクセスできないことを示しています。</span><span class="sxs-lookup"><span data-stu-id="5609b-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="5609b-p103">アプリのエラー: このアプリを起動できませんでした。このダイアログを閉じて問題を無視するか、[再起動] をクリックしてもう一度お試しください。</span><span class="sxs-lookup"><span data-stu-id="5609b-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="5609b-117">Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム ](https://support.microsoft.com/kb/2986156/)をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="5609b-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="5609b-118">エラー: オブジェクトがプロパティまたはメソッド 'defineProperty' をサポートしていません</span><span class="sxs-lookup"><span data-stu-id="5609b-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="5609b-119">Internet Explorerが互換モードで実行されていないことを確認します。</span><span class="sxs-lookup"><span data-stu-id="5609b-119">Confirm that Internet Explorer is not running in Compatibility Mode.</span></span> <span data-ttu-id="5609b-120">[ツール] > \*\*[互換表示設定] \*\* に移動します。</span><span class="sxs-lookup"><span data-stu-id="5609b-120">Go to Tools > **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="5609b-p105">ブラウザーのバージョンがサポートされていないため、アプリを読み込めませんでした。サポートされているブラウザーのバージョンの一覧についてはここをクリックしてください。</span><span class="sxs-lookup"><span data-stu-id="5609b-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="5609b-p106">ブラウザーが HTML5 のローカル ストレージをサポートしていることを確認するか、Internet Explorer の設定をリセットします。サポートされているブラウザーの詳細については、「[Office アドインを実行するための要件](../concepts/requirements-for-running-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5609b-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a><span data-ttu-id="5609b-125">アドインをインストールすると、ステータス バーに "アドイン読み込み中のエラー" と表示される</span><span class="sxs-lookup"><span data-stu-id="5609b-125">When installing an add-in, you see "Error loading add-in" in the status bar</span></span>

1. <span data-ttu-id="5609b-126">Office を終了します。</span><span class="sxs-lookup"><span data-stu-id="5609b-126">Close Office.</span></span>
2. <span data-ttu-id="5609b-127">マニフェストが有効であることを確認する</span><span class="sxs-lookup"><span data-stu-id="5609b-127">Verify that the manifest is valid</span></span>
3. <span data-ttu-id="5609b-128">アドインを再起動する</span><span class="sxs-lookup"><span data-stu-id="5609b-128">Restart the add-in</span></span>
4. <span data-ttu-id="5609b-129">もう一度アドインをインストールする。</span><span class="sxs-lookup"><span data-stu-id="5609b-129">Install the add-in again.</span></span>

<span data-ttu-id="5609b-130">また、フィードバックを寄せることができます。Windows または Mac 用 Excel を使用している場合は、Excel から直接 Office の機能拡張チームにフィードバックを送信できます。</span><span class="sxs-lookup"><span data-stu-id="5609b-130">You can also give us feedback: if using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="5609b-131">これを行うには、[**ファイル**] | [**フィードバック**] | [**問題点、改善の報告**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="5609b-131">To do this, select **File** | **Feedback** | **Send a Frown**.</span></span> <span data-ttu-id="5609b-132">問題点、改善の報告により、問題を理解するために必要なログが提供されます。</span><span class="sxs-lookup"><span data-stu-id="5609b-132">Sending a frown provides the necessary logs to understand the issue.</span></span>

## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="5609b-133">Outlook アドインが正常に機能しない</span><span class="sxs-lookup"><span data-stu-id="5609b-133">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="5609b-134">Windows で実行され、[Internet Explorer を使用している](../concepts/browsers-used-by-office-web-add-ins.md) Outlook アドインが正常に機能しない場合は、Internet Explorer でスクリプトのデバッグを有効にしてみてください。</span><span class="sxs-lookup"><span data-stu-id="5609b-134">If an Outlook add-in running on Windows and [using Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="5609b-135">[ツール] > [**インターネットオプション**の詳細] に移動  >  **Advanced**します。</span><span class="sxs-lookup"><span data-stu-id="5609b-135">Go to Tools > **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="5609b-136">**[参照]** で、**[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオフにします。</span><span class="sxs-lookup"><span data-stu-id="5609b-136">Under **Browsing**, uncheck **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="5609b-p108">これらの設定は、問題のトラブルシューティングを行う場合にのみチェックボックスをオフにすることをお勧めします。チェックボックスをオフにしたままにすると、参照時にメッセージが表示されます。問題が解決したら、**[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオンにしてください。</span><span class="sxs-lookup"><span data-stu-id="5609b-p108">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="5609b-140">Office 2013 でアドインがアクティブにならない</span><span class="sxs-lookup"><span data-stu-id="5609b-140">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="5609b-141">ユーザーが次の手順を実行したときに、アドインがアクティブにならない場合があります。</span><span class="sxs-lookup"><span data-stu-id="5609b-141">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="5609b-142">Office 2013 で自分の Microsoft アカウントでサインインする。</span><span class="sxs-lookup"><span data-stu-id="5609b-142">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="5609b-143">自分の Microsoft アカウントの 2 段階検証を有効にする。</span><span class="sxs-lookup"><span data-stu-id="5609b-143">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="5609b-144">アドインを挿入しようとする際に、メッセージに従って ID の確認を行う。</span><span class="sxs-lookup"><span data-stu-id="5609b-144">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="5609b-145">Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム](https://support.microsoft.com/kb/2986156/)をダウンロードしてください。</span><span class="sxs-lookup"><span data-stu-id="5609b-145">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="5609b-146">アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題</span><span class="sxs-lookup"><span data-stu-id="5609b-146">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="5609b-147">アドインのマニフェストでの問題をデバッグするには、「[Office アドインのマニフェストを検証する](troubleshoot-manifest.md)」および「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5609b-147">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="5609b-148">アドイン ダイアログ ボックスを表示できない</span><span class="sxs-lookup"><span data-stu-id="5609b-148">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="5609b-p109">Office アドインを使用するとき、ユーザーは、ダイアログ ボックスの表示を許可するよう求められます。ユーザーが **[許可]** を選択すると、次のエラー メッセージが発生します。</span><span class="sxs-lookup"><span data-stu-id="5609b-p109">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="5609b-p110">"ブラウザーのセキュリティ設定により、ダイアログ ボックスを作成できませんでした。別のブラウザーを試すか、アドレス バーに表示される [URL] とドメインが同じセキュリティ ゾーンに存在するようにブラウザーを構成してください。"</span><span class="sxs-lookup"><span data-stu-id="5609b-p110">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![ダイアログ ボックスのエラー メッセージのスクリーン ショット](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="5609b-154">**影響を受けるブラウザー**</span><span class="sxs-lookup"><span data-stu-id="5609b-154">**Affected browsers**</span></span>|<span data-ttu-id="5609b-155">**影響を受けるプラットフォーム**</span><span class="sxs-lookup"><span data-stu-id="5609b-155">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="5609b-156">Internet Explorer、Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="5609b-156">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="5609b-157">Office on the web</span><span class="sxs-lookup"><span data-stu-id="5609b-157">Office on the web</span></span>|

<span data-ttu-id="5609b-p111">この問題を解決するために、エンド ユーザーまたは管理者は、Internet Explore の信頼済みサイトのリストにアドインのドメインを追加することができます。Internet Explorer または Microsoft Edge ブラウザーのどちらを使用していても、同じ手順を使用します。</span><span class="sxs-lookup"><span data-stu-id="5609b-p111">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5609b-160">アドインを信頼しない場合は、信頼済みサイトのリストにアドインの URL を追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="5609b-160">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="5609b-161">URL を信頼済みサイトのリストに追加する方法:</span><span class="sxs-lookup"><span data-stu-id="5609b-161">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="5609b-162">[**コントロール パネル**]で、[**インターネット オプション**]  >  [**セキュリティ**] に移動します。</span><span class="sxs-lookup"><span data-stu-id="5609b-162">In **Control Panel**, go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="5609b-163">**[信頼済みサイト]** ゾーンを選択して、**[サイト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="5609b-163">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="5609b-164">エラー メッセージに表示される URL を入力して、**[追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="5609b-164">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="5609b-p112">アドインの使用をもう一度お試しください。問題が続く場合は、他のセキュリティ ゾーンの設定を変えて、アドインのドメインが Office アプリケーションのアドレス バーに表示される URL と同じゾーンに存在するようにします。</span><span class="sxs-lookup"><span data-stu-id="5609b-p112">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="5609b-p113">この問題は、ポップアップ モードでダイアログ API が使用されているときに発生します。この問題を防ぐには、[displayInFrame](/javascript/api/office/office.ui) フラグを使います。そのために、ページが iframe 内の表示をサポートしている必要があります。次の例は、フラグの使用方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="5609b-p113">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="5609b-171">リボン ボタンとメニュー項目が含まれているアドイン コマンドへの変更が反映されない</span><span class="sxs-lookup"><span data-stu-id="5609b-171">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="5609b-172">リボン ボタンのアイコンのファイル名やメニュー アイテムのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。</span><span class="sxs-lookup"><span data-stu-id="5609b-172">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="5609b-173">Windows の場合:</span><span class="sxs-lookup"><span data-stu-id="5609b-173">For Windows:</span></span>
<span data-ttu-id="5609b-174">フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` の内容を削除する</span><span class="sxs-lookup"><span data-stu-id="5609b-174">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="5609b-175">Mac の場合: </span><span class="sxs-lookup"><span data-stu-id="5609b-175">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="5609b-176">iOS の場合: </span><span class="sxs-lookup"><span data-stu-id="5609b-176">For iOS:</span></span>
<span data-ttu-id="5609b-p114">アドイン内の JavaScript から `window.location.reload(true)` を呼び出して強制的に再読み込みします。または、Office を再インストールしてください。</span><span class="sxs-lookup"><span data-stu-id="5609b-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="5609b-179">JavaScript、HTML、CSS などの静的ファイルへの変更は有効になりません</span><span class="sxs-lookup"><span data-stu-id="5609b-179">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="5609b-180">ブラウザーがこれらのファイルをキャッシュしている可能性があります。</span><span class="sxs-lookup"><span data-stu-id="5609b-180">The browser may be caching these files.</span></span> <span data-ttu-id="5609b-181">これを防ぐには、開発時にクライアント側のキャッシュをオフにします。</span><span class="sxs-lookup"><span data-stu-id="5609b-181">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="5609b-182">詳細は、使用しているサーバーの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="5609b-182">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="5609b-183">ほとんどの場合、HTTP 応答に特定のヘッダーを追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5609b-183">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="5609b-184">次の設定をお勧めします。</span><span class="sxs-lookup"><span data-stu-id="5609b-184">We suggest the following set:</span></span>

- <span data-ttu-id="5609b-185">Cache Control: 「プライベート、キャッシュなし、ストアなし」</span><span class="sxs-lookup"><span data-stu-id="5609b-185">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="5609b-186">Pragma: 「no-cache」</span><span class="sxs-lookup"><span data-stu-id="5609b-186">Pragma: "no-cache"</span></span>
- <span data-ttu-id="5609b-187">有効期限: 「-1」</span><span class="sxs-lookup"><span data-stu-id="5609b-187">Expires: "-1"</span></span>

<span data-ttu-id="5609b-188">Node.JS Express サーバーでこれを行う例については、「[この app.js ファイル](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js)について」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5609b-188">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="5609b-189">ASP.NET プロジェクトの例については、「[この cshtml ファイル](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)について」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5609b-189">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="5609b-190">アドインがインターネット インフォメーション サービス (IIS) にホストされている場合は、次を web.config に追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="5609b-190">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="5609b-191">これらの手順が最初に動作しない場合は、ブラウザーのキャッシュをクリアする必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="5609b-191">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="5609b-192">これは、ブラウザーの UI を使用して行います。</span><span class="sxs-lookup"><span data-stu-id="5609b-192">Do this through the UI of the browser.</span></span> <span data-ttu-id="5609b-193">画面の端の UI でエッジ キャッシュをクリアしようとすると、正常にクリアされないことがあります。</span><span class="sxs-lookup"><span data-stu-id="5609b-193">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="5609b-194">その場合は、Windows コマンド プロンプトで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="5609b-194">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="see-also"></a><span data-ttu-id="5609b-195">関連項目</span><span class="sxs-lookup"><span data-stu-id="5609b-195">See also</span></span>

- [<span data-ttu-id="5609b-196">Office on the web でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="5609b-196">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="5609b-197">iPad または Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="5609b-197">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="5609b-198">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="5609b-198">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="5609b-199">Office アドインのマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="5609b-199">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="5609b-200">ランタイム ログを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="5609b-200">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
