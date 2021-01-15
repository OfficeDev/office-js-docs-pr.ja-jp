---
title: アドインを使用したOfficeエラーのトラブルシューティング
description: アドインの開発エラーをトラブルシューティングするOffice説明します。
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: 48216230db4bf90ca53ef10d98786877bd3905c2
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771425"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a><span data-ttu-id="4e55b-103">アドインを使用したOfficeエラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="4e55b-103">Troubleshoot development errors with Office Add-ins</span></span>

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="4e55b-104">アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題</span><span class="sxs-lookup"><span data-stu-id="4e55b-104">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="4e55b-105">アドインのマニフェストでの問題をデバッグするには、「[Office アドインのマニフェストを検証する](troubleshoot-manifest.md)」および「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4e55b-105">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="4e55b-106">リボン ボタンとメニュー項目が含まれているアドイン コマンドへの変更が反映されない</span><span class="sxs-lookup"><span data-stu-id="4e55b-106">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="4e55b-107">リボン ボタンのアイコンのファイル名やメニュー アイテムのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。</span><span class="sxs-lookup"><span data-stu-id="4e55b-107">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="4e55b-108">Windows の場合:</span><span class="sxs-lookup"><span data-stu-id="4e55b-108">For Windows:</span></span>

<span data-ttu-id="4e55b-109">フォルダーの内容を削除し `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 、フォルダーの内容が存在する場合は削除 `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` します。</span><span class="sxs-lookup"><span data-stu-id="4e55b-109">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, and delete the contents of the folder `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\`, if it exists.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="4e55b-110">Mac の場合: </span><span class="sxs-lookup"><span data-stu-id="4e55b-110">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="4e55b-111">iOS の場合: </span><span class="sxs-lookup"><span data-stu-id="4e55b-111">For iOS:</span></span>
<span data-ttu-id="4e55b-p101">アドイン内の JavaScript から `window.location.reload(true)` を呼び出して強制的に再読み込みします。または、Office を再インストールしてください。</span><span class="sxs-lookup"><span data-stu-id="4e55b-p101">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="4e55b-114">JavaScript、HTML、CSS などの静的ファイルへの変更は有効になりません</span><span class="sxs-lookup"><span data-stu-id="4e55b-114">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="4e55b-115">ブラウザーがこれらのファイルをキャッシュしている可能性があります。</span><span class="sxs-lookup"><span data-stu-id="4e55b-115">The browser may be caching these files.</span></span> <span data-ttu-id="4e55b-116">これを防ぐには、開発時にクライアント側のキャッシュをオフにします。</span><span class="sxs-lookup"><span data-stu-id="4e55b-116">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="4e55b-117">詳細は、使用しているサーバーの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="4e55b-117">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="4e55b-118">ほとんどの場合、HTTP 応答に特定のヘッダーを追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4e55b-118">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="4e55b-119">次の設定をお勧めします。</span><span class="sxs-lookup"><span data-stu-id="4e55b-119">We suggest the following set:</span></span>

- <span data-ttu-id="4e55b-120">Cache Control: 「プライベート、キャッシュなし、ストアなし」</span><span class="sxs-lookup"><span data-stu-id="4e55b-120">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="4e55b-121">Pragma: 「no-cache」</span><span class="sxs-lookup"><span data-stu-id="4e55b-121">Pragma: "no-cache"</span></span>
- <span data-ttu-id="4e55b-122">有効期限: 「-1」</span><span class="sxs-lookup"><span data-stu-id="4e55b-122">Expires: "-1"</span></span>

<span data-ttu-id="4e55b-123">Node.JS Express サーバーでこれを行う例については、「[この app.js ファイル](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js)について」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4e55b-123">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="4e55b-124">ASP.NET プロジェクトの例については、「[この cshtml ファイル](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)について」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4e55b-124">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="4e55b-125">アドインがインターネット インフォメーション サービス (IIS) にホストされている場合は、次を web.config に追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="4e55b-125">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="4e55b-126">これらの手順が最初に動作しない場合は、ブラウザーのキャッシュをクリアする必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="4e55b-126">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="4e55b-127">これは、ブラウザーの UI を使用して行います。</span><span class="sxs-lookup"><span data-stu-id="4e55b-127">Do this through the UI of the browser.</span></span> <span data-ttu-id="4e55b-128">画面の端の UI でエッジ キャッシュをクリアしようとすると、正常にクリアされないことがあります。</span><span class="sxs-lookup"><span data-stu-id="4e55b-128">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="4e55b-129">その場合は、Windows コマンド プロンプトで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="4e55b-129">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a><span data-ttu-id="4e55b-130">プロパティ値に加えた変更は行われたので、エラー メッセージはありません</span><span class="sxs-lookup"><span data-stu-id="4e55b-130">Changes made to property values don't happen and there is no error message</span></span>

<span data-ttu-id="4e55b-131">プロパティが読み取り専用である場合は、そのプロパティのリファレンス ドキュメントを確認してください。</span><span class="sxs-lookup"><span data-stu-id="4e55b-131">Check the reference documentation for the property to see if it is read only.</span></span> <span data-ttu-id="4e55b-132">また、読み取り専用のOffice JS の [TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) 定義も指定します。</span><span class="sxs-lookup"><span data-stu-id="4e55b-132">Also, the [TypeScript definitions](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="4e55b-133">読み取り専用プロパティを設定しようとすると、書き込み操作はサイレント モードで失敗し、エラーはスローされます。</span><span class="sxs-lookup"><span data-stu-id="4e55b-133">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="4e55b-134">次の例では、誤って読み取り専用プロパティの設定を試 [Chart.id。](/javascript/api/excel/excel.chart#id)「一部 [のプロパティを直接設定できない」も参照してください](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)。</span><span class="sxs-lookup"><span data-stu-id="4e55b-134">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id). See also [Some properties cannot be set directly](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a><span data-ttu-id="4e55b-135">エラーが表示される: "このアドインは使用できなくなりました"</span><span class="sxs-lookup"><span data-stu-id="4e55b-135">Getting error: "This add-in is no longer available"</span></span>

<span data-ttu-id="4e55b-136">このエラーの原因の一部を次に示します。</span><span class="sxs-lookup"><span data-stu-id="4e55b-136">The following are some of the causes of this error.</span></span> <span data-ttu-id="4e55b-137">その他の原因が見つかった場合は、ページの下部にあるフィードバック ツールを使用してご連絡ください。</span><span class="sxs-lookup"><span data-stu-id="4e55b-137">If you discover additional causes, please tell us with the feedback tool at the bottom of the page.</span></span>

- <span data-ttu-id="4e55b-138">アプリを使用しているVisual Studio、サイドローディングに問題がある可能性があります。</span><span class="sxs-lookup"><span data-stu-id="4e55b-138">If you are using Visual Studio, there may be a problem with the sideloading.</span></span> <span data-ttu-id="4e55b-139">ホストとホストのすべてのインスタンスOffice閉Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="4e55b-139">Close all instances of the Office host and Visual Studio.</span></span> <span data-ttu-id="4e55b-140">再起動Visual Studio、もう一度 F5 キーを押してみてください。</span><span class="sxs-lookup"><span data-stu-id="4e55b-140">Restart Visual Studio and try pressing F5 again.</span></span>
- <span data-ttu-id="4e55b-141">アドインのマニフェストは、一元展開、SharePoint カタログ、ネットワーク共有など、展開場所から削除されました。</span><span class="sxs-lookup"><span data-stu-id="4e55b-141">The add-in's manifest has been removed from its deployment location, such as Centralized Deployment, a SharePoint catalog, or a network share.</span></span>
- <span data-ttu-id="4e55b-142">マニフェスト内の [ID 要素](../reference/manifest/id.md) の値は、展開されたコピーで直接変更されています。</span><span class="sxs-lookup"><span data-stu-id="4e55b-142">The value of the [ID](../reference/manifest/id.md) element in the manifest has been changed directly in the deployed copy.</span></span> <span data-ttu-id="4e55b-143">何らかの理由でこの ID を変更する場合は、まず Office ホストからアドインを削除してから、元のマニフェストを変更されたマニフェストに置き換える必要があります。</span><span class="sxs-lookup"><span data-stu-id="4e55b-143">If for any reason, you want to change this ID, first remove the add-in from the Office host, then replace the original manifest with the changed manifest.</span></span> <span data-ttu-id="4e55b-144">多くの場合、元のOfficeトレースを削除するには、キャッシュをクリアする必要があります。</span><span class="sxs-lookup"><span data-stu-id="4e55b-144">You many need to clear the Office cache to remove all traces of the original.</span></span> <span data-ttu-id="4e55b-145">「リボン ボタンや [メニュー項目を](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) 含むアドイン コマンドに対する変更は、この記事の前の方では有効ではありません。」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="4e55b-145">See the section [Changes to add-in commands including ribbon buttons and menu items do not take effect](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) earlier in this article.</span></span>
- <span data-ttu-id="4e55b-146">アドインのマニフェストには、マニフェストの Resources セクションのどこにも定義されていないものがあります。または、アドインが使用される場所とセクションで定義されている場所のスペルが一致しません。 `resid` [](../reference/manifest/resources.md) `resid` `<Resources>`</span><span class="sxs-lookup"><span data-stu-id="4e55b-146">The add-in's manifest has a `resid` that is not defined anywhere in the [Resources](../reference/manifest/resources.md) section of the manifest, or there is a mismatch in the spelling of the `resid` between where it is used and where it is defined in the `<Resources>` section.</span></span>
- <span data-ttu-id="4e55b-147">マニフェストのどこかに 32 文字を超える `resid` 属性があります。</span><span class="sxs-lookup"><span data-stu-id="4e55b-147">There is a `resid` attribute somewhere in the manifest with more than 32 characters.</span></span> <span data-ttu-id="4e55b-148">属性 `resid` と、セクション内の対応するリソースの属性は `id` `<Resources>` 、32 文字を超えることはできません。</span><span class="sxs-lookup"><span data-stu-id="4e55b-148">A `resid` attribute, and the `id` attribute of the corresponding resource in the `<Resources>` section, cannot be more than 32 characters.</span></span>
- <span data-ttu-id="4e55b-149">アドインにはカスタム アドイン コマンドがありますが、それをサポートしないプラットフォーム上で実行しようとしている。</span><span class="sxs-lookup"><span data-stu-id="4e55b-149">The add-in has a custom Add-in Command but you are trying to run it on a platform that doesn't support them.</span></span> <span data-ttu-id="4e55b-150">詳細については、アドイン コマンド [の要件セットを参照してください](../reference/requirement-sets/add-in-commands-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="4e55b-150">For more information, see [Add-in commands requirement sets](../reference/requirement-sets/add-in-commands-requirement-sets.md).</span></span>

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a><span data-ttu-id="4e55b-151">アドインは Edge では動作しませんが、他のブラウザーで動作します</span><span class="sxs-lookup"><span data-stu-id="4e55b-151">Add-in doesn't work on Edge but it works on other browsers</span></span>

<span data-ttu-id="4e55b-152">「Microsoft [Edge の問題のトラブルシューティング」を参照してください](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。</span><span class="sxs-lookup"><span data-stu-id="4e55b-152">See [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span></span>

## <a name="excel-add-in-throws-errors-but-not-consistently"></a><span data-ttu-id="4e55b-153">Excel アドインはエラーをスローしますが、一貫してスローしません</span><span class="sxs-lookup"><span data-stu-id="4e55b-153">Excel add-in throws errors, but not consistently</span></span>

<span data-ttu-id="4e55b-154">考 [えられる原因については、「Excel アドインの](../excel/excel-add-ins-troubleshooting.md) トラブルシューティング」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4e55b-154">See [Troubleshoot Excel add-ins](../excel/excel-add-ins-troubleshooting.md) for possible causes.</span></span>

## <a name="see-also"></a><span data-ttu-id="4e55b-155">関連項目</span><span class="sxs-lookup"><span data-stu-id="4e55b-155">See also</span></span>

- [<span data-ttu-id="4e55b-156">Office on the web でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="4e55b-156">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)
- [<span data-ttu-id="4e55b-157">iPad または Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="4e55b-157">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="4e55b-158">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="4e55b-158">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="4e55b-159">Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能</span><span class="sxs-lookup"><span data-stu-id="4e55b-159">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="4e55b-160">Office アドインのマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="4e55b-160">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="4e55b-161">ランタイム ログを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="4e55b-161">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="4e55b-162">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="4e55b-162">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
