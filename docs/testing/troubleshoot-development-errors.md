---
title: Office アドインでの開発エラーのトラブルシューティング
description: Office アドインの開発エラーをトラブルシューティングする方法について説明します。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 5801146165446352ec806f6f832e9976f96467ac
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409410"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a><span data-ttu-id="f5f3d-103">Office アドインでの開発エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="f5f3d-103">Troubleshoot development errors with Office Add-ins</span></span>

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="f5f3d-104">アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題</span><span class="sxs-lookup"><span data-stu-id="f5f3d-104">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="f5f3d-105">アドインのマニフェストでの問題をデバッグするには、「[Office アドインのマニフェストを検証する](troubleshoot-manifest.md)」および「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-105">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="f5f3d-106">リボン ボタンとメニュー項目が含まれているアドイン コマンドへの変更が反映されない</span><span class="sxs-lookup"><span data-stu-id="f5f3d-106">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="f5f3d-107">リボン ボタンのアイコンのファイル名やメニュー アイテムのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-107">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="f5f3d-108">Windows の場合:</span><span class="sxs-lookup"><span data-stu-id="f5f3d-108">For Windows:</span></span>

<span data-ttu-id="f5f3d-109">フォルダーの内容を削除 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` し、フォルダーの内容を削除し `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` ます (存在する場合)。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-109">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, and delete the contents of the folder `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\`, if it exists.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="f5f3d-110">Mac の場合: </span><span class="sxs-lookup"><span data-stu-id="f5f3d-110">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="f5f3d-111">iOS の場合: </span><span class="sxs-lookup"><span data-stu-id="f5f3d-111">For iOS:</span></span>
<span data-ttu-id="f5f3d-p101">アドイン内の JavaScript から `window.location.reload(true)` を呼び出して強制的に再読み込みします。または、Office を再インストールしてください。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-p101">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="f5f3d-114">JavaScript、HTML、CSS などの静的ファイルへの変更は有効になりません</span><span class="sxs-lookup"><span data-stu-id="f5f3d-114">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="f5f3d-115">ブラウザーがこれらのファイルをキャッシュしている可能性があります。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-115">The browser may be caching these files.</span></span> <span data-ttu-id="f5f3d-116">これを防ぐには、開発時にクライアント側のキャッシュをオフにします。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-116">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="f5f3d-117">詳細は、使用しているサーバーの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-117">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="f5f3d-118">ほとんどの場合、HTTP 応答に特定のヘッダーを追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-118">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="f5f3d-119">次の設定をお勧めします。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-119">We suggest the following set:</span></span>

- <span data-ttu-id="f5f3d-120">Cache Control: 「プライベート、キャッシュなし、ストアなし」</span><span class="sxs-lookup"><span data-stu-id="f5f3d-120">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="f5f3d-121">Pragma: 「no-cache」</span><span class="sxs-lookup"><span data-stu-id="f5f3d-121">Pragma: "no-cache"</span></span>
- <span data-ttu-id="f5f3d-122">有効期限: 「-1」</span><span class="sxs-lookup"><span data-stu-id="f5f3d-122">Expires: "-1"</span></span>

<span data-ttu-id="f5f3d-123">Node.JS Express サーバーでこれを行う例については、「[この app.js ファイル](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js)について」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-123">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="f5f3d-124">ASP.NET プロジェクトの例については、「[この cshtml ファイル](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)について」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-124">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="f5f3d-125">アドインがインターネット インフォメーション サービス (IIS) にホストされている場合は、次を web.config に追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-125">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="f5f3d-126">これらの手順が最初に動作しない場合は、ブラウザーのキャッシュをクリアする必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-126">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="f5f3d-127">これは、ブラウザーの UI を使用して行います。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-127">Do this through the UI of the browser.</span></span> <span data-ttu-id="f5f3d-128">画面の端の UI でエッジ キャッシュをクリアしようとすると、正常にクリアされないことがあります。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-128">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="f5f3d-129">その場合は、Windows コマンド プロンプトで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-129">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a><span data-ttu-id="f5f3d-130">プロパティ値に対する変更は行われず、エラーメッセージもありません。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-130">Changes made to property values don't happen and there is no error message</span></span>

<span data-ttu-id="f5f3d-131">プロパティの参照ドキュメントが読み取り専用かどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-131">Check the reference documentation for the property to see if it is read only.</span></span> <span data-ttu-id="f5f3d-132">また、Office JS の [TypeScript 定義](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) は、読み取り専用のオブジェクトプロパティを指定します。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-132">Also, the [TypeScript definitions](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="f5f3d-133">読み取り専用プロパティを設定しようとすると、エラーがスローされずに書き込み操作が失敗します。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-133">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="f5f3d-134">次の例では、誤って読み取り専用プロパティ [Chart.id](/javascript/api/excel/excel.chart#id)を設定しようとしています。一部の [プロパティを直接設定することはできません](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-134">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id). See also [Some properties cannot be set directly](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a><span data-ttu-id="f5f3d-135">アドインはエッジでは動作しませんが、他のブラウザーで動作します。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-135">Add-in doesn't work on Edge but it works on other browsers</span></span>

<span data-ttu-id="f5f3d-136">[Microsoft Edge の問題のトラブルシューティング](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-136">See [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span></span>

## <a name="excel-add-in-throws-errors-but-not-consistently"></a><span data-ttu-id="f5f3d-137">Excel アドインはエラーをスローしますが、一貫していません</span><span class="sxs-lookup"><span data-stu-id="f5f3d-137">Excel add-in throws errors, but not consistently</span></span>

<span data-ttu-id="f5f3d-138">考えられる原因については、「 [Excel アドインのトラブルシューティング](../excel/excel-add-ins-troubleshooting.md) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f5f3d-138">See [Troubleshoot Excel add-ins](../excel/excel-add-ins-troubleshooting.md) for possible causes.</span></span>

## <a name="see-also"></a><span data-ttu-id="f5f3d-139">関連項目</span><span class="sxs-lookup"><span data-stu-id="f5f3d-139">See also</span></span>

- [<span data-ttu-id="f5f3d-140">Office on the web でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="f5f3d-140">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)
- [<span data-ttu-id="f5f3d-141">iPad または Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="f5f3d-141">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="f5f3d-142">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="f5f3d-142">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="f5f3d-143">Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能</span><span class="sxs-lookup"><span data-stu-id="f5f3d-143">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="f5f3d-144">Office アドインのマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="f5f3d-144">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="f5f3d-145">ランタイム ログを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="f5f3d-145">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="f5f3d-146">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="f5f3d-146">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
