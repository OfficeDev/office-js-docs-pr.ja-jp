---
title: アドインを使用したOfficeのトラブルシューティング
description: アドインの開発エラーをトラブルシューティングするOffice説明します。
ms.date: 06/11/2021
localization_priority: Normal
ms.openlocfilehash: 8f0ceaf13041fa27c4e9e279646e979f132913b3
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349279"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a><span data-ttu-id="d2270-103">アドインを使用したOfficeのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="d2270-103">Troubleshoot development errors with Office Add-ins</span></span>

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="d2270-104">アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題</span><span class="sxs-lookup"><span data-stu-id="d2270-104">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="d2270-105">アドインのマニフェストでの問題をデバッグするには、「[Office アドインのマニフェストを検証する](troubleshoot-manifest.md)」および「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d2270-105">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="d2270-106">リボン ボタンとメニュー項目が含まれているアドイン コマンドへの変更が反映されない</span><span class="sxs-lookup"><span data-stu-id="d2270-106">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="d2270-107">リボン ボタンのアイコンのファイル名やメニュー アイテムのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。</span><span class="sxs-lookup"><span data-stu-id="d2270-107">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="d2270-108">Windows の場合:</span><span class="sxs-lookup"><span data-stu-id="d2270-108">For Windows:</span></span>

<span data-ttu-id="d2270-109">フォルダーの内容を削除し `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 、フォルダーの内容が存在する場合は `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` 削除します。</span><span class="sxs-lookup"><span data-stu-id="d2270-109">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, and delete the contents of the folder `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\`, if it exists.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="d2270-110">Mac の場合: </span><span class="sxs-lookup"><span data-stu-id="d2270-110">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="d2270-111">iOS の場合: </span><span class="sxs-lookup"><span data-stu-id="d2270-111">For iOS:</span></span>

<span data-ttu-id="d2270-p101">アドイン内の JavaScript から `window.location.reload(true)` を呼び出して強制的に再読み込みします。または、Office を再インストールしてください。</span><span class="sxs-lookup"><span data-stu-id="d2270-p101">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="d2270-114">JavaScript、HTML、CSS などの静的ファイルへの変更は有効になりません</span><span class="sxs-lookup"><span data-stu-id="d2270-114">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="d2270-115">ブラウザーがこれらのファイルをキャッシュしている可能性があります。</span><span class="sxs-lookup"><span data-stu-id="d2270-115">The browser may be caching these files.</span></span> <span data-ttu-id="d2270-116">これを防ぐには、開発時にクライアント側のキャッシュをオフにします。</span><span class="sxs-lookup"><span data-stu-id="d2270-116">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="d2270-117">詳細は、使用しているサーバーの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="d2270-117">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="d2270-118">ほとんどの場合、HTTP 応答に特定のヘッダーを追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d2270-118">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="d2270-119">次のセットをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="d2270-119">We suggest the following set.</span></span>

- <span data-ttu-id="d2270-120">Cache Control: 「プライベート、キャッシュなし、ストアなし」</span><span class="sxs-lookup"><span data-stu-id="d2270-120">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="d2270-121">Pragma: 「no-cache」</span><span class="sxs-lookup"><span data-stu-id="d2270-121">Pragma: "no-cache"</span></span>
- <span data-ttu-id="d2270-122">有効期限: 「-1」</span><span class="sxs-lookup"><span data-stu-id="d2270-122">Expires: "-1"</span></span>

<span data-ttu-id="d2270-123">Node.JS Express サーバーでこれを行う例については、「[この app.js ファイル](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js)について」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d2270-123">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="d2270-124">ASP.NET プロジェクトの例については、「[この cshtml ファイル](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)について」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d2270-124">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="d2270-125">アドインがインターネット インフォメーション サービス (IIS) にホストされている場合は、次を web.config に追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="d2270-125">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="d2270-126">これらの手順が最初に動作しない場合は、ブラウザーのキャッシュをクリアする必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="d2270-126">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="d2270-127">これは、ブラウザーの UI を使用して行います。</span><span class="sxs-lookup"><span data-stu-id="d2270-127">Do this through the UI of the browser.</span></span> <span data-ttu-id="d2270-128">画面の端の UI でエッジ キャッシュをクリアしようとすると、正常にクリアされないことがあります。</span><span class="sxs-lookup"><span data-stu-id="d2270-128">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="d2270-129">その場合は、Windows コマンド プロンプトで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="d2270-129">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a><span data-ttu-id="d2270-130">プロパティ値に加えた変更は発生し、エラー メッセージはありません</span><span class="sxs-lookup"><span data-stu-id="d2270-130">Changes made to property values don't happen and there is no error message</span></span>

<span data-ttu-id="d2270-131">プロパティが読み取り専用である場合は、プロパティのリファレンス ドキュメントを確認してください。</span><span class="sxs-lookup"><span data-stu-id="d2270-131">Check the reference documentation for the property to see if it is read only.</span></span> <span data-ttu-id="d2270-132">また[、JS の TypeScript 定義Office、](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)読み取り専用のオブジェクト プロパティを指定します。</span><span class="sxs-lookup"><span data-stu-id="d2270-132">Also, the [TypeScript definitions](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="d2270-133">読み取り専用プロパティを設定しようとすると、書き込み操作はサイレント モードで失敗し、エラーはスローされます。</span><span class="sxs-lookup"><span data-stu-id="d2270-133">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="d2270-134">次の例では、読み取り専用プロパティを誤って設定 [Chart.id。](/javascript/api/excel/excel.chart#id)「一部 [のプロパティを直接設定できない」も参照してください](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)。</span><span class="sxs-lookup"><span data-stu-id="d2270-134">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id). See also [Some properties cannot be set directly](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a><span data-ttu-id="d2270-135">エラーの取得: "このアドインは使用できなくなりました"</span><span class="sxs-lookup"><span data-stu-id="d2270-135">Getting error: "This add-in is no longer available"</span></span>

<span data-ttu-id="d2270-136">このエラーの原因の一部を次に示します。</span><span class="sxs-lookup"><span data-stu-id="d2270-136">The following are some of the causes of this error.</span></span> <span data-ttu-id="d2270-137">その他の原因が見つかった場合は、ページの下部にあるフィードバック ツールを使って教えて下さい。</span><span class="sxs-lookup"><span data-stu-id="d2270-137">If you discover additional causes, please tell us with the feedback tool at the bottom of the page.</span></span>

- <span data-ttu-id="d2270-138">アプリケーションを使用しているVisual Studio、サイドローディングに問題がある可能性があります。</span><span class="sxs-lookup"><span data-stu-id="d2270-138">If you are using Visual Studio, there may be a problem with the sideloading.</span></span> <span data-ttu-id="d2270-139">ホストとホストのすべてのインスタンスOffice閉じるVisual Studio。</span><span class="sxs-lookup"><span data-stu-id="d2270-139">Close all instances of the Office host and Visual Studio.</span></span> <span data-ttu-id="d2270-140">再起動してVisual Studio F5 キーを再度押してみてください。</span><span class="sxs-lookup"><span data-stu-id="d2270-140">Restart Visual Studio and try pressing F5 again.</span></span>
- <span data-ttu-id="d2270-141">アドインのマニフェストは、展開場所 (集中展開、SharePoint、ネットワーク共有など) から削除されています。</span><span class="sxs-lookup"><span data-stu-id="d2270-141">The add-in's manifest has been removed from its deployment location, such as Centralized Deployment, a SharePoint catalog, or a network share.</span></span>
- <span data-ttu-id="d2270-142">マニフェスト内の [ID 要素](../reference/manifest/id.md) の値は、展開されたコピーで直接変更されています。</span><span class="sxs-lookup"><span data-stu-id="d2270-142">The value of the [ID](../reference/manifest/id.md) element in the manifest has been changed directly in the deployed copy.</span></span> <span data-ttu-id="d2270-143">何らかの理由でこの ID を変更する場合は、まず Office ホストからアドインを削除してから、元のマニフェストを変更したマニフェストに置き換える必要があります。</span><span class="sxs-lookup"><span data-stu-id="d2270-143">If for any reason, you want to change this ID, first remove the add-in from the Office host, then replace the original manifest with the changed manifest.</span></span> <span data-ttu-id="d2270-144">多くの場合、元のトレースOffice削除するには、キャッシュをクリアする必要があります。</span><span class="sxs-lookup"><span data-stu-id="d2270-144">You many need to clear the Office cache to remove all traces of the original.</span></span> <span data-ttu-id="d2270-145">この記事の [「リボン ボタンや](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) メニュー項目を含むアドイン コマンドの変更」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d2270-145">See the section [Changes to add-in commands including ribbon buttons and menu items do not take effect](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) earlier in this article.</span></span>
- <span data-ttu-id="d2270-146">アドインのマニフェストには、マニフェストの `resid` [[リソース](../reference/manifest/resources.md)] セクションのどこにも定義されていないか、使用する場所とセクションで定義されている場所のスペルが一致しません。 `resid` `<Resources>`</span><span class="sxs-lookup"><span data-stu-id="d2270-146">The add-in's manifest has a `resid` that is not defined anywhere in the [Resources](../reference/manifest/resources.md) section of the manifest, or there is a mismatch in the spelling of the `resid` between where it is used and where it is defined in the `<Resources>` section.</span></span>
- <span data-ttu-id="d2270-147">マニフェストの `resid` どこかに 32 文字を超える属性があります。</span><span class="sxs-lookup"><span data-stu-id="d2270-147">There is a `resid` attribute somewhere in the manifest with more than 32 characters.</span></span> <span data-ttu-id="d2270-148">属性 `resid` と、セクション内の対応するリソースの属性は `id` `<Resources>` 、32 文字を超えることはできません。</span><span class="sxs-lookup"><span data-stu-id="d2270-148">A `resid` attribute, and the `id` attribute of the corresponding resource in the `<Resources>` section, cannot be more than 32 characters.</span></span>
- <span data-ttu-id="d2270-149">アドインにはカスタム アドイン コマンドがありますが、それをサポートしないプラットフォームで実行しようとしている。</span><span class="sxs-lookup"><span data-stu-id="d2270-149">The add-in has a custom Add-in Command but you are trying to run it on a platform that doesn't support them.</span></span> <span data-ttu-id="d2270-150">詳細については、「アドイン コマンド [の要件セット」を参照してください](../reference/requirement-sets/add-in-commands-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="d2270-150">For more information, see [Add-in commands requirement sets](../reference/requirement-sets/add-in-commands-requirement-sets.md).</span></span>

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a><span data-ttu-id="d2270-151">アドインは Edge では機能しませんが、他のブラウザーで動作します</span><span class="sxs-lookup"><span data-stu-id="d2270-151">Add-in doesn't work on Edge but it works on other browsers</span></span>

<span data-ttu-id="d2270-152">「[トラブルシューティングと問題Microsoft Edgeする」を参照してください](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。</span><span class="sxs-lookup"><span data-stu-id="d2270-152">See [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span></span>

## <a name="excel-add-in-throws-errors-but-not-consistently"></a><span data-ttu-id="d2270-153">Excelはエラーをスローしますが、一貫して発生しません</span><span class="sxs-lookup"><span data-stu-id="d2270-153">Excel add-in throws errors, but not consistently</span></span>

<span data-ttu-id="d2270-154">考[えられる原因Excel、アドインのトラブルシューティング](../excel/excel-add-ins-troubleshooting.md)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d2270-154">See [Troubleshoot Excel add-ins](../excel/excel-add-ins-troubleshooting.md) for possible causes.</span></span>

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a><span data-ttu-id="d2270-155">プロジェクトのマニフェスト スキーマ検証Visual Studioエラー</span><span class="sxs-lookup"><span data-stu-id="d2270-155">Manifest schema validation errors in Visual Studio projects</span></span>

<span data-ttu-id="d2270-156">マニフェスト ファイルを変更する必要がある新しい機能を使用している場合は、マニフェスト ファイルで検証エラー Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="d2270-156">If you are using newer features that require changes to the manifest file, you may get validation errors in Visual Studio.</span></span> <span data-ttu-id="d2270-157">たとえば、共有 JavaScript ランタイムを実装する要素を追加すると、 `<Runtimes>` 次の検証エラーが表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="d2270-157">For example, when adding the `<Runtimes>` element to implement the shared JavaScript runtime, you may see the following validation error.</span></span>

<span data-ttu-id="d2270-158">**名前空間 ' ' の要素 'Host' に、名前空間 ' に無効な子要素 http://schemas.microsoft.com/office/taskpaneappversionoverrides 'Runtimes' が含 http://schemas.microsoft.com/office/taskpaneappversionoverrides まれている**</span><span class="sxs-lookup"><span data-stu-id="d2270-158">**The element 'Host' in namespace 'http://schemas.microsoft.com/office/taskpaneappversionoverrides' has invalid child element 'Runtimes' in namespace 'http://schemas.microsoft.com/office/taskpaneappversionoverrides'**</span></span>

<span data-ttu-id="d2270-159">この場合は、使用する XSD ファイルVisual Studio最新バージョンに更新できます。</span><span class="sxs-lookup"><span data-stu-id="d2270-159">If this occurs, you can update the XSD files that Visual Studio uses to the latest versions.</span></span> <span data-ttu-id="d2270-160">最新のスキーマ バージョンは [[MS-OWEMXML]: 付録 A: 完全な XML スキーマです](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)。</span><span class="sxs-lookup"><span data-stu-id="d2270-160">The latest schema versions are at [[MS-OWEMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).</span></span>

### <a name="locate-the-xsd-files"></a><span data-ttu-id="d2270-161">XSD ファイルを見つける</span><span class="sxs-lookup"><span data-stu-id="d2270-161">Locate the XSD files</span></span>

1. <span data-ttu-id="d2270-162">Visual Studio でプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="d2270-162">Open your project in Visual Studio.</span></span>
1. <span data-ttu-id="d2270-163">ソリューション **エクスプローラーで、** ファイルを開manifest.xmlします。</span><span class="sxs-lookup"><span data-stu-id="d2270-163">In **Solution Explorer**, open the manifest.xml file.</span></span> <span data-ttu-id="d2270-164">マニフェストは、通常、ソリューションの下の最初のプロジェクトに含まれます。</span><span class="sxs-lookup"><span data-stu-id="d2270-164">The manifest is typically in the first project under your solution.</span></span>
1. <span data-ttu-id="d2270-165">[プロパティ **の表示**  >  **] ウィンドウ**(F4) を選択します。</span><span class="sxs-lookup"><span data-stu-id="d2270-165">Choose **View** > **Properties Window** (F4).</span></span>
1. <span data-ttu-id="d2270-166">[プロパティ **] ウィンドウで**、省略記号 (...) を選択して XML スキーマ エディター **を開** きます。</span><span class="sxs-lookup"><span data-stu-id="d2270-166">In the **Properties Window**, choose the ellipsis (...) to open the **XML Schemas** editor.</span></span> <span data-ttu-id="d2270-167">ここでは、プロジェクトで使用しているすべてのスキーマ ファイルの正確なフォルダーの場所を確認できます。</span><span class="sxs-lookup"><span data-stu-id="d2270-167">Here you can find the exact folder location of all schema files your project uses.</span></span>

### <a name="update-the-xsd-files"></a><span data-ttu-id="d2270-168">XSD ファイルを更新する</span><span class="sxs-lookup"><span data-stu-id="d2270-168">Update the XSD files</span></span>

1. <span data-ttu-id="d2270-169">更新する XSD ファイルをテキスト エディターで開きます。</span><span class="sxs-lookup"><span data-stu-id="d2270-169">Open the XSD file you want to update in a text editor.</span></span> <span data-ttu-id="d2270-170">検証エラーのスキーマ名は XSD ファイル名に関連付けされます。</span><span class="sxs-lookup"><span data-stu-id="d2270-170">The schema name from the validation error will correlate to the XSD file name.</span></span> <span data-ttu-id="d2270-171">たとえば **、TaskPaneAppVersionOverridesV1_0.xsd を開きます**。</span><span class="sxs-lookup"><span data-stu-id="d2270-171">For example, open **TaskPaneAppVersionOverridesV1_0.xsd**.</span></span>
1. <span data-ttu-id="d2270-172">[[MS-OWEMXML] で](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)更新されたスキーマを検索します。付録 A: 完全な XML スキーマです。</span><span class="sxs-lookup"><span data-stu-id="d2270-172">Locate the updated schema at [[MS-OWEMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).</span></span> <span data-ttu-id="d2270-173">たとえば、TaskPaneAppVersionOverridesV1_0 [は taskpaneappversionoverrides スキーマ です](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40)。</span><span class="sxs-lookup"><span data-stu-id="d2270-173">For example, TaskPaneAppVersionOverridesV1_0 is at [taskpaneappversionoverrides Schema](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).</span></span>
1. <span data-ttu-id="d2270-174">テキストをテキスト エディターにコピーします。</span><span class="sxs-lookup"><span data-stu-id="d2270-174">Copy the text into your text editor.</span></span>
1. <span data-ttu-id="d2270-175">更新された XSD ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="d2270-175">Save the updated XSD file.</span></span>
1. <span data-ttu-id="d2270-176">新Visual Studio XSD ファイルの変更を取得するには、次のコマンドを再起動します。</span><span class="sxs-lookup"><span data-stu-id="d2270-176">Restart Visual Studio to pick up the new XSD file changes.</span></span>

<span data-ttu-id="d2270-177">古い追加のスキーマに対して、前のプロセスを繰り返します。</span><span class="sxs-lookup"><span data-stu-id="d2270-177">You can repeat the previous process for any additional schemas that are out-of-date.</span></span>

## <a name="see-also"></a><span data-ttu-id="d2270-178">関連項目</span><span class="sxs-lookup"><span data-stu-id="d2270-178">See also</span></span>

- [<span data-ttu-id="d2270-179">Office on the web でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="d2270-179">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)
- [<span data-ttu-id="d2270-180">iPad または Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="d2270-180">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="d2270-181">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="d2270-181">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="d2270-182">Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能</span><span class="sxs-lookup"><span data-stu-id="d2270-182">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="d2270-183">Office アドインのマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="d2270-183">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="d2270-184">ランタイム ログを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="d2270-184">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="d2270-185">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="d2270-185">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
