---
title: テスト用に Office on the web で Office アドインをサイドロードする
description: サイドローディングOfficeして、Office on the webアドインをテストします。
ms.date: 04/14/2021
localization_priority: Normal
ms.openlocfilehash: e830ccbb6a4e325d6d70c3612492009b5e3d1570
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077219"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a><span data-ttu-id="7bfd4-103">テスト用に Office on the web で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="7bfd4-103">Sideload Office Add-ins in Office on the web for testing</span></span>

<span data-ttu-id="7bfd4-104">アドインをサイドロードすると、アドインを最初にアドイン カタログに含めずにアドインをインストールできます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-104">When you sideload an add-in, you're able to install the add-in without first putting it in the add-in catalog.</span></span> <span data-ttu-id="7bfd4-105">これは、アドインの表示方法と機能を確認できるので、アドインをテストおよび開発する場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-105">This is useful when testing and developing your add-in because you can see how your add-in will appear and function.</span></span>

<span data-ttu-id="7bfd4-106">アドインをサイドロードすると、アドインのマニフェストはブラウザーのローカル ストレージに格納されます。そのため、ブラウザーのキャッシュをクリアするか、別のブラウザーに切り替える場合は、アドインを再度サイドロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-106">When you sideload an add-in, the add-in's manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>

<span data-ttu-id="7bfd4-107">サイドローディングは、ホスト アプリケーションによって異なります (たとえば、Excel)。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-107">Sideloading varies between host applications (for example, Excel).</span></span>

> [!NOTE]
> <span data-ttu-id="7bfd4-108">この記事で説明するサイドローディングは、Excel、OneNote、PowerPoint、および Word でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-108">Sideloading as described in this article is supported on Excel, OneNote, PowerPoint, and Word.</span></span> <span data-ttu-id="7bfd4-109">Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」をご参照ください。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-109">To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a><span data-ttu-id="7bfd4-110">Office on the web で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="7bfd4-110">Sideload an Office Add-in in Office on the web</span></span>

<span data-ttu-id="7bfd4-111">このプロセスは、word、Excel、OneNote、PowerPointでのみ **サポート** されます。 </span><span class="sxs-lookup"><span data-stu-id="7bfd4-111">This process is supported for **Excel**, **OneNote**, **PowerPoint**, and **Word** only.</span></span> <span data-ttu-id="7bfd4-112">他のホスト アプリケーションについては、次のセクションの手動サイドローディング手順を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-112">For other host applications, see the manual sideloading instructions in the following section.</span></span> <span data-ttu-id="7bfd4-113">このサンプル プロジェクトでは、Yeoman ジェネレーターを使用して作成されたプロジェクトを、アドインに使用[Office想定しています](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-113">This example project assumes that you are using a project created with [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

1. <span data-ttu-id="7bfd4-114">[ファイル[Office on the web] を開きます](https://office.live.com/)。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-114">Open [Office on the web](https://office.live.com/).</span></span> <span data-ttu-id="7bfd4-115">[作成 **] オプション** を使用して、Excel、OneNote、PowerPoint、**または** Word **で** ドキュメントを作成 **します**。 </span><span class="sxs-lookup"><span data-stu-id="7bfd4-115">Using the **Create** option, make a document in **Excel**, **OneNote**, **PowerPoint**, or **Word**.</span></span> <span data-ttu-id="7bfd4-116">この新しいドキュメントで、リボン **で [共有** ] を選択し、[リンクのコピー] **を** 選択して URL をコピーします。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-116">In this new document, select **Share** in the ribbon, select **Copy Link**, and copy the URL.</span></span>

2. <span data-ttu-id="7bfd4-117">yo office プロジェクト ファイルのルート ディレクトリで、ファイルのpackage.js **開** きます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-117">In the root directory of your yo office project files, open the **package.json** file.</span></span> <span data-ttu-id="7bfd4-118">このファイル **の構成** セクション内に、プロパティを作成 `"document"` します。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-118">Within the **config** section of this file, create a `"document"` property.</span></span> <span data-ttu-id="7bfd4-119">コピーした URL をプロパティの値として貼り付 `"document"` けます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-119">Paste the URL you copied as the value for the `"document"` property.</span></span> <span data-ttu-id="7bfd4-120">たとえば、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-120">For example, yours will look something like this:</span></span>

    ```json
      "config": {
        "document": "<YOUR URL>",
        ...
      }
    ```

    > [!TIP]
    > <span data-ttu-id="7bfd4-121">Yeoman ジェネレーターを使用しないアドインを作成する場合は、次の項目を既存の URL に追加して、ドキュメントの URL にクエリ パラメーターを追加できます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-121">If you are creating an add-in not using our Yeoman generator, you can add query parameters to your document's URL, by appending the following to the existing URL:</span></span>

    - <span data-ttu-id="7bfd4-122">など、開発サーバー ポート `&wdaddindevserverport=3000` 。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-122">The dev server port, such as `&wdaddindevserverport=3000`.</span></span>
    - <span data-ttu-id="7bfd4-123">マニフェスト ファイル名 ( `&wdaddinmanifestfile=manifest1.xml` など)。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-123">The manifest file name, such as `&wdaddinmanifestfile=manifest1.xml`.</span></span>
    - <span data-ttu-id="7bfd4-124">マニフェスト GUID など `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143` 。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-124">The manifest GUID, such as `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143`.</span></span>

    > <span data-ttu-id="7bfd4-125">Yeoman ジェネレーターを使用している場合は、Yeoman ツールによってこの情報が自動的に追加されるので、この情報を追加する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-125">If you are using the Yeoman generator, adding this information is not necessary as the Yeoman tooling appends this information automatically.</span></span>
    > <span data-ttu-id="7bfd4-126">ただし、どちらの場合も、localhost からのみマニフェストを読み込み可能です。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-126">Note that in both cases, however, you can only load manifests from localhost.</span></span>

3. <span data-ttu-id="7bfd4-127">プロジェクトのルート ディレクトリから始まるコマンド ラインで、次のコマンドを実行します `npm run start:web` 。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-127">In the command line starting at the root directory of your project, run the following command: `npm run start:web`.</span></span>

4. <span data-ttu-id="7bfd4-128">このメソッドを初めて使用して、Web 上にアドインをサイドロードすると、開発者モードを有効にしてくださいというダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-128">The first time you use this method to sideload an add-in on the web, you'll see a dialog asking you to enable developer mode.</span></span> <span data-ttu-id="7bfd4-129">[今すぐ開発者モードを **有効にする] のチェック ボックスをオンにして\*\*\*\*、[OK] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-129">Select the checkbox for **Enable Developer Mode now** and select **OK**.</span></span>

5. <span data-ttu-id="7bfd4-130">2 番目のダイアログ Office ボックスが表示されます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-130">You will see a second dialog box, asking if you wish to register an Office Add-in manifest from your computer.</span></span> <span data-ttu-id="7bfd4-131">[はい] を **選択する必要があります**。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-131">You should select **Yes**.</span></span>

6. <span data-ttu-id="7bfd4-132">アドインがインストールされています。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-132">Your add-in is installed.</span></span> <span data-ttu-id="7bfd4-133">アドイン コマンドの場合は、リボンまたはコンテキスト メニューに表示されます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-133">If it is an add-in command, it should appear on either the ribbon or the context menu.</span></span> <span data-ttu-id="7bfd4-134">作業ウィンドウ アドインの場合は、作業ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-134">If it is a task pane add-in, the task pane should appear.</span></span>

## <a name="sideload-an-office-add-in-in-office-on-the-web-manually"></a><span data-ttu-id="7bfd4-135">手動でOfficeアドインをサイドOffice on the webする</span><span class="sxs-lookup"><span data-stu-id="7bfd4-135">Sideload an Office Add-in in Office on the web manually</span></span>

<span data-ttu-id="7bfd4-136">このメソッドはコマンド ラインを使用しないので、ホスト アプリケーション内のコマンド (コマンド など) を使用Excel。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-136">This method doesn't use the command line and can be accomplished using commands only within the host application (such as Excel).</span></span>

1. <span data-ttu-id="7bfd4-137">[ファイル[Office on the web] を開きます](https://office.live.com/)。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-137">Open [Office on the web](https://office.live.com/).</span></span> <span data-ttu-id="7bfd4-138">[Word] 、または **[Excel]** で **ドキュメントを開** PowerPoint。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-138">Open a document in **Excel**, **Word**, or **PowerPoint**.</span></span> <span data-ttu-id="7bfd4-139">[アドイン **]** セクションのリボンの [挿入] タブ **で、[アドイン**] Office **を選択します**。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-139">On the **Insert** tab on the ribbon in the **Add-ins** section, choose **Office Add-ins**.</span></span>

1. <span data-ttu-id="7bfd4-140">[アドイン **Office]** ダイアログで **、[MY ADD-INS]** タブを選択し、[自分のアドインの管理] を選択し、[マイ アップロード] をクリック **します**。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-140">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>

    ![右上Officeの [アドインの管理] というドロップダウンが表示された [Office アドイン] ダイアログボックスと、その下に [アップロード マイ アドイン] というオプションが表示されます。](../images/office-add-ins-my-account.png)

1. <span data-ttu-id="7bfd4-142">アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-142">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>

    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

1. <span data-ttu-id="7bfd4-p111">アドインがインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-p111">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
> <span data-ttu-id="7bfd4-147">元の WebView Officeを使用Microsoft Edgeアドインをテストするには、追加の構成手順が必要です。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-147">To test your Office Add-in with Microsoft Edge with the original WebView (EdgeHTML), an additional configuration step is required.</span></span> <span data-ttu-id="7bfd4-148">コマンド プロンプトWindows、次の行を実行します `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes` 。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-148">In a Windows Command Prompt, run the following line: `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`.</span></span> <span data-ttu-id="7bfd4-149">この機能は、Officeベースのエッジ WebView2 をChromium場合は必要ありません。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-149">This is not required when Office is using the Chromium-based Edge WebView2.</span></span> <span data-ttu-id="7bfd4-150">詳細については、「アドインで使用Office[ブラウザー」を参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-150">For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

## <a name="sideload-an-office-add-in"></a><span data-ttu-id="7bfd4-151">アドインをサイドOfficeする</span><span class="sxs-lookup"><span data-stu-id="7bfd4-151">Sideload an Office Add-in</span></span>

1. <span data-ttu-id="7bfd4-152">アカウントにサインインMicrosoft 365します。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-152">Sign in to your Microsoft 365 account.</span></span>

2. <span data-ttu-id="7bfd4-153">ツール バーの起動ツールで [App 起動ツール] を開き、[Excel]、[Word]、または **[PowerPoint]** を選択し、新しいドキュメントを作成します。 </span><span class="sxs-lookup"><span data-stu-id="7bfd4-153">Open the App Launcher on the left end of the toolbar and select **Excel**, **Word**, or **PowerPoint**, and then create a new document.</span></span>

3. <span data-ttu-id="7bfd4-154">手順 3 から 6 は、前のセクション「**Office on the web で Office アドインをサイドロードする**」のものと同じです。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-154">Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office on the web**.</span></span>

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="7bfd4-155">Visual Studio の使用時にアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="7bfd4-155">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="7bfd4-156">アドインの開発にVisual Studio場合、サイドロードするプロセスは、Web への手動サイドローディングに似ています。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-156">If you're using Visual Studio to develop your add-in, the process to sideload is similar to manual sideloading to the web.</span></span> <span data-ttu-id="7bfd4-157">アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。唯一の違いは、マニフェストの **SourceURL** 要素の値を更新して、アドインが展開されている完全な URL を含める必要がある点です。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-157">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="7bfd4-158">アドインは Visual Studio から Office on the web にサイドロードできますが、Visual Studio からはデバッグできません。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-158">Although you can sideload add-ins from Visual Studio to Office on the web, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="7bfd4-159">デバッグするには、ブラウザー デバッグ ツールを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-159">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="7bfd4-160">詳細については、「[Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-160">For more information, see [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="7bfd4-161">Visual Studio で、[**表示**]  >  [**プロパティ ウィンドウ**] の順に選択して [**プロパティ**] ウィンドウを表示させます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-161">In Visual Studio, show the **Properties** window by choosing **View** > **Properties Window**.</span></span>
2. <span data-ttu-id="7bfd4-162">[**ソリューション エクスプローラー**] で Web プロジェクトを選択します。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-162">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="7bfd4-163">プロジェクトのプロパティが [**プロパティ**] ウィンドウに表示されます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-163">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="7bfd4-164">[プロパティ] ウィンドウで、[**SSL URL**] をコピーします。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-164">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="7bfd4-165">アドイン プロジェクトで、マニフェスト XML ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-165">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="7bfd4-166">編集しているのがソース XML であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-166">Be sure you are editing the source XML.</span></span> <span data-ttu-id="7bfd4-167">一部の種類のプロジェクトでは、Visual Studio は XML のビジュアル ビューを開きますが、これは次の手順で使用できません。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-167">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="7bfd4-168">**~remoteAppUrl/** のすべてのインスタンスを検索し、先ほどコピーした SSL URL と置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-168">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="7bfd4-169">プロジェクトの種類に応じていくつかの置換が表示され、新しい URL の表示は `https://localhost:44300/Home.html` に似たものになりま。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-169">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="7bfd4-170">XML ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-170">Save the XML file.</span></span>
7. <span data-ttu-id="7bfd4-171">Web プロジェクトを右クリックして、[**デバッグ**]  >  [**新しいインスタンスを開始**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-171">Right click the web project and choose **Debug** > **Start new instance**.</span></span> <span data-ttu-id="7bfd4-172">これにより、Office を起動することなく Web プロジェクトが実行されます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-172">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="7bfd4-173">前述の「[Office on the web で Office アドインをサイドロードする](#sideload-an-office-add-in-in-office-on-the-web)」で説明した手順を使用して、Office on the web からアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-173">From Office on the web, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office on the web](#sideload-an-office-add-in-in-office-on-the-web).</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="7bfd4-174">サイドロードされたアドインを削除する</span><span class="sxs-lookup"><span data-stu-id="7bfd4-174">Remove a sideloaded add-in</span></span>

<span data-ttu-id="7bfd4-175">ブラウザーのキャッシュをクリアすると、以前にサイドロードされたアドインを削除できます。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-175">You can remove a previously sideloaded add-in by clearing your browser's cache.</span></span> <span data-ttu-id="7bfd4-176">アドインのマニフェストに変更を加えた場合 (たとえば、アイコンのファイル名やアドイン コマンドのテキストを更新する) 場合は[、Office](clear-cache.md)キャッシュをクリアしてから、更新されたマニフェストを使用してアドインを再読み込みする必要があります。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-176">If you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you may need to [clear the Office cache](clear-cache.md) and then re-sideload the add-in using the updated manifest.</span></span> <span data-ttu-id="7bfd4-177">これを実行することにより、アドインは更新されたマニフェストの記載どおりに Office で表示されるようになります。</span><span class="sxs-lookup"><span data-stu-id="7bfd4-177">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="see-also"></a><span data-ttu-id="7bfd4-178">関連項目</span><span class="sxs-lookup"><span data-stu-id="7bfd4-178">See also</span></span>

- [<span data-ttu-id="7bfd4-179">iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="7bfd4-179">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="7bfd4-180">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="7bfd4-180">Sideload Outlook add-ins for testing</span></span>](../outlook/sideload-outlook-add-ins-for-testing.md)
- [<span data-ttu-id="7bfd4-181">Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="7bfd4-181">Clear the Office cache</span></span>](clear-cache.md)
