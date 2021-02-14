---
title: テスト用に Office on the web で Office アドインをサイドロードする
description: サイドロードOffice Web 上のOfficeアドインをテストします。
ms.date: 02/11/2021
localization_priority: Normal
ms.openlocfilehash: f81fbc163135be5a616e7b44e604cb842da9870b
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238065"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a><span data-ttu-id="16134-103">テスト用に Office on the web で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="16134-103">Sideload Office Add-ins in Office on the web for testing</span></span>

<span data-ttu-id="16134-104">アドインをサイドロードすると、最初にアドイン カタログにアドインを置かずにアドインをインストールできます。</span><span class="sxs-lookup"><span data-stu-id="16134-104">When you sideload an add-in, you're able to install the add-in without first putting it in the add-in catalog.</span></span> <span data-ttu-id="16134-105">これは、アドインの表示方法と機能を確認できるアドインをテストおよび開発するときに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="16134-105">This is useful when testing and developing your add-in because you can see how your add-in will appear and function.</span></span>

<span data-ttu-id="16134-106">アドインをサイドロードすると、アドインのマニフェストはブラウザーのローカル ストレージに格納されます。そのため、ブラウザーのキャッシュをクリアするか、別のブラウザーに切り替える場合は、アドインを再びサイドロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="16134-106">When you sideload an add-in, the add-in's manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>

<span data-ttu-id="16134-107">サイドロードは、ホスト アプリケーション (Excel など) によって異なります。</span><span class="sxs-lookup"><span data-stu-id="16134-107">Sideloading varies between host applications (for example, Excel).</span></span>

> [!NOTE]
> <span data-ttu-id="16134-p102">この記事で説明するようにサイドロードは、Excel、OneNote、PowerPoint、および Word でサポートされています。Outlook アドインをサイドロードするには、「テスト用に Outlook アドイン [をサイドロードする」を参照してください](../outlook/sideload-outlook-add-ins-for-testing.md)。</span><span class="sxs-lookup"><span data-stu-id="16134-p102">Sideloading as described in this article is supported on Excel, OneNote, PowerPoint, and Word. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a><span data-ttu-id="16134-110">Office on the web で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="16134-110">Sideload an Office Add-in in Office on the web</span></span>

<span data-ttu-id="16134-111">このプロセスは、Excel、OneNote、PowerPoint、**および** **Word でのみサポート** されています。  </span><span class="sxs-lookup"><span data-stu-id="16134-111">This process is supported for **Excel**, **OneNote**, **PowerPoint**, and **Word** only.</span></span> <span data-ttu-id="16134-112">その他のホスト アプリケーションについては、次のセクションの手動サイドロード手順を参照してください。</span><span class="sxs-lookup"><span data-stu-id="16134-112">For other host applications, see the manual sideloading instructions in the following section.</span></span> <span data-ttu-id="16134-113">このサンプル プロジェクトでは、アドイン用の Yeoman ジェネレーターで作成された [プロジェクトOffice想定しています](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="16134-113">This example project assumes that you are using a project created with [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

1. <span data-ttu-id="16134-114">Web [Officeを開きます](https://office.live.com/)。</span><span class="sxs-lookup"><span data-stu-id="16134-114">Open [Office on the web](https://office.live.com/).</span></span> <span data-ttu-id="16134-115">[作成 **] オプション** を使用して、Excel、OneNote、PowerPoint、または Word **でドキュメント** を作成 **します**。 </span><span class="sxs-lookup"><span data-stu-id="16134-115">Using the **Create** option, make a document in **Excel**, **OneNote**, **PowerPoint**, or **Word**.</span></span> <span data-ttu-id="16134-116">この新しいドキュメントで、リボンで **[共有**]を選択し、[リンクのコピー] を選択して URL をコピーします。</span><span class="sxs-lookup"><span data-stu-id="16134-116">In this new document, select **Share** in the ribbon, select **Copy Link**, and copy the URL.</span></span>

2. <span data-ttu-id="16134-117">yo office プロジェクト ファイルのルート ディレクトリで、ファイルのpackage.js **開** きます。</span><span class="sxs-lookup"><span data-stu-id="16134-117">In the root directory of your yo office project files, open the **package.json** file.</span></span> <span data-ttu-id="16134-118">このファイル **のスクリプト** セクション内にプロパティを作成 `"document"` します。</span><span class="sxs-lookup"><span data-stu-id="16134-118">Within the **scripts** section of this file, create a `"document"` property.</span></span> <span data-ttu-id="16134-119">コピーした URL をプロパティの値として貼り付 `"document"` けます。</span><span class="sxs-lookup"><span data-stu-id="16134-119">Paste the URL you copied as the value for the `"document"` property.</span></span> <span data-ttu-id="16134-120">たとえば、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="16134-120">For example, yours will look something like this:</span></span>

    ```json
      "scripts": {
        "document": "<YOUR URL>",
        ...
      }
    ```

    > [!TIP]
    > <span data-ttu-id="16134-121">Yeoman ジェネレーターを使用しないアドインを作成する場合は、既存の URL に次を追加して、ドキュメントの URL にクエリ パラメーターを追加できます。</span><span class="sxs-lookup"><span data-stu-id="16134-121">If you are creating an add-in not using our Yeoman generator, you can add query parameters to your document's URL, by appending the following to the existing URL:</span></span>

    - <span data-ttu-id="16134-122">開発サーバーのポート。次に例を示します `&wdaddindevserverport=3000` 。</span><span class="sxs-lookup"><span data-stu-id="16134-122">The dev server port, such as `&wdaddindevserverport=3000`.</span></span>
    - <span data-ttu-id="16134-123">マニフェスト ファイル名。次に例を示します `&wdaddinmanifestfile=manifest1.xml` 。</span><span class="sxs-lookup"><span data-stu-id="16134-123">The manifest file name, such as `&wdaddinmanifestfile=manifest1.xml`.</span></span>
    - <span data-ttu-id="16134-124">マニフェスト GUID。次に例を示します `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143` 。</span><span class="sxs-lookup"><span data-stu-id="16134-124">The manifest GUID, such as `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143`.</span></span>

    > <span data-ttu-id="16134-125">Yeoman ジェネレーターを使用している場合、Yeoman ツールによってこの情報が自動的に追加されるので、この情報を追加する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="16134-125">If you are using the Yeoman generator, adding this information is not necessary as the Yeoman tooling appends this information automatically.</span></span>
    > <span data-ttu-id="16134-126">ただし、どちらの場合も、localhost からしかマニフェストを読み込めない点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="16134-126">Note that in both cases, however, you can only load manifests from localhost.</span></span>

3. <span data-ttu-id="16134-127">プロジェクトのルート ディレクトリから始まるコマンド ラインで、次のコマンドを実行します `npm run start:web` 。</span><span class="sxs-lookup"><span data-stu-id="16134-127">In the command line starting at the root directory of your project, run the following command: `npm run start:web`.</span></span>

4. <span data-ttu-id="16134-128">このメソッドを初めて使用して Web 上にアドインをサイドロードすると、開発者モードを有効にしてくださいというダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="16134-128">The first time you use this method to sideload an add-in on the web, you'll see a dialog asking you to enable developer mode.</span></span> <span data-ttu-id="16134-129">[開発者モードを今すぐ有効 **にする] チェック ボックスをオンにして\*\*\*\*、[OK] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="16134-129">Select the checkbox for **Enable Developer Mode now** and select **OK**.</span></span>

5. <span data-ttu-id="16134-130">2 番目のダイアログ ボックスが表示されます。コンピューターからアドイン マニフェストOffice登録を求めるダイアログ ボックスが表示されます。</span><span class="sxs-lookup"><span data-stu-id="16134-130">You will see a second dialog box, asking if you wish to register an Office Add-in manifest from your computer.</span></span> <span data-ttu-id="16134-131">[はい] を選択 **する必要があります**。</span><span class="sxs-lookup"><span data-stu-id="16134-131">You should select **Yes**.</span></span>

6. <span data-ttu-id="16134-132">アドインがインストールされている。</span><span class="sxs-lookup"><span data-stu-id="16134-132">Your add-in is installed.</span></span> <span data-ttu-id="16134-133">アドイン コマンドの場合は、リボンまたはコンテキスト メニューに表示されます。</span><span class="sxs-lookup"><span data-stu-id="16134-133">If it is an add-in command, it should appear on either the ribbon or the context menu.</span></span> <span data-ttu-id="16134-134">作業ウィンドウ アドインの場合は、作業ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="16134-134">If it is a task pane add-in, the task pane should appear.</span></span>

## <a name="sideload-an-office-add-in-in-office-on-the-web-manually"></a><span data-ttu-id="16134-135">Web 上のOfficeアドインを手動Officeサイドロードする</span><span class="sxs-lookup"><span data-stu-id="16134-135">Sideload an Office Add-in in Office on the web manually</span></span>

<span data-ttu-id="16134-136">このメソッドはコマンド ラインを使用し、ホスト アプリケーション (Excel など) 内でのみコマンドを使用して実行できます。</span><span class="sxs-lookup"><span data-stu-id="16134-136">This method doesn't use the command line and can be accomplished using commands only within the host application (such as Excel).</span></span>

1. <span data-ttu-id="16134-137">Web [Officeを開きます](https://office.live.com/)。</span><span class="sxs-lookup"><span data-stu-id="16134-137">Open [Office on the web](https://office.live.com/).</span></span> <span data-ttu-id="16134-138">**Excel、Word、または** PowerPoint でドキュメント **を開きます**。</span><span class="sxs-lookup"><span data-stu-id="16134-138">Open a document in **Excel**, **Word**, or **PowerPoint**.</span></span> <span data-ttu-id="16134-139">[アドイン **]** セクションのリボンの[挿入] タブで、[アドイン] Office **選択します**。</span><span class="sxs-lookup"><span data-stu-id="16134-139">On the **Insert** tab on the ribbon in the **Add-ins** section, choose **Office Add-ins**.</span></span>

1. <span data-ttu-id="16134-140">[アドインOffice] ダイアログで、[**マイ** アドイン] タブを選択し、[マイ アドインの管理] を選択して、[マイ アドインのアップロード] を **選択します**。</span><span class="sxs-lookup"><span data-stu-id="16134-140">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>

    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

1. <span data-ttu-id="16134-142">アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="16134-142">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>

    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

1. <span data-ttu-id="16134-p111">アドインがインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="16134-p111">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
> <span data-ttu-id="16134-147">Microsoft Edge Office元の WebView (EdgeHTML) を使用してアドインをテストするには、追加の構成手順が必要です。</span><span class="sxs-lookup"><span data-stu-id="16134-147">To test your Office Add-in with Microsoft Edge with the original WebView (EdgeHTML), an additional configuration step is required.</span></span> <span data-ttu-id="16134-148">Windows コマンド プロンプトで、次の行を実行します `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes` 。</span><span class="sxs-lookup"><span data-stu-id="16134-148">In a Windows Command Prompt, run the following line: `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`.</span></span> <span data-ttu-id="16134-149">Chromium ベースの Edge WebView2 Office使用している場合、これは必要ありません。</span><span class="sxs-lookup"><span data-stu-id="16134-149">This is not required when Office is using the Chromium-based Edge WebView2.</span></span> <span data-ttu-id="16134-150">詳細については、「アドイン [で使用されるブラウザー」Office参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="16134-150">For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

## <a name="sideload-an-office-add-in"></a><span data-ttu-id="16134-151">新しいアドインOfficeサイドロードする</span><span class="sxs-lookup"><span data-stu-id="16134-151">Sideload an Office Add-in</span></span>

1. <span data-ttu-id="16134-152">Microsoft 365 アカウントにサインインします。</span><span class="sxs-lookup"><span data-stu-id="16134-152">Sign in to your Microsoft 365 account.</span></span>

2. <span data-ttu-id="16134-153">ツールバーの左側起動ツールアプリ アプリを開き **、Excel、Word、** または **PowerPoint** を選択して、新しいドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="16134-153">Open the App Launcher on the left end of the toolbar and select **Excel**, **Word**, or **PowerPoint**, and then create a new document.</span></span>

3. <span data-ttu-id="16134-154">手順 3 から 6 は、前のセクション「**Office on the web で Office アドインをサイドロードする**」のものと同じです。</span><span class="sxs-lookup"><span data-stu-id="16134-154">Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office on the web**.</span></span>

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="16134-155">Visual Studio の使用時にアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="16134-155">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="16134-156">アドインの開発に Visual Studioを使用している場合、サイドロードするプロセスは Web への手動サイドロードに似ています。</span><span class="sxs-lookup"><span data-stu-id="16134-156">If you're using Visual Studio to develop your add-in, the process to sideload is similar to manual sideloading to the web.</span></span> <span data-ttu-id="16134-157">アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。唯一の違いは、マニフェストの **SourceURL** 要素の値を更新して、アドインが展開されている完全な URL を含める必要がある点です。</span><span class="sxs-lookup"><span data-stu-id="16134-157">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="16134-158">アドインは Visual Studio から Office on the web にサイドロードできますが、Visual Studio からはデバッグできません。</span><span class="sxs-lookup"><span data-stu-id="16134-158">Although you can sideload add-ins from Visual Studio to Office on the web, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="16134-159">デバッグするには、ブラウザー デバッグ ツールを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="16134-159">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="16134-160">詳細については、「[Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="16134-160">For more information, see [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="16134-161">Visual Studio で、[**表示**]  >  [**プロパティ ウィンドウ**] の順に選択して [**プロパティ**] ウィンドウを表示させます。</span><span class="sxs-lookup"><span data-stu-id="16134-161">In Visual Studio, show the **Properties** window by choosing **View** > **Properties Window**.</span></span>
2. <span data-ttu-id="16134-162">[**ソリューション エクスプローラー**] で Web プロジェクトを選択します。</span><span class="sxs-lookup"><span data-stu-id="16134-162">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="16134-163">プロジェクトのプロパティが [**プロパティ**] ウィンドウに表示されます。</span><span class="sxs-lookup"><span data-stu-id="16134-163">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="16134-164">[プロパティ] ウィンドウで、[**SSL URL**] をコピーします。</span><span class="sxs-lookup"><span data-stu-id="16134-164">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="16134-165">アドイン プロジェクトで、マニフェスト XML ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="16134-165">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="16134-166">編集しているのがソース XML であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="16134-166">Be sure you are editing the source XML.</span></span> <span data-ttu-id="16134-167">一部の種類のプロジェクトでは、Visual Studio は XML のビジュアル ビューを開きますが、これは次の手順で使用できません。</span><span class="sxs-lookup"><span data-stu-id="16134-167">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="16134-168">**~remoteAppUrl/** のすべてのインスタンスを検索し、先ほどコピーした SSL URL と置き換えます。</span><span class="sxs-lookup"><span data-stu-id="16134-168">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="16134-169">プロジェクトの種類に応じていくつかの置換が表示され、新しい URL の表示は `https://localhost:44300/Home.html` に似たものになりま。</span><span class="sxs-lookup"><span data-stu-id="16134-169">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="16134-170">XML ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="16134-170">Save the XML file.</span></span>
7. <span data-ttu-id="16134-171">Web プロジェクトを右クリックして、[**デバッグ**]  >  [**新しいインスタンスを開始**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="16134-171">Right click the web project and choose **Debug** > **Start new instance**.</span></span> <span data-ttu-id="16134-172">これにより、Office を起動することなく Web プロジェクトが実行されます。</span><span class="sxs-lookup"><span data-stu-id="16134-172">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="16134-173">前述の「[Office on the web で Office アドインをサイドロードする](#sideload-an-office-add-in-in-office-on-the-web)」で説明した手順を使用して、Office on the web からアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="16134-173">From Office on the web, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office on the web](#sideload-an-office-add-in-in-office-on-the-web).</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="16134-174">サイドロードされたアドインを削除する</span><span class="sxs-lookup"><span data-stu-id="16134-174">Remove a sideloaded add-in</span></span>

<span data-ttu-id="16134-175">ブラウザーのキャッシュをクリアすることで、以前にサイドロードされたアドインを削除できます。</span><span class="sxs-lookup"><span data-stu-id="16134-175">You can remove a previously sideloaded add-in by clearing your browser's cache.</span></span> <span data-ttu-id="16134-176">アドインのマニフェストに変更を加えた場合 (たとえば、アイコンのファイル名やアドイン コマンドのテキストを更新する場合 [)、Office](clear-cache.md) キャッシュをクリアしてから、更新されたマニフェストを使用してアドインを再サイドロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="16134-176">If you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you may need to [clear the Office cache](clear-cache.md) and then re-sideload the add-in using the updated manifest.</span></span> <span data-ttu-id="16134-177">これを実行することにより、アドインは更新されたマニフェストの記載どおりに Office で表示されるようになります。</span><span class="sxs-lookup"><span data-stu-id="16134-177">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="see-also"></a><span data-ttu-id="16134-178">関連項目</span><span class="sxs-lookup"><span data-stu-id="16134-178">See also</span></span>

- [<span data-ttu-id="16134-179">iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="16134-179">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="16134-180">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="16134-180">Sideload Outlook add-ins for testing</span></span>](../outlook/sideload-outlook-add-ins-for-testing.md)
- [<span data-ttu-id="16134-181">Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="16134-181">Clear the Office cache</span></span>](clear-cache.md)
