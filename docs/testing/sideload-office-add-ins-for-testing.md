---
title: テスト用に Office on the web で Office アドインをサイドロードする
description: Office on the web で Office アドインをサイドロードしてテストをする
ms.date: 02/18/2020
localization_priority: Normal
ms.openlocfilehash: 869cabec737c39d7dded04fe7c52011347e0f314
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163585"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a><span data-ttu-id="f5ac1-103">テスト用に Office on the web で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="f5ac1-103">Sideload Office Add-ins in Office on the web for testing</span></span>

<span data-ttu-id="f5ac1-104">サイドロードを使用することで、最初にアドイン カタログに置かなくても、テスト用に Office アドインをインストールすることができます。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-104">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading.</span></span> <span data-ttu-id="f5ac1-105">サイドロードは、Office 365 または Office on the web で実行できます。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-105">Sideloading can be done in either Office 365 or Office on the web.</span></span> <span data-ttu-id="f5ac1-106">2 つのプラットフォームで手順が少し異なります。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-106">The procedure is slightly different for the two platforms.</span></span>

<span data-ttu-id="f5ac1-107">アドインをサイドロードするとき、アドイン マニフェストはブラウザーのローカル ストレージに格納されます。そのため、ブラウザーのキャッシュを消去したり、別のブラウザーに切り替えたりする場合、アドインを再びサイドロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-107">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>

> [!NOTE]
> <span data-ttu-id="f5ac1-p102">この記事で説明したようにサイドロードは、Word、Excel、および PowerPoint でサポートされています。Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」をご参照ください。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

<span data-ttu-id="f5ac1-110">次のビデオでは、Office on the web またはデスクトップでアドインをサイドロードする手順について説明しています。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-110">The following video walks you through the process of sideloading your add-in in Office on the web or desktop.</span></span>

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a><span data-ttu-id="f5ac1-111">Office on the web で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="f5ac1-111">Sideload an Office Add-in in Office on the web</span></span>

1. <span data-ttu-id="f5ac1-112">[Microsoft Office on the web](https://office.live.com/) を開きます。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-112">Open [Microsoft Office on the web](https://office.live.com/).</span></span>

2. <span data-ttu-id="f5ac1-113">**[オンライン アプリを今すぐ開始する]** で、 **Excel**、 **Word**、または  **PowerPoint** を選択して、新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-113">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>

3. <span data-ttu-id="f5ac1-114">リボンの  **[挿入]** タブを開き、 **[アドイン]** セクションで、 **Office [アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-114">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>

4. <span data-ttu-id="f5ac1-115">**[Office アドイン]** ダイアログ ボックスで、**[個人用アドイン]** タブ、**[個人用アドインの管理]**、**[個人用アドインのアップロード]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-115">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>

    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="f5ac1-117">アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-117">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>

    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

6. <span data-ttu-id="f5ac1-p103">アドインがインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="f5ac1-122">Microsoft Edge で Office アドインをテストするには、次の 2 つの構成手順が必要です。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-122">To test your Office Add-in with Microsoft Edge, two configuration steps are required:</span></span> 
>
> - <span data-ttu-id="f5ac1-123">Windows コマンド プロンプトで、次のコマンドを実行します: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span><span class="sxs-lookup"><span data-stu-id="f5ac1-123">In a Windows Command Prompt, run the following line: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span></span>
>
> - <span data-ttu-id="f5ac1-124">Microsoft Edge の検索バーに "**about:flags**" と入力して開発者向け設定のオプションを表示します。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-124">Enter “**about:flags**” in the Microsoft Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="f5ac1-125">**[ローカルホスト ループバックを許可する]** オプションをオンにして、Microsoft Edge を再起動します。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-125">Check the “**Allow localhost loopback**” option and restart Microsoft Edge.</span></span>

>    ![[ローカルホスト ループバックを許可する] オプションがオンになった Microsoft Edge。](../images/allow-localhost-loopback.png)

## <a name="sideload-an-office-add-in-in-office-365"></a><span data-ttu-id="f5ac1-127">Office 365 で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="f5ac1-127">Sideload an Office Add-in in Office 365</span></span>

1. <span data-ttu-id="f5ac1-128">Office 365 アカウントにサインインします。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-128">Sign in to your Office 365 account.</span></span>

2. <span data-ttu-id="f5ac1-129">ツールバーの左端にあるアプリ起動ツールを開き、**Excel**、**Word**、または **PowerPoint** を選択して、新しいドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-129">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>

3. <span data-ttu-id="f5ac1-130">手順 3 から 6 は、前のセクション「**Office on the web で Office アドインをサイドロードする**」のものと同じです。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-130">Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office on the web**.</span></span>

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="f5ac1-131">Visual Studio の使用時にアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="f5ac1-131">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="f5ac1-132">アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-132">If you're using Visual Studio to develop your add-in, the process to sideload is similar.</span></span> <span data-ttu-id="f5ac1-133">アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。唯一の違いは、マニフェストの **SourceURL** 要素の値を更新して、アドインが展開されている完全な URL を含める必要がある点です。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-133">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="f5ac1-134">アドインは Visual Studio から Office on the web にサイドロードできますが、Visual Studio からはデバッグできません。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-134">Although you can sideload add-ins from Visual Studio to Office on the web, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="f5ac1-135">デバッグするには、ブラウザー デバッグ ツールを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-135">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="f5ac1-136">詳細については、「[Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-136">For more information, see [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="f5ac1-137">Visual Studio で、[**表示**]  ->  [**プロパティ ウィンドウ**] の順に選択して [**プロパティ**] ウィンドウを表示させます。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-137">In Visual Studio, show the **Properties** window by choosing **View** -> **Properties Window**.</span></span>
2. <span data-ttu-id="f5ac1-138">[**ソリューション エクスプローラー**] で Web プロジェクトを選択します。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-138">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="f5ac1-139">プロジェクトのプロパティが [**プロパティ**] ウィンドウに表示されます。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-139">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="f5ac1-140">[プロパティ] ウィンドウで、[**SSL URL**] をコピーします。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-140">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="f5ac1-141">アドイン プロジェクトで、マニフェスト XML ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-141">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="f5ac1-142">編集しているのがソース XML であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-142">Be sure you are editing the source XML.</span></span> <span data-ttu-id="f5ac1-143">一部の種類のプロジェクトでは、Visual Studio は XML のビジュアル ビューを開きますが、これは次の手順で使用できません。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-143">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="f5ac1-144">**~remoteAppUrl/** のすべてのインスタンスを検索し、先ほどコピーした SSL URL と置き換えます。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-144">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="f5ac1-145">プロジェクトの種類に応じていくつかの置換が表示され、新しい URL の表示は `https://localhost:44300/Home.html` に似たものになりま。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-145">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="f5ac1-146">XML ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-146">Save the XML file.</span></span>
7. <span data-ttu-id="f5ac1-147">Web プロジェクトを右クリックして、[**デバッグ**]  ->  [**新しいインスタンスを開始**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-147">Right click the web project and choose **Debug** -> **Start new instance**.</span></span> <span data-ttu-id="f5ac1-148">これにより、Office を起動することなく Web プロジェクトが実行されます。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-148">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="f5ac1-149">前述の「[Office on the web で Office アドインをサイドロードする](#sideload-an-office-add-in-in-office-on-the-web)」で説明した手順を使用して、Office on the web からアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-149">From Office on the web, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office on the web](#sideload-an-office-add-in-in-office-on-the-web).</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="f5ac1-150">サイドロードアドインを削除する</span><span class="sxs-lookup"><span data-stu-id="f5ac1-150">Remove a sideloaded add-in</span></span>

<span data-ttu-id="f5ac1-151">以前のサイドロードアドインを削除するには、ブラウザーのキャッシュをクリアする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-151">You can remove a previously sideloaded add-in by clearing your browser's cache.</span></span> <span data-ttu-id="f5ac1-152">また、アドインのマニフェストを変更した場合 (たとえば、アイコンの更新ファイル名やアドインコマンドのテキスト)、キャッシュをクリアし、更新されたマニフェストを使用してアドインを再サイドロードする必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-152">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you may need to clear the cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="f5ac1-153">これを実行することにより、アドインは更新されたマニフェストの記載どおりに Office で表示されるようになります。</span><span class="sxs-lookup"><span data-stu-id="f5ac1-153">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>
