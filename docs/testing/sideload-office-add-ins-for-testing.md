---
title: テスト用に Office Online で Office アドインをサイドロードする
description: Office Online で Office アドインをサイドロードしてテストをする
ms.date: 04/29/2019
localization_priority: Priority
ms.openlocfilehash: 2bcab7b41fa7f5b9590aacc19645253ee822eeb8
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/01/2019
ms.locfileid: "33517087"
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a><span data-ttu-id="3b4ba-103">テスト用に Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="3b4ba-103">Sideload Office Add-ins in Office Online for testing</span></span>

<span data-ttu-id="3b4ba-104">サイドロードを使用することで、最初にアドイン カタログに置かなくても、テスト用に Office アドインをインストールすることができます。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-104">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading.</span></span> <span data-ttu-id="3b4ba-105">サイドロードは、Office 365 または Office Online 上のいずれかで実行できます。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-105">Sideloading can be done in either Office 365 or Office Online.</span></span> <span data-ttu-id="3b4ba-106">2 つのプラットフォームで手順が少し異なります。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-106">The procedure is slightly different for the two platforms.</span></span> 

<span data-ttu-id="3b4ba-107">アドインをサイドロードするとき、アドイン マニフェストはブラウザーのローカル ストレージに格納されます。そのため、ブラウザーのキャッシュを消去したり、別のブラウザーに切り替えたりする場合、アドインを再びサイドロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-107">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>


> [!NOTE]
> <span data-ttu-id="3b4ba-p102">この記事で説明したようにサイドロードは、Word、Excel、および PowerPoint でサポートされています。Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](/outlook/add-ins/sideload-outlook-add-ins-for-testing)」をご参照ください。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

<span data-ttu-id="3b4ba-110">次のビデオでは、Office デスクトップまたは Office Online でアドインをサイドロードする手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-110">The following video walks you through the process of sideloading your add-in in Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-online"></a><span data-ttu-id="3b4ba-111">Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="3b4ba-111">Sideload an Office Add-in in Office Online</span></span>

1. <span data-ttu-id="3b4ba-112">[Microsoft Office Online](https://office.live.com/) を開きます。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-112">Open [Microsoft Office Online](https://office.live.com/).</span></span>
    
2. <span data-ttu-id="3b4ba-113">**[オンライン アプリを今すぐ開始する]** で、 **Excel**、 **Word**、または  **PowerPoint** を選択して、新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-113">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>
    
3. <span data-ttu-id="3b4ba-114">リボンの  **[挿入]** タブを開き、 **[アドイン]** セクションで、 **Office [アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-114">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="3b4ba-115">**[Office アドイン]** ダイアログ ボックスで、**[個人用アドイン]** タブ、**[個人用アドインの管理]**、**[個人用アドインのアップロード]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-115">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="3b4ba-117">アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-117">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

6. <span data-ttu-id="3b4ba-p103">アドインがインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="3b4ba-122">Office アドインを Microsoft Edge でテストするには、2 つの構成手順が必要です。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-122">To test your Office Add-in with Edge, two configuration steps are required:</span></span> 
>
> - <span data-ttu-id="3b4ba-123">Windows コマンド プロンプトで、次のコマンドを実行します: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span><span class="sxs-lookup"><span data-stu-id="3b4ba-123">In a Windows Command Prompt, run the following line: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span></span>
>
> - <span data-ttu-id="3b4ba-124">Microsoft Edge の検索バーに "**about:flags**" と入力して開発者向け設定のオプションを表示させます。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-124">Enter “**about:flags**” in the Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="3b4ba-125">[**ローカルホスト ループバックを許可する**] オプションをオンにし、Microsoft Edge を再起動します。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-125">Check the “**Allow localhost loopback**” option and restart Edge.</span></span>

>    ![[ローカルホスト ループバックを許可する] オプションがオンになった Microsoft Edge。](../images/allow-localhost-loopback.png)


## <a name="sideload-an-office-add-in-in-office-365"></a><span data-ttu-id="3b4ba-127">Office 365 で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="3b4ba-127">Sideload an Office Add-in in Office 365</span></span>

1. <span data-ttu-id="3b4ba-128">Office 365 アカウントにサインインします。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-128">Sign in to your Office 365 account.</span></span>
    
2. <span data-ttu-id="3b4ba-129">ツールバーの左端にあるアプリ起動ツールを開き、**Excel**、**Word**、または **PowerPoint** を選択して、新しいドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-129">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>
    
3. <span data-ttu-id="3b4ba-130">手順 3 ~ 6 は、前のセクションにある**Office Online で Office アドインをサイドロードする**と同じです。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-130">Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office Online**.</span></span>


## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="3b4ba-131">Visual Studio の使用時にアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="3b4ba-131">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="3b4ba-132">アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-132">If you're using Visual Studio to develop your add-in, the process to sideload is similar.</span></span> <span data-ttu-id="3b4ba-133">アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。唯一の違いは、マニフェストの **SourceURL** 要素の値を更新して、アドインが展開されている完全な URL を含める必要がある点です。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-133">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="3b4ba-134">アドインを Visual Studio から Office Online にサイドロードすることはできますが、Visual Studio からはデバッグできません。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-134">Although you can sideload add-ins from Visual Studio to Office Online, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="3b4ba-135">デバッグするには、ブラウザー デバッグ ツールを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-135">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="3b4ba-136">詳細については、「[Office Online でアドインをデバッグする](debug-add-ins-in-office-online.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-136">For more information, see [Debug add-ins in Office Online](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="3b4ba-137">Visual Studio で、[**表示**]  ->  [**プロパティ ウィンドウ**] の順に選択して [**プロパティ**] ウィンドウを表示させます。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-137">In Visual Studio, show the **Properties** window by choosing **View** -> **Properties Window**.</span></span>
2. <span data-ttu-id="3b4ba-138">[**ソリューション エクスプローラー**] で Web プロジェクトを選択します。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-138">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="3b4ba-139">プロジェクトのプロパティが [**プロパティ**] ウィンドウに表示されます。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-139">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="3b4ba-140">[プロパティ] ウィンドウで、[**SSL URL**] をコピーします。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-140">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="3b4ba-141">アドイン プロジェクトで、マニフェスト XML ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-141">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="3b4ba-142">編集しているのがソース XML であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-142">Be sure you are editing the source XML.</span></span> <span data-ttu-id="3b4ba-143">一部の種類のプロジェクトでは、Visual Studio は XML のビジュアル ビューを開きますが、これは次の手順で使用できません。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-143">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="3b4ba-144">**~remoteAppUrl/** のすべてのインスタンスを検索し、先ほどコピーした SSL URL と置き換えます。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-144">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="3b4ba-145">プロジェクトの種類に応じていくつかの置換が表示され、新しい URL の表示は `https://localhost:44300/Home.html` に似たものになりま。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-145">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="3b4ba-146">XML ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-146">Save the XML file.</span></span>
7. <span data-ttu-id="3b4ba-147">Web プロジェクトを右クリックして、[**デバッグ**]  ->  [**新しいインスタンスを開始**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-147">Right click the web project and choose **Debug** -> **Start new instance**.</span></span> <span data-ttu-id="3b4ba-148">これにより、Office を起動することなく Web プロジェクトが実行されます。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-148">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="3b4ba-149">上記の「[Office Online で Office アドインをサイドロードする](#sideload-an-office-add-in-in-office-online)」で説明された手順を使用して、Office Online からアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="3b4ba-149">From Office Online, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office Online](#sideload-an-office-add-in-in-office-online).</span></span>
