---
title: テスト用に Office Online で Office アドインをサイドロードする
description: Office Online で Office アドインをサイドロードしてテストをする
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 8870e955ca30c4a3b35f2b51e0e16a3ee634960d
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871718"
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a><span data-ttu-id="2c00c-103">テスト用に Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="2c00c-103">Sideload Office Add-ins in Office Online for testing</span></span>

<span data-ttu-id="2c00c-104">サイドロードを使用することで、最初にアドイン カタログに置かなくても、テスト用に Office アドインをインストールすることができます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-104">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading.</span></span> <span data-ttu-id="2c00c-105">サイドロードは、Office 365 または Office Online 上のいずれかで実行できます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-105">Sideloading can be done in either Office 365 or Office Online.</span></span> <span data-ttu-id="2c00c-106">2 つのプラットフォームで手順が少し異なります。</span><span class="sxs-lookup"><span data-stu-id="2c00c-106">The procedure is slightly different for the two platforms.</span></span> 

<span data-ttu-id="2c00c-107">アドインをサイドロードするとき、アドイン マニフェストはブラウザーのローカル ストレージに格納されます。そのため、ブラウザーのキャッシュを消去したり、別のブラウザーに切り替えたりする場合、アドインを再びサイドロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="2c00c-107">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>


> [!NOTE]
> <span data-ttu-id="2c00c-p102">この記事で説明したようにサイドロードは、Word、Excel、および PowerPoint でサポートされています。Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](/outlook/add-ins/sideload-outlook-add-ins-for-testing)」をご参照ください。</span><span class="sxs-lookup"><span data-stu-id="2c00c-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

<span data-ttu-id="2c00c-110">次のビデオでは、Office デスクトップまたは Office Online でアドインをサイドロードする手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-110">The following video walks you through the process of sideloading your add-in in Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-365"></a><span data-ttu-id="2c00c-111">Office 365 で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="2c00c-111">Sideload an Office Add-in in Office 365</span></span>


1. <span data-ttu-id="2c00c-112">Office 365 アカウントにサインインします。</span><span class="sxs-lookup"><span data-stu-id="2c00c-112">Sign in to your Office 365 account.</span></span>
    
2. <span data-ttu-id="2c00c-113">ツールバーの左端にあるアプリ起動ツールを開き、**Excel**、**Word**、または **PowerPoint** を選択して、新しいドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-113">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>
    
3. <span data-ttu-id="2c00c-114">リボンの  **[挿入]** タブを開き、 **[アドイン]** セクションで、 **Office [アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-114">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="2c00c-115">**[Office アドイン]** ダイアログ ボックスで、**[自分の所属組織]** タブ、**[個人用アドインのアップロード]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-115">On the  **Office Add-ins** dialog, select the **MY ORGANIZATION** tab, and then **Upload My Add-in**.</span></span>
    
    ![左上隅近くの、リンクが付いている Office アドインのダイアログ。タイトルは、[マイ アドインのアップロード]](../images/office-add-ins.png)

5.  <span data-ttu-id="2c00c-117">アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-117">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ](../images/upload-add-in.png)

6. <span data-ttu-id="2c00c-p103">アドイン がインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.</span></span>
    

## <a name="sideload-an-office-add-in-in-office-online"></a><span data-ttu-id="2c00c-122">Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="2c00c-122">Sideload an Office Add-in in Office Online</span></span>


1. <span data-ttu-id="2c00c-123">[Microsoft Office Online](https://office.live.com/) を開きます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-123">Open [Microsoft Office Online](https://office.live.com/).</span></span>
    
2. <span data-ttu-id="2c00c-124">**[オンライン アプリを今すぐ開始する]** で、 **Excel**、 **Word**、または  **PowerPoint** を選択して、新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-124">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>
    
3. <span data-ttu-id="2c00c-125">リボンの  **[挿入]** タブを開き、 **[アドイン]** セクションで、 **Office [アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-125">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="2c00c-126">**[Office アドイン]** ダイアログ ボックスで、**[個人用アドイン]** タブ、**[個人用アドインの管理]**、**[個人用アドインのアップロード]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-126">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="2c00c-128">アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-128">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

6. <span data-ttu-id="2c00c-p104">アドインがインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-p104">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="2c00c-133">Office アドインを Microsoft Edge でテストするには、2 つの構成手順が必要です。</span><span class="sxs-lookup"><span data-stu-id="2c00c-133">To test your Office Add-in with Edge, two configuration steps are required:</span></span> 
>
> - <span data-ttu-id="2c00c-134">Windows コマンド プロンプトで、次のコマンドを実行します: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span><span class="sxs-lookup"><span data-stu-id="2c00c-134">In a Windows Command Prompt, run the following line: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span></span>
>
> - <span data-ttu-id="2c00c-135">Microsoft Edge の検索バーに "**about:flags**" と入力して開発者向け設定のオプションを表示させます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-135">Enter “**about:flags**” in the Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="2c00c-136">[**ローカルホスト ループバックを許可する**] オプションをオンにし、Microsoft Edge を再起動します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-136">Check the “**Allow localhost loopback**” option and restart Edge.</span></span>

>    ![[ローカルホスト ループバックを許可する] オプションがオンになった Microsoft Edge。](../images/allow-localhost-loopback.png)

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="2c00c-138">Visual Studio の使用時にアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="2c00c-138">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="2c00c-139">アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。</span><span class="sxs-lookup"><span data-stu-id="2c00c-139">If you're using Visual Studio to develop your add-in, the process to sideload is similar.</span></span> <span data-ttu-id="2c00c-140">アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。唯一の違いは、マニフェストの **SourceURL** 要素の値を更新して、アドインが展開されている完全な URL を含める必要がある点です。</span><span class="sxs-lookup"><span data-stu-id="2c00c-140">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="2c00c-141">アドインを Visual Studio から Office Online にサイドロードすることはできますが、Visual Studio からはデバッグできません。</span><span class="sxs-lookup"><span data-stu-id="2c00c-141">Although you can sideload add-ins from Visual Studio to Office Online, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="2c00c-142">デバッグするには、ブラウザー デバッグ ツールを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2c00c-142">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="2c00c-143">詳細については、「[Office Online でアドインをデバッグする](debug-add-ins-in-office-online.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2c00c-143">For more information, see [Debug add-ins in Office Online](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="2c00c-144">Visual Studio で、[**表示**]  ->  [**プロパティ ウィンドウ**] の順に選択して [**プロパティ**] ウィンドウを表示させます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-144">In Visual Studio, show the **Properties** window by choosing **View** -> **Properties Window**.</span></span>
2. <span data-ttu-id="2c00c-145">[**ソリューション エクスプローラー**] で Web プロジェクトを選択します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-145">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="2c00c-146">プロジェクトのプロパティが [**プロパティ**] ウィンドウに表示されます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-146">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="2c00c-147">[プロパティ] ウィンドウで、[**SSL URL**] をコピーします。</span><span class="sxs-lookup"><span data-stu-id="2c00c-147">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="2c00c-148">アドイン プロジェクトで、マニフェスト XML ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-148">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="2c00c-149">編集しているのがソース XML であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-149">Be sure you are editing the source XML.</span></span> <span data-ttu-id="2c00c-150">一部の種類のプロジェクトでは、Visual Studio は XML のビジュアル ビューを開きますが、これは次の手順で使用できません。</span><span class="sxs-lookup"><span data-stu-id="2c00c-150">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="2c00c-151">**~remoteAppUrl/** のすべてのインスタンスを検索し、先ほどコピーした SSL URL と置き換えます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-151">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="2c00c-152">プロジェクトの種類に応じていくつかの置換が表示され、新しい URL の表示は `https://localhost:44300/Home.html` に似たものになりま。</span><span class="sxs-lookup"><span data-stu-id="2c00c-152">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="2c00c-153">XML ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-153">Save the XML file.</span></span>
7. <span data-ttu-id="2c00c-154">Web プロジェクトを右クリックして、[**デバッグ**]  ->  [**新しいインスタンスを開始**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="2c00c-154">Right click the web project and choose **Debug** -> **Start new instance**.</span></span> <span data-ttu-id="2c00c-155">これにより、Office を起動することなく Web プロジェクトが実行されます。</span><span class="sxs-lookup"><span data-stu-id="2c00c-155">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="2c00c-156">上記の「[Office Online で Office アドインをサイドロードする](#sideload-an-office-add-in-in-office-online)」で説明された手順を使用して、Office Online からアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="2c00c-156">From Office Online, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office Online](#sideload-an-office-add-in-in-office-online).</span></span>
