---
title: テスト用に Office Online で Office アドインをサイドロードする
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 69b255545525ff667618c9f8bd1e1b7953592967
ms.sourcegitcommit: 58af795c3d0393a4b1f6425fa1cbdca1e48fb473
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/29/2018
ms.locfileid: "20138850"
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a><span data-ttu-id="0dca3-102">テスト用に Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="0dca3-102">Sideload Office Add-ins in Office Online for testing</span></span>

<span data-ttu-id="0dca3-p101">まずアドイン カタログに置かなくても、サイドロードを使用すると、テスト用に Office アドインをインストールすることができます。サイドロードは、Office 365 または Office Online 上のいずれかで実行できます。2 つのプラットフォームで手順が少し異なります。</span><span class="sxs-lookup"><span data-stu-id="0dca3-p101">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading. Sideloading can be done on either Office 365 or Office Online. The procedure is slightly different for the two platforms.</span></span> 

<span data-ttu-id="0dca3-106">アドインをサイドロードするとき、アドイン マニフェストはブラウザーのローカル ストレージに格納されます。そのため、ブラウザーのキャッシュを消去したり、別のブラウザーに切り替えたりする場合、アドインを再びサイドロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0dca3-106">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>


> [!NOTE]
> <span data-ttu-id="0dca3-p102">この記事で説明したようにサイドロードは、Word、Excel、および PowerPoint でサポートされています。Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing)」をご参照ください。</span><span class="sxs-lookup"><span data-stu-id="0dca3-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

<span data-ttu-id="0dca3-109">次のビデオでは、Office デスクトップまたは Office Online のアドインをサイドロードする手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="0dca3-109">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-on-office-365"></a><span data-ttu-id="0dca3-110">Office アドインを Office 365 にサイドロードする</span><span class="sxs-lookup"><span data-stu-id="0dca3-110">Sideload an Office Add-in on Office 365</span></span>


1. <span data-ttu-id="0dca3-111">Office 365 サイトにサインインします。</span><span class="sxs-lookup"><span data-stu-id="0dca3-111">Sign in to your Office 365 account.</span></span>
    
2. <span data-ttu-id="0dca3-112">ツールバーの左端にあるアプリ起動ツールを開き、**Excel**、**Word**、または **PowerPoint** を選択して、新しいドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="0dca3-112">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>
    
3. <span data-ttu-id="0dca3-113">リボンの  **[挿入]** タブを開き、 **[アドイン]** セクションで、 **Office [アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0dca3-113">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="0dca3-114">**[Office アドイン]** ダイアログ ボックスで、**[自分の所属組織]** タブ、**[個人用アドインのアップロード]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="0dca3-114">On the  **Office Add-ins** dialog, select the **MY ORGANIZATION** tab, and then **Upload My Add-in**.</span></span>
    
    ![左上隅近くの、リンクが付いている Office アドインのダイアログ。タイトルは、[マイ アドインのアップロード]](../images/office-add-ins.png)

5.  <span data-ttu-id="0dca3-116">アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0dca3-116">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ](../images/upload-add-in.png)

6. <span data-ttu-id="0dca3-p103">アドイン がインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0dca3-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.</span></span>
    

## <a name="sideload-an-office-add-in-on-office-online"></a><span data-ttu-id="0dca3-121">Office アドイン を Office Online にサイドロードする</span><span class="sxs-lookup"><span data-stu-id="0dca3-121">Sideload an Office Add-in on Office Online</span></span>


1. <span data-ttu-id="0dca3-122">[Microsoft Office Online](https://office.live.com/) を開きます。</span><span class="sxs-lookup"><span data-stu-id="0dca3-122">Open [Microsoft Office Online](https://office.live.com/).</span></span>
    
2. <span data-ttu-id="0dca3-123">**[オンライン アプリを今すぐ開始する]** で、 **Excel**、 **Word**、または  **PowerPoint** を選択して、新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="0dca3-123">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>
    
3. <span data-ttu-id="0dca3-124">リボンの  **[挿入]** タブを開き、 **[アドイン]** セクションで、 **Office [アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0dca3-124">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="0dca3-125">**[Office アドイン]** ダイアログ ボックスで、**[個人用アドイン]** タブ、**[個人用アドインの管理]**、**[個人用アドインのアップロード]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="0dca3-125">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="0dca3-127">アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0dca3-127">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

6. <span data-ttu-id="0dca3-p104">アドインがインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0dca3-p104">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="0dca3-132">Office アドインを Edge でテストするには、Edge の検索バーに "**abou:flags**" を入力し、[開発者設定] オプションを表示します。</span><span class="sxs-lookup"><span data-stu-id="0dca3-132">To test your Office Add-in with Edge, enter “**about:flags**” in the Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="0dca3-133">"**ローカルホスト ループバックを許可する**" オプションにチェックを入れ、Edgeを再起動します。</span><span class="sxs-lookup"><span data-stu-id="0dca3-133">Check the “**Allow localhost loopback**” option and restart Edge.</span></span>

>    ![Edge の [ローカルホスト ループバックを許可する] オプションにチェックを入れます。](../images/allow-localhost-loopback.png)

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="0dca3-135">Visual Studio の使用時にアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="0dca3-135">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="0dca3-p106">アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。唯一の違いは、マニフェストの **SourceURL** 要素の値を更新して、アドインが展開されている完全な URL を含める必要がある点です。</span><span class="sxs-lookup"><span data-stu-id="0dca3-p106">If you're using Visual Studio to develop your add-in, the process to sideload is similar. The only difference is that you will have to update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span> 

<span data-ttu-id="0dca3-p107">現在アドインを開発している場合、アドイン manifest.xml ファイルを見つけて、**SourceLocation** 要素の値を更新することにより、絶対 URI を含めます。Visual Studio は、localhost を展開するためのトークンを配置します。</span><span class="sxs-lookup"><span data-stu-id="0dca3-p107">If you're currently developing your add-in, locate your add-in manifest.xml file, and update the **SourceLocation** element value to include an absolute URI. Visual Studio will put in a token for your localhost deployment.</span></span>

<span data-ttu-id="0dca3-140">例:</span><span class="sxs-lookup"><span data-stu-id="0dca3-140">For example:</span></span> 

```xml
<SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
```
