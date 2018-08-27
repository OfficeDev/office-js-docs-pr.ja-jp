---
title: テスト用に Office アドインをサイドロードする
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: 94803a2c610fc869aefb6c77d53965981778e62e
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925123"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="e6fd7-102">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="e6fd7-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="e6fd7-103">マニフェストをネットワーク ファイル共有に公開することで、Windows 上で実行されている Office クライアントにテスト用の Office アドインをインストールできます (以下の手順)。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-103">You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share.</span></span>

> [!NOTE]
> <span data-ttu-id="e6fd7-104">[**yo office **ツール](https://github.com/OfficeDev/generator-office)を使用してアドイン プロジェクトを作成した場合、お客様に適した別のサイドロードの方法があります。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-104">If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you.</span></span> <span data-ttu-id="e6fd7-105">詳細は、[sideload コマンドを使用した Sideload Office アドイン](sideload-office-addin-using-sideload-command.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-105">Sideload Office Add-ins using the sideload command</span></span>

<span data-ttu-id="e6fd7-106">この記事は、Windows 上の Word、Excel、または PowerPoint アドインのテストにのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-106">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows.</span></span> <span data-ttu-id="e6fd7-107">別のプラットフォームでテストする場合、または Outlook アドインをテストする場合は、次のトピックのいずれかを参照してアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-107">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="e6fd7-108">テスト用に Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="e6fd7-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="e6fd7-109">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="e6fd7-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="e6fd7-110">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="e6fd7-110">Sideload Outlook Add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)


<span data-ttu-id="e6fd7-111">次のビデオでは、共有フォルダ カタログを使用して Office デスクトップまたは Office Online のアドインをサイドロードする手順を説明します。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="e6fd7-112">フォルダーの共有</span><span class="sxs-lookup"><span data-stu-id="e6fd7-112">Share a folder</span></span>

1. <span data-ttu-id="e6fd7-113">アドインをホストさせようとしている Windows コンピューターで、共有フォルダー カタログとして使用するつもりのフォルダーの親フォルダーまたはドライブ文字に移動します。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-113">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="e6fd7-114">フォルダーのコンテキスト メニューを (右クリックして) 開き、**[プロパティ]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-114">Open the context menu for the folder (right-click) and choose **Properties**.</span></span>

3. <span data-ttu-id="e6fd7-115">**[共有]** タブを開きます。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-115">Open the **Sharing** tab.</span></span>

4. <span data-ttu-id="e6fd7-p103">「**相手を選んでください**」ページで、自分自身とアドインを共有する相手を追加します。相手がセキュリティ グループのメンバー全員の場合は、そのグループを追加できます。少なくとも、フォルダーへの**読み取り/書き込み**アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-p103">On the **Choose people ...** page, add yourself and and anyone else with whom you want to share your add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder.</span></span> 

5. <span data-ttu-id="e6fd7-119">**[共有]** > **[完了]** > **[閉じる]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-119">Choose **Share** > **Done** > **Close**.</span></span>


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="e6fd7-120">信頼できるカタログとしてその共有フォルダーを指定します。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-120">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="e6fd7-121">Excel、Word、または PowerPoint で新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-121">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="e6fd7-122">**[ファイル]** タブを選び、**[オプション]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-122">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="e6fd7-123">**[セキュリティ センター]** を選び、**[セキュリティ センターの設定]** ボタンを選びます。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-123">Choose **Trust Center**, and then choose the  **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="e6fd7-124">**[信頼されているアドイン カタログ]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-124">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="e6fd7-125">**[カタログの URL]** ボックスで、共有フォルダー カタログへの完全なネットワーク パスを入力し、**[カタログの追加]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-125">In the  **Catalog Url** box, enter the full network path to the shared folder catalog, and then choose **Add Catalog**.</span></span>
    
6. <span data-ttu-id="e6fd7-126">**[メニューに表示する]** チェック ボックスをオンにし、**[OK]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-126">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

7. <span data-ttu-id="e6fd7-127">Office アプリケーションを閉じると変更内容が有効になります。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-127">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="e6fd7-128">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="e6fd7-128">Sideload your add-in</span></span>


1. <span data-ttu-id="e6fd7-p104">テストするアドインのマニフェスト XML ファイルを共有フォルダー カタログに置きます。Web アプリケーションを Web サーバーに展開します。必ずマニフェスト ファイルの **SourceLocation** 要素で URL を指定してください。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-p104">Put the manifest file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="e6fd7-132">Excel、Word、または PowerPoint で、リボンの **[挿入]** タブにある **[個人用アドイン]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-132">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="e6fd7-133">**[Office アドイン]** ダイアログ ボックスの上部にある **[共有フォルダー]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-133">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="e6fd7-134">アドインの名前を選び、**[OK]** を選択して、アドインを挿入します。</span><span class="sxs-lookup"><span data-stu-id="e6fd7-134">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="e6fd7-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="e6fd7-135">See also</span></span>

- [<span data-ttu-id="e6fd7-136">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="e6fd7-136">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="e6fd7-137">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="e6fd7-137">Publish your Office Add-in</span></span>](../publish/publish.md)
    
