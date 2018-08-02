---
title: テスト用に Office アドインをサイドロードする
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: 42af5d0665fc6cb1135103789adcb4414c4763ff
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703806"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="99a29-102">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="99a29-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="99a29-103">次のいずれかの方法で、Windows 上で実行されている Office クライアントにテスト用の Office アドインをインストールできます。</span><span class="sxs-lookup"><span data-stu-id="99a29-103">You can install an Office Add-in for testing in an Office client running on Windows by one of the following methods:</span></span>

- <span data-ttu-id="99a29-104">共有フォルダ カタログを使用してマニフェストをネットワーク ファイル共有に公開する（以下の手順）</span><span class="sxs-lookup"><span data-stu-id="99a29-104">Using a shared folder catalog to publish the manifest to a network file share (instructions below)</span></span>
- [<span data-ttu-id="99a29-105">アドイン プロジェクト フォルダのルートから「**npm run sideload**」コマンドを実行。</span><span class="sxs-lookup"><span data-stu-id="99a29-105">Running the "**npm run sideload**" command from the root of the add-in project folder.</span></span>](sideload-office-addin-using-sideload-command.md)

    > [!NOTE]
    > <span data-ttu-id="99a29-106">「npm run sideload」メソッドは、Excel、Word、および PowerPoint アドインでのみ機能します）。</span><span class="sxs-lookup"><span data-stu-id="99a29-106">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins).</span></span>

<span data-ttu-id="99a29-107">Word、Excel、PowerPoint のアドインを Windows でテストしない場合は、以下のいずれかのトピックを参照してアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="99a29-107">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="99a29-108">テスト用に Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="99a29-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="99a29-109">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="99a29-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

<span data-ttu-id="99a29-110">次のビデオでは、共有フォルダ カタログを使用して Office デスクトップまたは Office Online のアドインをサイドロードする手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="99a29-110">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="99a29-111">フォルダーの共有</span><span class="sxs-lookup"><span data-stu-id="99a29-111">Share a folder</span></span>

1. <span data-ttu-id="99a29-112">アドインをホストさせようとしている Windows コンピューターで、共有フォルダー カタログとして使用するつもりのフォルダーの親フォルダーまたはドライブ文字に移動します。</span><span class="sxs-lookup"><span data-stu-id="99a29-112">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="99a29-113">フォルダーのコンテキスト メニューを (右クリックして) 開き、**[プロパティ]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="99a29-113">Open the context menu for the folder (right-click) and choose **Properties**.</span></span>

3. <span data-ttu-id="99a29-114">**[共有]** タブを開きます。</span><span class="sxs-lookup"><span data-stu-id="99a29-114">Open the **Sharing** tab.</span></span>

4. <span data-ttu-id="99a29-p101">「**相手を選んでください**」ページで、自分自身とアドインを共有する相手を追加します。相手がセキュリティ グループのメンバー全員の場合は、そのグループを追加できます。少なくとも、フォルダーへの**読み取り/書き込み**アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="99a29-p101">On the **Choose people ...** page, add yourself and and anyone else with whom you want to share your add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder.</span></span> 

5. <span data-ttu-id="99a29-118">**[共有]** > **[完了]** > **[閉じる]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="99a29-118">Choose **Share** > **Done** > **Close**.</span></span>


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="99a29-119">信頼できるカタログとしてその共有フォルダーを指定します。</span><span class="sxs-lookup"><span data-stu-id="99a29-119">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="99a29-120">Excel、Word、または PowerPoint で新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="99a29-120">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="99a29-121">**[ファイル]** タブを選び、**[オプション]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="99a29-121">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="99a29-122">**[セキュリティ センター]** を選び、**[セキュリティ センターの設定]** ボタンを選びます。</span><span class="sxs-lookup"><span data-stu-id="99a29-122">Choose **Trust Center**, and then choose the  **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="99a29-123">**[信頼されているアドイン カタログ]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="99a29-123">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="99a29-124">**[カタログの URL]** ボックスで、共有フォルダー カタログへの完全なネットワーク パスを入力し、**[カタログの追加]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="99a29-124">In the  **Catalog Url** box, enter the full network path to the shared folder catalog, and then choose **Add Catalog**.</span></span>
    
6. <span data-ttu-id="99a29-125">**[メニューに表示する]** チェック ボックスをオンにし、**[OK]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="99a29-125">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

7. <span data-ttu-id="99a29-126">Office アプリケーションを閉じると変更内容が有効になります。</span><span class="sxs-lookup"><span data-stu-id="99a29-126">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="99a29-127">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="99a29-127">Sideload your add-in</span></span>

1. <span data-ttu-id="99a29-p102">テストするアドインのマニフェスト ファイルを共有フォルダー カタログに置きます。なお、Web サーバーに Web アプリケーション自体を展開します。必ずマニフェスト ファイルの **SourceLocation** 要素で URL を指定してください。</span><span class="sxs-lookup"><span data-stu-id="99a29-p102">Put the manifest file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="99a29-131">Excel、Word、または PowerPoint で、リボンの **[挿入]** タブにある **[個人用アドイン]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="99a29-131">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="99a29-132">**[Office アドイン]** ダイアログ ボックスの上部にある **[共有フォルダー]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="99a29-132">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="99a29-133">アドインの名前を選び、**[OK]** を選択して、アドインを挿入します。</span><span class="sxs-lookup"><span data-stu-id="99a29-133">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="99a29-134">関連項目</span><span class="sxs-lookup"><span data-stu-id="99a29-134">See also</span></span>

- [<span data-ttu-id="99a29-135">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="99a29-135">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="99a29-136">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="99a29-136">Publish your Office Add-in</span></span>](../publish/publish.md)
    
