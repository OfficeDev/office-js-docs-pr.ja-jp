---
title: テスト用に Office アドインをサイドロードする
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 24c7719969ddc59d8bb6e525af804515331a51ad
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33619045"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="1b343-102">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="1b343-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="1b343-103">共有フォルダー カタログを使用して、マニフェストをネットワークのファイル共有に発行することで、Windows を実行する Office クライアントのテストのために Office アドインをインストールすることができます。</span><span class="sxs-lookup"><span data-stu-id="1b343-103">You can install an Office Add-in for testing in an Office client running on Windows by publishing the manifest to a network file share (instructions below).</span></span>

> [!NOTE]
> <span data-ttu-id="1b343-104">アドイン プロジェクトが [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)で作成された場合、アドインをサイドロードするのに役に立つ別の方法があります。</span><span class="sxs-lookup"><span data-stu-id="1b343-104">If your add-in project was created with the [yo office](https://github.com/OfficeDev/generator-office) tool, there is an alternative way of sideloading it that might work for you.</span></span> <span data-ttu-id="1b343-105">詳細については、「[サイドロード コマンドを使用して Office アドインをサイドロードする](sideload-office-addin-using-sideload-command.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1b343-105">For details, see [Sideload Office Add-ins using the sideload command](sideload-office-addin-using-sideload-command.md).</span></span>

<span data-ttu-id="1b343-106">この記事は、Windows での Word、Excel、PowerPoint、および Project アドインのテストにのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="1b343-106">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows.</span></span> <span data-ttu-id="1b343-107">異なるプラットフォームでのテストまたは Outlook アドインのテストをする場合は、以下の、アドインのサイドロードに関するいずれかのトピックを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1b343-107">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="1b343-108">テスト用に Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="1b343-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="1b343-109">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="1b343-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="1b343-110">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="1b343-110">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

<span data-ttu-id="1b343-111">次のビデオでは、共有フォルダー カタログを使用して Office デスクトップまたは Office Online でアドインをサイドロードする手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="1b343-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online using a shared folder catalog.</span></span>  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a><span data-ttu-id="1b343-112">フォルダーの共有</span><span class="sxs-lookup"><span data-stu-id="1b343-112">Share a folder</span></span>

1. <span data-ttu-id="1b343-113">アドインをホストさせようとしている Windows コンピューターで、共有フォルダー カタログとして使用するつもりのフォルダーの親フォルダーまたはドライブ文字に移動します。</span><span class="sxs-lookup"><span data-stu-id="1b343-113">In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="1b343-114">共有フォルダー カタログとして使用するフォルダーのコンテキスト メニューを開き (フォルダーを右クリック)、[**プロパティ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="1b343-114">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="1b343-115">[**プロパティ**] ダイアログ ボックス内で [**共有**] タブを選択し、[**共有**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="1b343-115">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![[共有] タブと [共有] ボタンが強調表示されているフォルダーの [プロパティ] ダイアログ](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="1b343-117">[**ネットワーク アクセス**] ダイアログ ウィンドウで自分自身とアドインを共有する相手のユーザーまたはグループを追加します。</span><span class="sxs-lookup"><span data-stu-id="1b343-117">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="1b343-118">最低でも、フォルダーへの**読み取り/書き込み**アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="1b343-118">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="1b343-119">共有する相手の選択が完了したら、[**共有**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="1b343-119">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="1b343-120">「**ユーザーのフォルダーは共有されています**」という確認メッセージが表示されたら、フォルダー名のすぐ後に表示される完全なネットワーク パスを書き留めます。</span><span class="sxs-lookup"><span data-stu-id="1b343-120">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="1b343-121">(この記事の次のセクションで説明する通り、[共有フォルダーを信頼できるカタログとして指定する](#specify-the-shared-folder-as-a-trusted-catalog)際に、このネットワーク パスを [**カタログの URL**] として入力する必要があります。) [**完了**] を選択して [**ネットワーク アクセス**] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="1b343-121">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![共有パスが強調表示された [ネットワーク アクセス] ダイアログ](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="1b343-123">[**閉じる**] を選択して、[**プロパティ**] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="1b343-123">Choose the **Close** button to close the **Properties** dialog window.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="1b343-124">共有フォルダーを信頼できるカタログとして指定する</span><span class="sxs-lookup"><span data-stu-id="1b343-124">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="1b343-125">Excel、Word、PowerPoint、または Project で新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="1b343-125">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="1b343-126">[**ファイル**] タブを選択し、[**オプション**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="1b343-126">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="1b343-127">[**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="1b343-127">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="1b343-128">[**信頼されているアドイン カタログ**] を選びます。</span><span class="sxs-lookup"><span data-stu-id="1b343-128">Choose **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="1b343-129">[**カタログの URL**] ボックスで、先ほど[共有](#share-a-folder)したフォルダーの完全なネットワーク パスを入力します。</span><span class="sxs-lookup"><span data-stu-id="1b343-129">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="1b343-130">フォルダーを共有した際に完全なネットワーク パスを書き留めておかなかった場合は、次のスクリーン ショットに示されるように、フォルダーの [**プロパティ**] ダイアログ ウィンドウから取得できます。</span><span class="sxs-lookup"><span data-stu-id="1b343-130">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![[共有] タブとネットワーク パスが強調表示されているフォルダーの [プロパティ] ダイアログ](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="1b343-132">[**カタロ URL**] ボックスにフォルダーの完全なネットワーク パスを入力したら、[**カタログの追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="1b343-132">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="1b343-133">新しく追加されたアイテムの [**メニューに表示する**] チェック ボックスをオンにし、[**OK**] を選択して [**セキュリティ センター** ] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="1b343-133">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![カタログが選択されている [セキュリティ センター] ダイアログ](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="1b343-135">[**OK**] をクリックして [**Word のオプション**] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="1b343-135">Choose the **OK** button to close the **Word Options** dialog window.</span></span>

9. <span data-ttu-id="1b343-136">Office アプリケーションを閉じてからもう一度開くと変更内容が有効になります。</span><span class="sxs-lookup"><span data-stu-id="1b343-136">Close and reopen the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="1b343-137">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="1b343-137">Sideload your add-in</span></span>


1. <span data-ttu-id="1b343-138">テストするアドインのマニフェスト XML ファイルを共有フォルダー カタログに置きます。</span><span class="sxs-lookup"><span data-stu-id="1b343-138">Put the manifest XML file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="1b343-139">なお、Web アプリケーション自体を Web サーバーに展開します。</span><span class="sxs-lookup"><span data-stu-id="1b343-139">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="1b343-140">必ずマニフェスト ファイルの **SourceLocation** 要素で URL を指定してください。</span><span class="sxs-lookup"><span data-stu-id="1b343-140">Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="1b343-141">Excel、Word、または PowerPoint で、リボンの **[挿入]** タブにある **[個人用アドイン]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="1b343-141">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span> <span data-ttu-id="1b343-142">Projectで、リボンの [**Project**]タブの [**個人用アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="1b343-142">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span> 

3. <span data-ttu-id="1b343-143">**[Office アドイン]** ダイアログ ボックスの上部にある **[共有フォルダー]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="1b343-143">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="1b343-144">アドインの名前を選び、**[OK]** を選択して、アドインを挿入します。</span><span class="sxs-lookup"><span data-stu-id="1b343-144">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="1b343-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="1b343-145">See also</span></span>

- [<span data-ttu-id="1b343-146">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="1b343-146">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="1b343-147">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="1b343-147">Publish your Office Add-in</span></span>](../publish/publish.md)
    
