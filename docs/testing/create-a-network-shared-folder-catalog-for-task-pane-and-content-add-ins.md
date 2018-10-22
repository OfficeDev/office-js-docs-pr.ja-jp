---
title: テスト用に Office アドインをサイドロードする
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 6ee8e4e9a2413b34cb8991b09d61e16888a0e6a6
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640023"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="ba881-102">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="ba881-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="ba881-103">マニフェストをネットワーク ファイル共有に公開することで、Windows 上で実行されている Office クライアントにテスト用の Office アドインをインストールできます (以下の手順)。</span><span class="sxs-lookup"><span data-stu-id="ba881-103">You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share.</span></span>

> [!NOTE]
> <span data-ttu-id="ba881-p101">アドイン プロジェクトを [**yo office** ツール](https://github.com/OfficeDev/generator-office)とともにを作成した場合、動作するようにサイドロードする別の方法があります。詳細については、 [sideload コマンドを使用するSideload Office アドイン ](sideload-office-addin-using-sideload-command.md)参照してください。</span><span class="sxs-lookup"><span data-stu-id="ba881-p101">If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you. For details, see [Sideload Office Add-ins using the sideload command](sideload-office-addin-using-sideload-command.md).</span></span>

<span data-ttu-id="ba881-p102">この資料は、Word、Excel、または PowerPoint のアドインの Windows のテストにのみ適用されます。Outlook のアドインをテストするのには別のプラットフォーム上でテストする場合は、アドインをサイドロードするの 次のトピックのいずれかを参照してください。</span><span class="sxs-lookup"><span data-stu-id="ba881-p102">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows. If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="ba881-108">テスト用に Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="ba881-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="ba881-109">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="ba881-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="ba881-110">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="ba881-110">Sideload Outlook add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)


<span data-ttu-id="ba881-111">次のビデオでは、共有フォルダ カタログを使用して Office デスクトップまたは Office Online のアドインをサイドロードする手順を説明します。</span><span class="sxs-lookup"><span data-stu-id="ba881-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="ba881-112">フォルダーの共有</span><span class="sxs-lookup"><span data-stu-id="ba881-112">Share a folder</span></span>

1. <span data-ttu-id="ba881-113">アドインをホストさせようとしている Windows コンピューターで、共有フォルダー カタログとして使用するつもりのフォルダーの親フォルダーまたはドライブ文字に移動します。</span><span class="sxs-lookup"><span data-stu-id="ba881-113">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="ba881-114">共有フォルダー (フォルダーを右クリック) カタログとして使用したいフォルダをコンテキスト メニューから開き、 **プロパティ**を選択します。</span><span class="sxs-lookup"><span data-stu-id="ba881-114">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="ba881-115"> *\*[プロパティ]**  ダイアログ ウィンドウ内で  *\*[共有]**  タブを開き、 *\*[共有]** ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="ba881-115">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![[フォルダー プロパティ] ダイアログ ボックス、[共有] タブと[共有] ボタンが強調表示されます。](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="ba881-117"> *\*ネットワーク アクセス** のダイアログ ウィンドウ内で、アドインを共有したい、他のユーザーおよび/またはグループと自分を追加します。</span><span class="sxs-lookup"><span data-stu-id="ba881-117">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="ba881-118">少なくとも、フォルダーへの**読み取り/書き込み**アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="ba881-118">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="ba881-119">共有するユーザーを選択したら、 **[共有]** ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="ba881-119">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="ba881-120"> *\*フォルダーが共有されました*\*の確認が表示されたら、フォルダー名の直後に表示される完全なネットワーク パスのメモを作成します。</span><span class="sxs-lookup"><span data-stu-id="ba881-120">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="ba881-121">(この資料の次のセクションで説明されているように 、この値を入力する必要とする [信頼できるカタログとして共有フォルダーを指定する場合、](#specify-the-shared-folder-as-a-trusted-catalog)**カタログの Url** としてこの値を入力する必要があります。) **ネットワーク アクセス** のダイアログ ウィンドウを閉じるには、[ **完了** ] を選択します。</span><span class="sxs-lookup"><span data-stu-id="ba881-121">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![強調表示されている共有パスを使用するネットワーク アクセスのダイアログ](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="ba881-123"> *\*[プロパティ** ] ダイアログ ウィンドウを閉じるには、 [ *\*閉じる** ] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="ba881-123">Choose the **Close** button to close the **Workbook Connections** dialog box.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="ba881-124">信頼できるカタログとしてその共有フォルダーを指定します。</span><span class="sxs-lookup"><span data-stu-id="ba881-124">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="ba881-125">Excel、Word、または PowerPoint で新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="ba881-125">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="ba881-126">**[ファイル]** タブを選択して、**[オプション]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ba881-126">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="ba881-127">[**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="ba881-127">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="ba881-128">**[信頼されているアドイン カタログ]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="ba881-128">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="ba881-129"> *\*[カタログの Url]**  ボックスで、以前に[共有](#share-a-folder) したフォルダーの完全なネットワーク パスを入力します。</span><span class="sxs-lookup"><span data-stu-id="ba881-129">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="ba881-130">フォルダーを共有するときにフォルダーの完全なネットワーク パスをメモしていなかった場合に、フォルダーの **[プロパティ]**  ダイアログ ウィンドウで、次のスクリーン ショットが示すように操作して得ることができます。</span><span class="sxs-lookup"><span data-stu-id="ba881-130">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![強調表示されている共有 タブと ネットワーク パスを持つフォルダーのプロパティ ダイアログ ボックス](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="ba881-132"> *\*カタログの Url** ] ボックスに、フォルダーの完全なネットワーク パスを入力した後に、 *\*［カタログの追加]**  ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="ba881-132">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="ba881-133">新しく追加した項目の **[メニューに表示する]** ]チェック ボックスを選択し、 [ **OK** ] ボタンを選択して、[ **セキュリティ センター** ] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="ba881-133">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![カタログが選択されている[トラスト センター] ダイアログ ボックス](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="ba881-135">[ **OK** ] ボタンを選択して、[ **Word のオプション** ] ダイアログ ウィンドウを閉じます</span><span class="sxs-lookup"><span data-stu-id="ba881-135">Choose the  **OK** button to close the **Internet Options** dialog box.</span></span>

9. <span data-ttu-id="ba881-136">Office アプリケーションを閉じるて再度開くと、変更内容が有効になります。</span><span class="sxs-lookup"><span data-stu-id="ba881-136">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="ba881-137">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="ba881-137">Sideload your add-in</span></span>


1. <span data-ttu-id="ba881-138">テストするアドインのマニフェスト XMLファイルを共有フォルダー カタログに置きます。</span><span class="sxs-lookup"><span data-stu-id="ba881-138">Put the manifest file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="ba881-139">Web サーバーに web application自体を展開することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="ba881-139">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="ba881-140"> *\*SourceLocation** 要素マニフェスト ファイルの URL を指定することを確認します。</span><span class="sxs-lookup"><span data-stu-id="ba881-140">Deploy the web application itself to a web server and specify the URL in the  **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="ba881-141">Excel、Word、または PowerPoint で、リボンの **[挿入]** タブにある **[個人用アドイン]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="ba881-141">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="ba881-142">**[Office アドイン]** ダイアログ ボックスの上部にある **[共有フォルダー]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="ba881-142">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="ba881-143">アドインの名前を選び、**[OK]** を選択して、アドインを挿入します。</span><span class="sxs-lookup"><span data-stu-id="ba881-143">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="ba881-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="ba881-144">See also</span></span>

- [<span data-ttu-id="ba881-145">マニフェストの問題を検証し、トラブルシューティングを行う</span><span class="sxs-lookup"><span data-stu-id="ba881-145">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="ba881-146">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="ba881-146">Publish your Office Add-in</span></span>](../publish/publish.md)
    
