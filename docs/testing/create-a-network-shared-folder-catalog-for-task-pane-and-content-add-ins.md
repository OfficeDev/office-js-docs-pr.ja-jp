---
title: テスト用に Office アドインをサイドロードする
description: ''
ms.date: 12/06/2019
localization_priority: Priority
ms.openlocfilehash: bb926b09d9381574d22e7634a578adac141e1f8f
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814480"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="e01f1-102">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="e01f1-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="e01f1-103">共有フォルダー カタログを使用して、マニフェストをネットワークのファイル共有に発行することで、Windows を実行する Office クライアントのテストのために Office アドインをインストールすることができます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-103">You can install an Office Add-in for testing in an Office client running on Windows by publishing the manifest to a network file share (instructions below).</span></span>

> [!NOTE]
> <span data-ttu-id="e01f1-104">アドイン プロジェクトが十分に新しい [Office 用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office) バージョンで作成されている場合、アドインは `npm start` を実行すると自動的に Office デスクトップ クライアントにサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="e01f1-104">If your add-in project was created with a sufficiently recent version of the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), the add-in will automatically sideload in the Office desktop client when you run `npm start`.</span></span>

<span data-ttu-id="e01f1-105">この記事は、Windows での Word、Excel、PowerPoint、および Project アドインのテストにのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-105">This article applies only to testing Word, Excel, PowerPoint, and Project add-ins on Windows.</span></span> <span data-ttu-id="e01f1-106">異なるプラットフォームでのテストまたは Outlook アドインのテストをする場合は、以下の、アドインのサイドロードに関するいずれかのトピックを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e01f1-106">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="e01f1-107">テスト用に Office on the web で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="e01f1-107">Sideload Office Add-ins in Office on the web for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="e01f1-108">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="e01f1-108">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="e01f1-109">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="e01f1-109">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

<span data-ttu-id="e01f1-110">次のビデオでは、共有フォルダー カタログを使用して Office on the web またはデスクトップでアドインをサイドロードする手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-110">The following video walks you through the process of sideloading your add-in in Office on the web or desktop using a shared folder catalog.</span></span>  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a><span data-ttu-id="e01f1-111">フォルダーの共有</span><span class="sxs-lookup"><span data-stu-id="e01f1-111">Share a folder</span></span>

1. <span data-ttu-id="e01f1-112">アドインをホストさせようとしている Windows コンピューターで、共有フォルダー カタログとして使用するつもりのフォルダーの親フォルダーまたはドライブ文字に移動します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-112">In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="e01f1-113">共有フォルダー カタログとして使用するフォルダーのコンテキスト メニューを開き (フォルダーを右クリック)、[**プロパティ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-113">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="e01f1-114">[**プロパティ**] ダイアログ ボックス内で [**共有**] タブを選択し、[**共有**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-114">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![[共有] タブと [共有] ボタンが強調表示されているフォルダーの [プロパティ] ダイアログ](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="e01f1-116">[**ネットワーク アクセス**] ダイアログ ウィンドウで自分自身とアドインを共有する相手のユーザーまたはグループを追加します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-116">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="e01f1-117">最低でも、フォルダーへの**読み取り/書き込み**アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="e01f1-117">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="e01f1-118">共有する相手の選択が完了したら、[**共有**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-118">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="e01f1-119">「**ユーザーのフォルダーは共有されています**」という確認メッセージが表示されたら、フォルダー名のすぐ後に表示される完全なネットワーク パスを書き留めます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-119">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="e01f1-120">(この記事の次のセクションで説明する通り、[共有フォルダーを信頼できるカタログとして指定する](#specify-the-shared-folder-as-a-trusted-catalog)際に、このネットワーク パスを [**カタログの URL**] として入力する必要があります。) [**完了**] を選択して [**ネットワーク アクセス**] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-120">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![共有パスが強調表示された [ネットワーク アクセス] ダイアログ](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="e01f1-122">[**閉じる**] を選択して、[**プロパティ**] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-122">Choose the **Close** button to close the **Properties** dialog window.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="e01f1-123">共有フォルダーを信頼できるカタログとして指定する</span><span class="sxs-lookup"><span data-stu-id="e01f1-123">Specify the shared folder as a trusted catalog</span></span> 

### <a name="configure-the-trust-manually"></a><span data-ttu-id="e01f1-124">信頼を手動で構成する</span><span class="sxs-lookup"><span data-stu-id="e01f1-124">Configure the trust manually</span></span>
      
1. <span data-ttu-id="e01f1-125">Excel、Word、PowerPoint、または Project で新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-125">Open a new document in Excel, Word, PowerPoint, or Project.</span></span>
    
2. <span data-ttu-id="e01f1-126">[**ファイル**] タブを選択し、[**オプション**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-126">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="e01f1-127">[**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-127">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="e01f1-128">[**信頼されているアドイン カタログ**] を選びます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-128">Choose **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="e01f1-129">[**カタログの URL**] ボックスで、先ほど[共有](#share-a-folder)したフォルダーの完全なネットワーク パスを入力します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-129">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="e01f1-130">フォルダーを共有した際に完全なネットワーク パスを書き留めておかなかった場合は、次のスクリーン ショットに示されるように、フォルダーの [**プロパティ**] ダイアログ ウィンドウから取得できます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-130">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![[共有] タブとネットワーク パスが強調表示されているフォルダーの [プロパティ] ダイアログ](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="e01f1-132">[**カタロ URL**] ボックスにフォルダーの完全なネットワーク パスを入力したら、[**カタログの追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-132">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="e01f1-133">新しく追加されたアイテムの [**メニューに表示する**] チェック ボックスをオンにし、[**OK**] を選択して [**セキュリティ センター** ] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-133">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![カタログが選択されている [セキュリティ センター] ダイアログ](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="e01f1-135">[**OK**] をクリックして [**Word のオプション**] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-135">Choose the **OK** button to close the **Word Options** dialog window.</span></span>

9. <span data-ttu-id="e01f1-136">Office アプリケーションを閉じてからもう一度開くと変更内容が有効になります。</span><span class="sxs-lookup"><span data-stu-id="e01f1-136">Close and reopen the Office application so your changes will take effect.</span></span>

### <a name="configure-the-trust-with-a-registry-script"></a><span data-ttu-id="e01f1-137">レジストリ スクリプトを使用して信頼を構成する</span><span class="sxs-lookup"><span data-stu-id="e01f1-137">Configure the trust with a Registry script</span></span>

1. <span data-ttu-id="e01f1-138">テキスト エディターで、TrustNetworkShareCatalog.reg という名前のファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-138">In a text editor, such as Notepad, create a file named ItemMetadata.xml.</span></span> 

2. <span data-ttu-id="e01f1-139">次に示すコンテンツをファイルに追加します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-139">Add the following content to the file:</span></span>

    ```
    Windows Registry Editor Version 5.00
    
    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-random-GUID-here-}]
    "Id"="{-random-GUID-here-}"
    "Url"="\\\\-share-\\-folder-"
    "Flags"=dword:00000001
    ```
3. <span data-ttu-id="e01f1-140">[GUID ジェネレーター](https://guidgenerator.com/)など、多数のオンライン GUID 生成ツールのいずれかを使用してランダムな GUID を生成し、TrustNetworkShareCatalog.reg ファイル内で*両方の場所*の文字列「-random-GUID-here-」を GUID に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-140">Use one of the many online GUID generation tools, such as [GUID Generator](https://guidgenerator.com/), to generate a random GUID, and within the TrustNetworkShareCatalog.reg file, replace the string "-random-GUID-here-" *in both places* with the GUID.</span></span> <span data-ttu-id="e01f1-141">(引用符 `{}` 記号は残しておく必要があります)。</span><span class="sxs-lookup"><span data-stu-id="e01f1-141">(The enclosing `{}` symbols should remain.)</span></span>

4. <span data-ttu-id="e01f1-142">`Url` 値を、以前[共有](#share-a-folder)したフォルダーへの完全なネットワーク パスに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-142">In the `Url` box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="e01f1-143">(URL の `\` 文字は 2 倍にする必要があります。) フォルダーを共有した際に完全なネットワーク パスを書き留めておかなかった場合は、次のスクリーン ショットに示されるように、フォルダーの [**プロパティ**] ダイアログ ウィンドウから取得できます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-143">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![[共有] タブとネットワーク パスが強調表示されているフォルダーの [プロパティ] ダイアログ](../images/sideload-windows-properties-dialog-2.png)
    
5. <span data-ttu-id="e01f1-145">ファイルは、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="e01f1-145">The method should now look like the following.</span></span> <span data-ttu-id="e01f1-146">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-146">Save it.</span></span>

    ```
    Windows Registry Editor Version 5.00
    
    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{01234567-89ab-cedf-0123-456789abcedf}]
    "Id"="{01234567-89ab-cedf-0123-456789abcedf}"
    "Url"="\\\\TestServer\\OfficeAddinManifests"
    "Flags"=dword:00000001
    ```

6. <span data-ttu-id="e01f1-147">*すべて*の Office アプリケーションを閉じます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-147">Close all Office applications, including Excel.</span></span>

7. <span data-ttu-id="e01f1-148">ダブルクリックするなど、実行可能ファイルと同様に TrustNetworkShareCatalog.reg 実行します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-148">Run the TrustNetworkShareCatalog.reg just as you would any executable, such as double-clicking it.</span></span>

## <a name="sideload-your-add-in"></a><span data-ttu-id="e01f1-149">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="e01f1-149">Sideload your add-in</span></span>

1. <span data-ttu-id="e01f1-150">テストするアドインのマニフェスト XML ファイルを共有フォルダー カタログに置きます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-150">Put the manifest XML file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="e01f1-151">なお、Web アプリケーション自体を Web サーバーに展開します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-151">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="e01f1-152">必ずマニフェスト ファイルの **SourceLocation** 要素で URL を指定してください。</span><span class="sxs-lookup"><span data-stu-id="e01f1-152">Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="e01f1-153">Excel、Word、または PowerPoint で、リボンの **[挿入]** タブにある **[個人用アドイン]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-153">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span> <span data-ttu-id="e01f1-154">Projectで、リボンの [**Project**]タブの [**個人用アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-154">In Project, select **My Add-ins** on the **Project** tab of the ribbon.</span></span> 

3. <span data-ttu-id="e01f1-155">**[Office アドイン]** ダイアログ ボックスの上部にある **[共有フォルダー]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="e01f1-155">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="e01f1-156">アドインの名前を選び、**[追加]** を選択して、アドインを挿入します。</span><span class="sxs-lookup"><span data-stu-id="e01f1-156">Select the name of the add-in and choose **Add** to insert the add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="e01f1-157">関連項目</span><span class="sxs-lookup"><span data-stu-id="e01f1-157">See also</span></span>

- [<span data-ttu-id="e01f1-158">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="e01f1-158">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="e01f1-159">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="e01f1-159">Publish your Office Add-in</span></span>](../publish/publish.md)
    
