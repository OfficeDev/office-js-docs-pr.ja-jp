---
title: ネットワーク共有からのテスト用に Office アドインをサイドロードする
description: ネットワーク共有からテスト用に Office アドインをサイドロードする方法について説明します。
ms.date: 06/02/2020
localization_priority: Normal
ms.openlocfilehash: 268fb79c6340aa2d0b8e8278683a0c47b3b60c0e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611247"
---
# <a name="sideload-office-add-ins-for-testing-from-a-network-share"></a><span data-ttu-id="fbf39-103">ネットワーク共有からのテスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="fbf39-103">Sideload Office Add-ins for testing from a network share</span></span>

<span data-ttu-id="fbf39-104">Windows 上の Office クライアントで Office アドインをテストするには、マニフェストをネットワークファイル共有に発行します (以下の手順を参照)。</span><span class="sxs-lookup"><span data-stu-id="fbf39-104">You can test an Office Add-in in an Office client that is on Windows by publishing the manifest to a network file share (instructions below).</span></span> <span data-ttu-id="fbf39-105">この展開オプションは、localhost での開発とテストが完了し、ローカル以外のサーバーまたはクラウドアカウントからアドインをテストする場合に使用することを目的としています。</span><span class="sxs-lookup"><span data-stu-id="fbf39-105">This deployment option is intended to be used when you have completed development and testing on a localhost and want to test the add-in from a non-local server or cloud account.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fbf39-106">ネットワーク共有による展開は、運用環境のアドインではサポートされていません。このメソッドには、次の制限があります。</span><span class="sxs-lookup"><span data-stu-id="fbf39-106">Deployment by network share is not supported for production add-ins. This method has the following limitations:</span></span>
> 
> - <span data-ttu-id="fbf39-107">アドインは、Windows コンピューターにのみインストールできます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-107">The add-in can only be installed on Windows computers.</span></span>
> - <span data-ttu-id="fbf39-108">新しいバージョンのアドインでリボンが変更された場合、各ユーザーはアドインを再インストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="fbf39-108">If a new version of an add-in changes the ribbon, each user will have to reinstall the add-in.</span></span>


> [!NOTE]
> <span data-ttu-id="fbf39-109">アドイン プロジェクトが十分に新しい [Office 用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office) バージョンで作成されている場合、アドインは `npm start` を実行すると自動的に Office デスクトップ クライアントにサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="fbf39-109">If your add-in project was created with a sufficiently recent version of the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), the add-in will automatically sideload in the Office desktop client when you run `npm start`.</span></span>

<span data-ttu-id="fbf39-110">この記事は、Word、Excel、PowerPoint、およびプロジェクトアドインのテストにのみ適用され、Windows 上でのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-110">This article applies only to testing Word, Excel, PowerPoint, and Project add-ins and only on Windows.</span></span> <span data-ttu-id="fbf39-111">異なるプラットフォームでのテストまたは Outlook アドインのテストをする場合は、以下の、アドインのサイドロードに関するいずれかのトピックを参照してください。</span><span class="sxs-lookup"><span data-stu-id="fbf39-111">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="fbf39-112">テスト用に Office on the web で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="fbf39-112">Sideload Office Add-ins in Office on the web for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="fbf39-113">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="fbf39-113">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="fbf39-114">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="fbf39-114">Sideload Outlook add-ins for testing</span></span>](../outlook/sideload-outlook-add-ins-for-testing.md)

<span data-ttu-id="fbf39-115">次のビデオでは、共有フォルダー カタログを使用して Office on the web またはデスクトップでアドインをサイドロードする手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-115">The following video walks you through the process of sideloading your add-in in Office on the web or desktop using a shared folder catalog.</span></span>  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a><span data-ttu-id="fbf39-116">フォルダーの共有</span><span class="sxs-lookup"><span data-stu-id="fbf39-116">Share a folder</span></span>

1. <span data-ttu-id="fbf39-117">アドインをホストさせようとしている Windows コンピューターで、共有フォルダー カタログとして使用するつもりのフォルダーの親フォルダーまたはドライブ文字に移動します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-117">In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="fbf39-118">共有フォルダー カタログとして使用するフォルダーのコンテキスト メニューを開き (フォルダーを右クリック)、[**プロパティ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-118">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="fbf39-119">[**プロパティ**] ダイアログ ボックス内で [**共有**] タブを選択し、[**共有**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-119">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![[共有] タブと [共有] ボタンが強調表示されているフォルダーの [プロパティ] ダイアログ](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="fbf39-121">[**ネットワーク アクセス**] ダイアログ ウィンドウで自分自身とアドインを共有する相手のユーザーまたはグループを追加します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-121">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="fbf39-122">最低でも、フォルダーへの**読み取り/書き込み**アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="fbf39-122">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="fbf39-123">共有する相手の選択が完了したら、[**共有**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-123">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="fbf39-124">「**ユーザーのフォルダーは共有されています**」という確認メッセージが表示されたら、フォルダー名のすぐ後に表示される完全なネットワーク パスを書き留めます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-124">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="fbf39-125">(この記事の次のセクションで説明する通り、[共有フォルダーを信頼できるカタログとして指定する](#specify-the-shared-folder-as-a-trusted-catalog)際に、このネットワーク パスを [**カタログの URL**] として入力する必要があります。) [**完了**] を選択して [**ネットワーク アクセス**] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-125">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![共有パスが強調表示された [ネットワーク アクセス] ダイアログ](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="fbf39-127">[**閉じる**] を選択して、[**プロパティ**] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-127">Choose the **Close** button to close the **Properties** dialog window.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="fbf39-128">共有フォルダーを信頼できるカタログとして指定する</span><span class="sxs-lookup"><span data-stu-id="fbf39-128">Specify the shared folder as a trusted catalog</span></span>

### <a name="configure-the-trust-manually"></a><span data-ttu-id="fbf39-129">信頼を手動で構成する</span><span class="sxs-lookup"><span data-stu-id="fbf39-129">Configure the trust manually</span></span>

1. <span data-ttu-id="fbf39-130">Excel、Word、PowerPoint、または Project で新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-130">Open a new document in Excel, Word, PowerPoint, or Project.</span></span>

2. <span data-ttu-id="fbf39-131">[**ファイル**] タブを選択し、[**オプション**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-131">Choose the **File** tab, and then choose **Options**.</span></span>

3. <span data-ttu-id="fbf39-132">[**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-132">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>

4. <span data-ttu-id="fbf39-133">[**信頼されているアドイン カタログ**] を選びます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-133">Choose **Trusted Add-in Catalogs**.</span></span>

5. <span data-ttu-id="fbf39-134">[**カタログの URL**] ボックスで、先ほど[共有](#share-a-folder)したフォルダーの完全なネットワーク パスを入力します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-134">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="fbf39-135">フォルダーを共有した際に完全なネットワーク パスを書き留めておかなかった場合は、次のスクリーン ショットに示されるように、フォルダーの [**プロパティ**] ダイアログ ウィンドウから取得できます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-135">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span>

    ![[共有] タブとネットワーク パスが強調表示されているフォルダーの [プロパティ] ダイアログ](../images/sideload-windows-properties-dialog-2.png)

6. <span data-ttu-id="fbf39-137">[**カタロ URL**] ボックスにフォルダーの完全なネットワーク パスを入力したら、[**カタログの追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-137">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="fbf39-138">新しく追加されたアイテムの [**メニューに表示する**] チェック ボックスをオンにし、[**OK**] を選択して [**セキュリティ センター** ] ダイアログ ウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-138">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![カタログが選択されている [セキュリティ センター] ダイアログ](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="fbf39-140">**[OK** ] ボタンをクリックして [**オプション**] ダイアログウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-140">Choose the **OK** button to close the **Options** dialog window.</span></span>

9. <span data-ttu-id="fbf39-141">Office アプリケーションを閉じてからもう一度開くと変更内容が有効になります。</span><span class="sxs-lookup"><span data-stu-id="fbf39-141">Close and reopen the Office application so your changes will take effect.</span></span>

### <a name="configure-the-trust-with-a-registry-script"></a><span data-ttu-id="fbf39-142">レジストリ スクリプトを使用して信頼を構成する</span><span class="sxs-lookup"><span data-stu-id="fbf39-142">Configure the trust with a Registry script</span></span>

1. <span data-ttu-id="fbf39-143">テキスト エディターで、TrustNetworkShareCatalog.reg という名前のファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-143">In a text editor, create a file named TrustNetworkShareCatalog.reg.</span></span>

2. <span data-ttu-id="fbf39-144">次に示すコンテンツをファイルに追加します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-144">Add the following content to the file:</span></span>

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-random-GUID-here-}]
    "Id"="{-random-GUID-here-}"
    "Url"="\\\\-share-\\-folder-"
    "Flags"=dword:00000001
    ```
3. <span data-ttu-id="fbf39-145">[GUID ジェネレーター](https://guidgenerator.com/)など、多数のオンライン GUID 生成ツールのいずれかを使用してランダムな GUID を生成し、TrustNetworkShareCatalog.reg ファイル内で*両方の場所*の文字列「-random-GUID-here-」を GUID に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-145">Use one of the many online GUID generation tools, such as [GUID Generator](https://guidgenerator.com/), to generate a random GUID, and within the TrustNetworkShareCatalog.reg file, replace the string "-random-GUID-here-" *in both places* with the GUID.</span></span> <span data-ttu-id="fbf39-146">(引用符 `{}` 記号は残しておく必要があります)。</span><span class="sxs-lookup"><span data-stu-id="fbf39-146">(The enclosing `{}` symbols should remain.)</span></span>

4. <span data-ttu-id="fbf39-147">`Url` 値を、以前[共有](#share-a-folder)したフォルダーへの完全なネットワーク パスに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-147">Replace the `Url` value with the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="fbf39-148">(URL の `\` 文字は 2 倍にする必要があります。) フォルダーを共有した際に完全なネットワーク パスを書き留めておかなかった場合は、次のスクリーン ショットに示されるように、フォルダーの [**プロパティ**] ダイアログ ウィンドウから取得できます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-148">(Note that any `\` characters in the URL must be doubled.) If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span>

    ![[共有] タブとネットワーク パスが強調表示されているフォルダーの [プロパティ] ダイアログ](../images/sideload-windows-properties-dialog-2.png)

5. <span data-ttu-id="fbf39-150">ファイルは、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="fbf39-150">The file should now look like the following.</span></span> <span data-ttu-id="fbf39-151">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-151">Save it.</span></span>

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{01234567-89ab-cedf-0123-456789abcedf}]
    "Id"="{01234567-89ab-cedf-0123-456789abcedf}"
    "Url"="\\\\TestServer\\OfficeAddinManifests"
    "Flags"=dword:00000001
    ```

6. <span data-ttu-id="fbf39-152">*すべて*の Office アプリケーションを閉じます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-152">Close *all* Office applications.</span></span>

7. <span data-ttu-id="fbf39-153">ダブルクリックするなど、実行可能ファイルと同様に TrustNetworkShareCatalog.reg 実行します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-153">Run the TrustNetworkShareCatalog.reg just as you would any executable, such as double-clicking it.</span></span>

## <a name="sideload-your-add-in"></a><span data-ttu-id="fbf39-154">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="fbf39-154">Sideload your add-in</span></span>

1. <span data-ttu-id="fbf39-155">テストするアドインのマニフェスト XML ファイルを共有フォルダー カタログに置きます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-155">Put the manifest XML file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="fbf39-156">なお、Web アプリケーション自体を Web サーバーに展開します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-156">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="fbf39-157">必ずマニフェスト ファイルの **SourceLocation** 要素で URL を指定してください。</span><span class="sxs-lookup"><span data-stu-id="fbf39-157">Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="fbf39-158">Excel、Word、または PowerPoint で、リボンの **[挿入]** タブにある **[個人用アドイン]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-158">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span> <span data-ttu-id="fbf39-159">Projectで、リボンの [**Project**]タブの [**個人用アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-159">In Project, select **My Add-ins** on the **Project** tab of the ribbon.</span></span>

3. <span data-ttu-id="fbf39-160">**[Office アドイン]** ダイアログ ボックスの上部にある **[共有フォルダー]** を選びます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-160">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="fbf39-161">アドインの名前を選び、**[追加]** を選択して、アドインを挿入します。</span><span class="sxs-lookup"><span data-stu-id="fbf39-161">Select the name of the add-in and choose **Add** to insert the add-in.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="fbf39-162">サイドロードアドインを削除する</span><span class="sxs-lookup"><span data-stu-id="fbf39-162">Remove a sideloaded add-in</span></span>

<span data-ttu-id="fbf39-163">コンピューター上の Office キャッシュをクリアすることによって、以前のサイドロードアドインを削除することができます。</span><span class="sxs-lookup"><span data-stu-id="fbf39-163">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="fbf39-164">Windows のキャッシュをクリアする方法については、記事「 [Office キャッシュをクリア](clear-cache.md#clear-the-office-cache-on-windows)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fbf39-164">Details on how to clear the cache on Windows can be found in the article [Clear the Office cache](clear-cache.md#clear-the-office-cache-on-windows).</span></span>

## <a name="see-also"></a><span data-ttu-id="fbf39-165">関連項目</span><span class="sxs-lookup"><span data-stu-id="fbf39-165">See also</span></span>

- [<span data-ttu-id="fbf39-166">Office アドインのマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="fbf39-166">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="fbf39-167">Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="fbf39-167">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="fbf39-168">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="fbf39-168">Publish your Office Add-in</span></span>](../publish/publish.md)
