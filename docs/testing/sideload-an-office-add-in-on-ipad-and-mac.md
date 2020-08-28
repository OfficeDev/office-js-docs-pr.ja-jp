---
title: テスト用に iPad と Mac で Office アドインをサイドロードする
description: サイドロードを使用して iPad および Mac で Office アドインをテストする
ms.date: 02/18/2020
localization_priority: Normal
ms.openlocfilehash: 1a1cb804a72aa182480d06009cf30b41a37276d2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292203"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="a503a-103">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="a503a-103">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="a503a-p101">Office on iOS でアドインの実行状態を確認するには、iTunes を利用してアドインのマニフェストを iPad にサイドロードするか、Office on Mac でアドインのマニフェストを直接サイドロードします。このアクションでは、実行中にブレークポイントを設定したり、アドインのコードをデバッグしたりできませんが、その動作を確認したり、UI が使いやすいかどうかや、適切にレンダリングされているかどうかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="a503a-p101">To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span>

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="a503a-106">Office on iOS の前提条件</span><span class="sxs-lookup"><span data-stu-id="a503a-106">Prerequisites for Office on iOS</span></span>

- <span data-ttu-id="a503a-107">[iTunes](https://www.apple.com/itunes/download/) がインストールされた Windows または Mac コンピューター。</span><span class="sxs-lookup"><span data-stu-id="a503a-107">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>

- <span data-ttu-id="a503a-108">[Excel on iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) がインストールされた iOS 8.2 以降の iPad と同期ケーブル。</span><span class="sxs-lookup"><span data-stu-id="a503a-108">An iPad running iOS 8.2 or later with [Excel on iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.</span></span>

- <span data-ttu-id="a503a-109">テスト対象アドインのマニフェスト .xml ファイル。</span><span class="sxs-lookup"><span data-stu-id="a503a-109">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="a503a-110">Office on Mac の前提条件</span><span class="sxs-lookup"><span data-stu-id="a503a-110">Prerequisites for Office on Mac</span></span>

- <span data-ttu-id="a503a-111">[Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) がインストールされていて OS X v10.10 "Yosemite" を実行している Mac。</span><span class="sxs-lookup"><span data-stu-id="a503a-111">A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>

- <span data-ttu-id="a503a-112">Word on Mac バージョン 15.18 (160109)。</span><span class="sxs-lookup"><span data-stu-id="a503a-112">Word on Mac version 15.18 (160109).</span></span>

- <span data-ttu-id="a503a-113">Excel on Mac バージョン 15.19 (160206)。</span><span class="sxs-lookup"><span data-stu-id="a503a-113">Excel on Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="a503a-114">PowerPoint on Mac バージョン 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="a503a-114">PowerPoint on Mac version 15.24 (160614)</span></span>

- <span data-ttu-id="a503a-115">テスト対象アドインのマニフェスト .xml ファイル。</span><span class="sxs-lookup"><span data-stu-id="a503a-115">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad"></a><span data-ttu-id="a503a-116">Excel または Word on iPad にアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="a503a-116">Sideload an add-in on Excel or Word on iPad</span></span>

1. <span data-ttu-id="a503a-117">同期ケーブルを使用し、iPad をコンピューターに接続します。</span><span class="sxs-lookup"><span data-stu-id="a503a-117">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="a503a-118">初めて iPad をコンピューターに接続している場合は、 **このコンピューターを信頼するかどうか**を確認するメッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="a503a-118">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="a503a-119">**[信頼する]** を選択して続行します。</span><span class="sxs-lookup"><span data-stu-id="a503a-119">Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="a503a-120">iTunes で、メニュー バーの下にある **[iPad]** のアイコンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="a503a-120">In iTunes, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="a503a-121">iTunes の左側の **[設定]** で、**[App]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="a503a-121">Under **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="a503a-122">iTunes の右側で、**[ファイル共有]** までスクロールしてから、**[アドイン]** 列で **[Excel]** または **[Word]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="a503a-122">On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="a503a-123">[ **Excel** ] 列または [ **Word ドキュメント** ] 列の下部で、[ **ファイルの追加**] を選択し、サイドロードするアドインの manifest.xml ファイルを選択します。</span><span class="sxs-lookup"><span data-stu-id="a503a-123">At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span>

6. <span data-ttu-id="a503a-124">iPad で Excel または Word アプリを開きます。</span><span class="sxs-lookup"><span data-stu-id="a503a-124">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="a503a-125">Excel または Word アプリが既に実行されている場合は、[ **ホーム** ] ボタンを選択し、アプリを閉じて再起動します。</span><span class="sxs-lookup"><span data-stu-id="a503a-125">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

7. <span data-ttu-id="a503a-126">ドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="a503a-126">Open a document.</span></span>

8. <span data-ttu-id="a503a-127">[**挿入**] タブで [**アドイン**] を選択します。サイドロードアドインは **、アドインの UI の**[**開発者**] 見出しの下に挿入できます。</span><span class="sxs-lookup"><span data-stu-id="a503a-127">Choose **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Excel アプリでアドインを挿入](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="a503a-129">Office on Mac にアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="a503a-129">Sideload an add-in in Office on Mac</span></span>

> [!NOTE]
> <span data-ttu-id="a503a-130">Mac に Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」をご参照ください。</span><span class="sxs-lookup"><span data-stu-id="a503a-130">To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="a503a-131">**ターミナル**を開き、次のいずれかのフォルダーに移動して、アドインのマニフェストファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="a503a-131">Open **Terminal** and go to one of the following folders where you'll save your add-in's manifest file.</span></span> <span data-ttu-id="a503a-132">`wef` フォルダーがコンピューター上に存在しない場合は、作成します。</span><span class="sxs-lookup"><span data-stu-id="a503a-132">If the `wef` folder doesn't exist on your computer, create it.</span></span>

    - <span data-ttu-id="a503a-133">Word の場合: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="a503a-133">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>    
    - <span data-ttu-id="a503a-134">Excel の場合: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="a503a-134">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="a503a-135">PowerPoint の場合: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="a503a-135">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>

2. <span data-ttu-id="a503a-136">コマンド**Finder** `open .` (ピリオドまたはドットを含む) を使用して、Finder でフォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="a503a-136">Open the folder in **Finder** using the command `open .` (including the period or dot).</span></span> <span data-ttu-id="a503a-137">アドインのマニフェスト ファイルをこのフォルダーにコピーします。</span><span class="sxs-lookup"><span data-stu-id="a503a-137">Copy your add-in's manifest file to this folder.</span></span>

    ![Office on Mac の Wef フォルダー](../images/all-my-files.png)

3. <span data-ttu-id="a503a-p106">Word を起動し、ドキュメントを開きます。既に起動している場合は、Word を再起動します。</span><span class="sxs-lookup"><span data-stu-id="a503a-p106">Open Word, and then open a document. Restart Word if it's already running.</span></span>

4. <span data-ttu-id="a503a-141">Word で、[アドインの**挿入**  >  **Add-ins**  >  **My Add-ins** ] (ドロップダウンメニュー) を選択し、アドインを選択します。</span><span class="sxs-lookup"><span data-stu-id="a503a-141">In Word, choose **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>

    ![Office on Mac の個人用アドイン](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="a503a-p107">サイドロードしたアドインは [個人用アドイン] ダイアログには表示されません。ドロップダウン メニュー内にのみ表示されます (**[挿入]** タブの [個人用アドイン] の右にある小さい下向き矢印)。サイドロードしたアドインは、このメニューの見出し **[開発者向けアドイン]** の下に一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="a503a-p107">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span>

5. <span data-ttu-id="a503a-146">アドインが Word に表示されることを確認します。</span><span class="sxs-lookup"><span data-stu-id="a503a-146">Verify that your add-in is displayed in Word.</span></span>

    ![Office on Mac に表示された Office アドイン](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="a503a-148">サイドロードアドインを削除する</span><span class="sxs-lookup"><span data-stu-id="a503a-148">Remove a sideloaded add-in</span></span>

<span data-ttu-id="a503a-149">コンピューター上の Office キャッシュをクリアすることによって、以前のサイドロードアドインを削除することができます。</span><span class="sxs-lookup"><span data-stu-id="a503a-149">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="a503a-150">各プラットフォームとアプリケーションのキャッシュをクリアする方法については、記事「 [Office キャッシュをクリア](clear-cache.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a503a-150">Details on how to clear the cache for each platform and application can be found in the article [Clear the Office cache](clear-cache.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a503a-151">関連項目</span><span class="sxs-lookup"><span data-stu-id="a503a-151">See also</span></span>

- [<span data-ttu-id="a503a-152">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="a503a-152">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
