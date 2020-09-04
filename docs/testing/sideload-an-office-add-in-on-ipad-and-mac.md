---
title: テスト用に iPad と Mac で Office アドインをサイドロードする
description: サイドロードを使用して、iPad と Mac で Office アドインをテストします。
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 7c5e9542c6e6f9abc96defde389b9543421b8529
ms.sourcegitcommit: 604361e55dee45c7a5d34c2fa6937693c154fc24
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/03/2020
ms.locfileid: "47364058"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="aa7ae-103">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="aa7ae-103">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="aa7ae-p101">Office on iOS でアドインの実行状態を確認するには、iTunes を利用してアドインのマニフェストを iPad にサイドロードするか、Office on Mac でアドインのマニフェストを直接サイドロードします。このアクションでは、実行中にブレークポイントを設定したり、アドインのコードをデバッグしたりできませんが、その動作を確認したり、UI が使いやすいかどうかや、適切にレンダリングされているかどうかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-p101">To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span>

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="aa7ae-106">Office on iOS の前提条件</span><span class="sxs-lookup"><span data-stu-id="aa7ae-106">Prerequisites for Office on iOS</span></span>

- <span data-ttu-id="aa7ae-107">[iTunes](https://www.apple.com/itunes/download/) がインストールされた Windows または Mac コンピューター。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-107">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>
  > [!IMPORTANT]
  > <span data-ttu-id="aa7ae-108">MacOS Catalina を実行している場合 [は、iTunes を使用でき](https://support.apple.com/HT210200) なくなりました。この記事の後半の「 [macos Catalina を使用して、Excel または Word でのアドインのサイドロード](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) 」の手順に従ってください。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-108">If you're running macOS Catalina, [iTunes is no longer available](https://support.apple.com/HT210200) so you should follow the instructions in the section [Sideload an add-in on Excel or Word on iPad using macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) later in this article.</span></span>

- <span data-ttu-id="aa7ae-109">[Excel](https://apps.apple.com/app/microsoft-excel/id586683407)または[Word](https://apps.apple.com/app/microsoft-word/id586447913)がインストールされた iOS 8.2 以降を実行している iPad と、同期ケーブル。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-109">An iPad running iOS 8.2 or later with [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) or [Word](https://apps.apple.com/app/microsoft-word/id586447913) installed, and a sync cable.</span></span>

- <span data-ttu-id="aa7ae-110">テスト対象アドインのマニフェスト .xml ファイル。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-110">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="aa7ae-111">Office on Mac の前提条件</span><span class="sxs-lookup"><span data-stu-id="aa7ae-111">Prerequisites for Office on Mac</span></span>

- <span data-ttu-id="aa7ae-112">[Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) がインストールされていて OS X v10.10 "Yosemite" を実行している Mac。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-112">A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>

- <span data-ttu-id="aa7ae-113">Word on Mac バージョン 15.18 (160109)。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-113">Word on Mac version 15.18 (160109).</span></span>

- <span data-ttu-id="aa7ae-114">Excel on Mac バージョン 15.19 (160206)。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-114">Excel on Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="aa7ae-115">PowerPoint on Mac バージョン 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="aa7ae-115">PowerPoint on Mac version 15.24 (160614)</span></span>

- <span data-ttu-id="aa7ae-116">テスト対象アドインのマニフェスト .xml ファイル。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-116">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a><span data-ttu-id="aa7ae-117">サイドロードを iTunes を使用して Excel または Word で iPad に追加する</span><span class="sxs-lookup"><span data-stu-id="aa7ae-117">Sideload an add-in on Excel or Word on iPad using iTunes</span></span>

1. <span data-ttu-id="aa7ae-118">同期ケーブルを使用し、iPad をコンピューターに接続します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-118">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="aa7ae-119">初めて iPad をコンピューターに接続している場合は、 **このコンピューターを信頼するかどうか**を確認するメッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-119">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="aa7ae-120">**[信頼する]** を選択して続行します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-120">Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="aa7ae-121">iTunes で、メニュー バーの下にある **[iPad]** のアイコンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-121">In iTunes, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="aa7ae-122">iTunes の左側の **[設定]** で、**[App]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-122">Under **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="aa7ae-123">iTunes の右側で、**[ファイル共有]** までスクロールしてから、**[アドイン]** 列で **[Excel]** または **[Word]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-123">On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="aa7ae-124">[ **Excel** ] 列または [ **Word ドキュメント** ] 列の下部で、[ **ファイルの追加**] を選択し、サイドロードするアドインの manifest.xml ファイルを選択します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-124">At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span>

6. <span data-ttu-id="aa7ae-125">iPad で Excel または Word アプリを開きます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-125">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="aa7ae-126">Excel または Word アプリが既に実行されている場合は、[ **ホーム** ] ボタンを選択し、アプリを閉じて再起動します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-126">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

7. <span data-ttu-id="aa7ae-127">ドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-127">Open a document.</span></span>

8. <span data-ttu-id="aa7ae-128">[**挿入**] タブで [**アドイン**] を選択します。 ([**挿入**] タブで、[**アドイン**] ボタンが表示されるまで、横にスクロールする必要がある場合があります)。サイドロードアドインは **、アドインの UI の**[**開発者**] 見出しの下に挿入できます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-128">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Excel アプリでアドインを挿入](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a><span data-ttu-id="aa7ae-130">MacOS Catalina を使用して、Excel または Word でアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-130">Sideload an add-in on Excel or Word on iPad using macOS Catalina</span></span>

> [!IMPORTANT]
> <span data-ttu-id="aa7ae-131">MacOS Catalina の導入により、 [Apple で廃止](https://support.apple.com/HT210200) された ITunes を Mac に、サイドロードアプリを **Finder**にするために必要な統合機能を使用しています。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-131">With the introduction of macOS Catalina, [Apple discontinued iTunes on Mac](https://support.apple.com/HT210200) and integrated functionality required to sideload apps into **Finder**.</span></span>

1. <span data-ttu-id="aa7ae-132">同期ケーブルを使用し、iPad をコンピューターに接続します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-132">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="aa7ae-133">初めて iPad をコンピューターに接続している場合は、 **このコンピューターを信頼するかどうか**を確認するメッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-133">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="aa7ae-134">**[信頼する]** を選択して続行します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-134">Choose **Trust** to continue.</span></span> <span data-ttu-id="aa7ae-135">また、これが新しい iPad であるかどうか、または1つを復元しているかどうかを尋ねられる場合もあります。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-135">You may also be asked if this is a new iPad or if you're restoring one.</span></span>

2. <span data-ttu-id="aa7ae-136">[Finder] の [ **場所**] で、メニューバーの下にある [ **iPad** ] アイコンを選択します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-136">In Finder, under **Locations**, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="aa7ae-137">ファインダーウィンドウの上部で、[ **ファイル**] をクリックし、[ **Excel** ] または [ **Word**] を見つけます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-137">On the top of the Finder window, click on **Files**, and then locate **Excel** or **Word**.</span></span>

4. <span data-ttu-id="aa7ae-138">別のファインダーウィンドウから、最初のファインダーウィンドウで、サイドロードするアドインの manifest.xml ファイルを **Excel** または **Word** ファイルにドラッグアンドドロップします。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-138">From a different Finder window, drag and drop the manifest.xml file of the add-in you want to side load onto the **Excel** or **Word** file in the first Finder window.</span></span>

5. <span data-ttu-id="aa7ae-139">iPad で Excel または Word アプリを開きます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-139">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="aa7ae-140">Excel または Word アプリが既に実行されている場合は、[ **ホーム** ] ボタンを選択し、アプリを閉じて再起動します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-140">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

6. <span data-ttu-id="aa7ae-141">ドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-141">Open a document.</span></span>

7. <span data-ttu-id="aa7ae-142">[**挿入**] タブで [**アドイン**] を選択します。 ([**挿入**] タブで、[**アドイン**] ボタンが表示されるまで、横にスクロールする必要がある場合があります)。サイドロードアドインは **、アドインの UI の**[**開発者**] 見出しの下に挿入できます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-142">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Excel アプリでアドインを挿入](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="aa7ae-144">Office on Mac にアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="aa7ae-144">Sideload an add-in in Office on Mac</span></span>

> [!NOTE]
> <span data-ttu-id="aa7ae-145">Mac に Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-an-add-in-in-outlook-on-the-desktop)」をご参照ください。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-145">To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-an-add-in-in-outlook-on-the-desktop).</span></span>

1. <span data-ttu-id="aa7ae-146">**ターミナル**を開き、次のいずれかのフォルダーに移動して、アドインのマニフェストファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-146">Open **Terminal** and go to one of the following folders where you'll save your add-in's manifest file.</span></span> <span data-ttu-id="aa7ae-147">`wef` フォルダーがコンピューター上に存在しない場合は、作成します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-147">If the `wef` folder doesn't exist on your computer, create it.</span></span>

    - <span data-ttu-id="aa7ae-148">Word の場合: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="aa7ae-148">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>
    - <span data-ttu-id="aa7ae-149">Excel の場合: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="aa7ae-149">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="aa7ae-150">PowerPoint の場合: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="aa7ae-150">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>

2. <span data-ttu-id="aa7ae-151">コマンド**Finder** `open .` (ピリオドまたはドットを含む) を使用して、Finder でフォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-151">Open the folder in **Finder** using the command `open .` (including the period or dot).</span></span> <span data-ttu-id="aa7ae-152">アドインのマニフェスト ファイルをこのフォルダーにコピーします。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-152">Copy your add-in's manifest file to this folder.</span></span>

    ![Office on Mac の Wef フォルダー](../images/all-my-files.png)

3. <span data-ttu-id="aa7ae-p108">Word を起動し、ドキュメントを開きます。既に起動している場合は、Word を再起動します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-p108">Open Word, and then open a document. Restart Word if it's already running.</span></span>

4. <span data-ttu-id="aa7ae-156">Word で、[アドインの**挿入**  >  **Add-ins**  >  **My Add-ins** ] (ドロップダウンメニュー) を選択し、アドインを選択します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-156">In Word, choose **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>

    ![Office on Mac の個人用アドイン](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="aa7ae-p109">サイドロードしたアドインは [個人用アドイン] ダイアログには表示されません。ドロップダウン メニュー内にのみ表示されます (**[挿入]** タブの [個人用アドイン] の右にある小さい下向き矢印)。サイドロードしたアドインは、このメニューの見出し **[開発者向けアドイン]** の下に一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-p109">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span>

5. <span data-ttu-id="aa7ae-161">アドインが Word に表示されることを確認します。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-161">Verify that your add-in is displayed in Word.</span></span>

    ![Office on Mac に表示された Office アドイン](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="aa7ae-163">サイドロードアドインを削除する</span><span class="sxs-lookup"><span data-stu-id="aa7ae-163">Remove a sideloaded add-in</span></span>

<span data-ttu-id="aa7ae-164">コンピューター上の Office キャッシュをクリアすることによって、以前のサイドロードアドインを削除することができます。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-164">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="aa7ae-165">各プラットフォームとアプリケーションのキャッシュをクリアする方法については、記事「 [Office キャッシュをクリア](clear-cache.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa7ae-165">Details on how to clear the cache for each platform and application can be found in the article [Clear the Office cache](clear-cache.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="aa7ae-166">関連項目</span><span class="sxs-lookup"><span data-stu-id="aa7ae-166">See also</span></span>

- [<span data-ttu-id="aa7ae-167">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="aa7ae-167">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
