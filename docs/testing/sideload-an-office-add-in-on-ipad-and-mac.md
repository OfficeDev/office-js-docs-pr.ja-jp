---
title: テスト用に iPad と Mac で Office アドインをサイドロードする
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e5ec6924917f2351da77c8b9a84eb8de77b3864e
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348129"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="26367-102">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="26367-102">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="26367-p101">Office for iOS でのアドインの実行状況を確認するには、iTunes を 使用してアドインのマニフェストを iPad にサイドロードするか、アドインのマニフェストを Office for Mac で直接サイドロードします。このアクションでは、ブレークポイントを設定したり、アドインのコードを実行中にデバッグすることはできませんが、その動作を確認したり、UI が使用可能であり適切にレンダリングされることを確認できます。</span><span class="sxs-lookup"><span data-stu-id="26367-p101">To see how your add-in will run in Office for iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office for Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span> 

## <a name="prerequisites-for-office-for-ios"></a><span data-ttu-id="26367-105">Office for iOS の前提条件</span><span class="sxs-lookup"><span data-stu-id="26367-105">Prerequisites for Office for iOS</span></span>

- <span data-ttu-id="26367-106">[iTunes](http://www.apple.com/itunes/download/) がインストールされた Windows または Mac コンピュータ。</span><span class="sxs-lookup"><span data-stu-id="26367-106">A Windows or Mac computer with [iTunes](http://www.apple.com/itunes/download/) installed.</span></span>
    
- <span data-ttu-id="26367-107">[Excel for iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) がインストールされた iOS 8.2 以上の iPad と同期ケーブル。</span><span class="sxs-lookup"><span data-stu-id="26367-107">An iPad running iOS 8.2 or later with [Excel for iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.</span></span>
    
- <span data-ttu-id="26367-108">テスト対象アドインのマニフェスト .xml ファイル。</span><span class="sxs-lookup"><span data-stu-id="26367-108">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="prerequisites-for-office-for-mac"></a><span data-ttu-id="26367-109">Office for Mac の前提条件</span><span class="sxs-lookup"><span data-stu-id="26367-109">Prerequisites for Office for Mac</span></span>

- <span data-ttu-id="26367-110">[Office for Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) をインストールし、OS X v10.10 "Yosemite" 以降で動作している Mac。</span><span class="sxs-lookup"><span data-stu-id="26367-110">A Mac running OS X v10.10 "Yosemite" or later with [Office for Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>
    
- <span data-ttu-id="26367-111">Word for Mac バージョン 15.18 (160109)。</span><span class="sxs-lookup"><span data-stu-id="26367-111">Word for Mac version 15.18 (160109).</span></span>
   
- <span data-ttu-id="26367-112">Excel for Mac バージョン 15.19 (160206)。</span><span class="sxs-lookup"><span data-stu-id="26367-112">Excel for Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="26367-113">PowerPoint for Mac バージョン 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="26367-113">PowerPoint for Mac version 15.24 (160614)</span></span>
    
- <span data-ttu-id="26367-114">テスト対象アドインのマニフェスト .xml ファイル。</span><span class="sxs-lookup"><span data-stu-id="26367-114">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="sideload-an-add-in-on-excel-or-word-for-ipad"></a><span data-ttu-id="26367-115">Excel for iPad または Word for iPad のアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="26367-115">Sideload an add-in on Excel or Word for iPad</span></span>

1. <span data-ttu-id="26367-p102">同期ケーブルを使用し、iPad をコンピューターに接続します。iPad を初めてコンピューターに接続する場合、**[このコンピューターを信頼しますか?]** というメッセージが表示されるので、**[信頼する]** を選択して続行します。</span><span class="sxs-lookup"><span data-stu-id="26367-p102">Use a sync cable to connect your iPad to your computer. If you're connecting the iPad to your computer for the first time, you'll be prompted with  **Trust This Computer?**. Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="26367-119">iTunes のメニュー バーの下にある **[iPad]** アイコンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="26367-119">In iTunes, choose the  **iPad** icon below the menu bar.</span></span>
    
    ![iTunes の iPad アイコン](../images/ipad.png)

3. <span data-ttu-id="26367-121">iTunes の左側にある  **[設定]** で、 **[App]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="26367-121">Under  **Settings** on the left side of iTunes, choose **Apps**.</span></span>
    
    ![iTunes App 設定](../images/file-settings-apps.png)

4. <span data-ttu-id="26367-123">iTunes の右側で、 **[ファイル共有]** までスクロールしてから、 **[アドイン]** 列で **[Excel]** または **[Word]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="26367-123">On the right side of iTunes, scroll down to  **File Sharing**, and then choose  **Excel** or **Word** in the **Add-ins** column.</span></span>
    
    ![iTunes ファイル共有](../images/file-sharing.png)

5. <span data-ttu-id="26367-125">**[Excel]** 列または **[Word ドキュメント]** 列の下部で、 **[ファイルの追加]** をクリックしてから、サイドロードするアドインのマニフェスト .xml ファイルを選択します。</span><span class="sxs-lookup"><span data-stu-id="26367-125">At the bottom of the  **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span> 
    
6. <span data-ttu-id="26367-p103">iPad で Excel または Word アプリを開きます。Excel または Word アプリがすでに実行されている場合は、 **[ホーム]** ボタンを選択し、アプリを閉じてから再起動します。</span><span class="sxs-lookup"><span data-stu-id="26367-p103">Open the Excel or Word app on your iPad. If the Excel or Word app is already running, choose the  **Home** button, and then close and restart the app.</span></span>
    
7. <span data-ttu-id="26367-128">ドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="26367-128">Open a document.</span></span>
    
8. <span data-ttu-id="26367-129">**[挿入]** タブで **[アドイン]** をクリックします。 **[アドイン]** UI の **[開発者]** という見出しの下に、サイドロードしたアドインが表示され、挿入するために選択できます。</span><span class="sxs-lookup"><span data-stu-id="26367-129">Choose  **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>
    
    ![Excel アプリでアドインを挿入](../images/excel-insert-add-in.png)


## <a name="sideload-an-add-in-on-office-for-mac"></a><span data-ttu-id="26367-131">Office for Mac でアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="26367-131">Sideload an add-in on Office for Mac</span></span>

> [!NOTE]
> <span data-ttu-id="26367-132">Outlook for Mac アドインをサイドロードするには、「[テスト用に Outlook アドインをサイドロードする](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)」をご参照ください。</span><span class="sxs-lookup"><span data-stu-id="26367-132">To sideload Outlook 2016 for Mac add-in, see [Sideload Outlook add-ins for testing](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

1. <span data-ttu-id="26367-p104">**ターミナル**を開き、アプリに対応するフォルダーに移動しアドインのマニフェスト ファイルを保存します。`wef` フォルダーがコンピューターにない場合には作成します。</span><span class="sxs-lookup"><span data-stu-id="26367-p104">Open  **Terminal** and go to one of the following folders where you'll save your add-in's manifest file. If the `wef` folder doesn't exist on your computer, create it.</span></span>
    
    - <span data-ttu-id="26367-135">Word の場合:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="26367-135">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`</span></span>    
    - <span data-ttu-id="26367-136">Excel の場合:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="26367-136">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`</span></span>
    - <span data-ttu-id="26367-137">PowerPoint の場合: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="26367-137">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`</span></span>
    
2. <span data-ttu-id="26367-p105">**Finder** で `open .` コマンド (ピリオドまたはドットを含む) を使用してフォルダーを開き、アドインのマニフェスト ファイルをこのフォルダーにコピーします。</span><span class="sxs-lookup"><span data-stu-id="26367-p105">Open the folder in  **Finder** using the command `open .` (including the period or dot). Copy your add-in's manifest file to this folder.</span></span>
    
    ![Office for Mac の Wef フォルダー](../images/all-my-files.png)

3. <span data-ttu-id="26367-p106">Word を起動し、ドキュメントを開きます。既に起動している場合は、Word を再起動します。</span><span class="sxs-lookup"><span data-stu-id="26367-p106">Open Word, and then open a document. Restart Word if it's already running.</span></span>
    
4. <span data-ttu-id="26367-143">Word で、**[挿入]** > **[アドイン]** > **[個人用アドイン]** (ドロップダウン メニュー) を選択し、アドインを選択します。</span><span class="sxs-lookup"><span data-stu-id="26367-143">In Word, choose  **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>
    
    ![Office for Mac のマイ アドイン](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="26367-p107">サイドロードしたアドインは [個人用アドイン] ダイアログには表示されず、ドロップダウン メニュー (**[挿入]** タブの [個人用アドイン] の右にある小さな下向き矢印) 内にのみ表示されます。サイドロードしたアドインは、このメニューの見出し **[開発者向けアドイン]** の下に一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="26367-p107">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span> 
    
5. <span data-ttu-id="26367-148">Word にアドインが表示されることを確認します。</span><span class="sxs-lookup"><span data-stu-id="26367-148">Verify that your add-in is displayed in Word.</span></span>
    
    ![Office for Mac で示される Office アドイン](../images/lorem-ipsum-wikipedia.png)
    
    > [!NOTE]
    > <span data-ttu-id="26367-p108">Office for Mac は、パフォーマンス上の理由からアドインを頻繁にキャッシュします。アドイン開発中に強制的にリロードする必要がある場合は、`Users/<usr>/Library/Containers/com.Microsoft.OsfWebHost/Data/` フォルダーを消去します。このフォルダーが存在しない場合は、`com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/` フォルダーのファイルを消去します。</span><span class="sxs-lookup"><span data-stu-id="26367-p108">Add-ins are cached often in Office for Mac, for performance reasons. If you need to force a reload of your add-in while you're developing it, you can clear the `Users/<usr>/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span>

## <a name="see-also"></a><span data-ttu-id="26367-153">関連項目</span><span class="sxs-lookup"><span data-stu-id="26367-153">See also</span></span>

- [<span data-ttu-id="26367-154">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="26367-154">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
    
