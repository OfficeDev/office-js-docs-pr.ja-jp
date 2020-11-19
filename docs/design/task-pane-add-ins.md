---
title: Office アドインの作業ウィンドウ
description: 作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: ed3f3b8fdf7cf62b6016fe8b03393de0d56dfb33
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132019"
---
# <a name="task-panes-in-office-add-ins"></a><span data-ttu-id="c0161-103">Office アドインの作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c0161-103">Task panes in Office Add-ins</span></span>

<span data-ttu-id="c0161-p101">作業ウィンドウは、通常 Word、PowerPoint、Excel、Outlook 内のウィンドウの右側に表示されるインターフェイスのサーフェスです。作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。機能を直接ドキュメントに埋め込む必要がない場合は、作業ウィンドウを使用します。</span><span class="sxs-lookup"><span data-stu-id="c0161-p101">Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.</span></span>

<span data-ttu-id="c0161-107">*図 1. 一般的な作業ウィンドウのレイアウト*</span><span class="sxs-lookup"><span data-stu-id="c0161-107">*Figure 1. Typical task pane layout*</span></span>

![上部にセクションタブ、左側には会社のロゴと会社名、右下に設定アイコンが表示される一般的な作業ウィンドウレイアウトの図](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a><span data-ttu-id="c0161-109">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="c0161-109">Best practices</span></span>

|<span data-ttu-id="c0161-110">するべきこと</span><span class="sxs-lookup"><span data-stu-id="c0161-110">Do</span></span>|<span data-ttu-id="c0161-111">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="c0161-111">Don't</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="c0161-112">タイトルにアドインの名前を含めます。</span><span class="sxs-lookup"><span data-stu-id="c0161-112">Include the name of your add-in in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="c0161-113">タイトルには会社名を追加しません。</span><span class="sxs-lookup"><span data-stu-id="c0161-113">Don't append your company name to the title.</span></span></li></ul>|
|<ul><li><span data-ttu-id="c0161-114">タイトルには短くわかりやすい名前を使用します。</span><span class="sxs-lookup"><span data-stu-id="c0161-114">Use short descriptive names in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="c0161-115">アドインのタイトルには、「アドイン」、「Word」、「Office」などの文字列を追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="c0161-115">Don't append strings such as "add-in," "for Word," or "for Office" to the title of your add-in.</span></span></li></ul>|
|<ul><li><span data-ttu-id="c0161-116">アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。</span><span class="sxs-lookup"><span data-stu-id="c0161-116">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span></li></ul>||
|<ul><li><span data-ttu-id="c0161-117">アドインを Outlook 内で使用する場合を除き、アドインの下部に BrandBar などのブランド化の要素を含めます。</span><span class="sxs-lookup"><span data-stu-id="c0161-117">Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.</span></span></li></ul>||

## <a name="variants"></a><span data-ttu-id="c0161-118">バリアント</span><span class="sxs-lookup"><span data-stu-id="c0161-118">Variants</span></span>

<span data-ttu-id="c0161-p102">次の画像は、1366x768 の解像度で Office アプリのリボンを使用して、さまざまな作業ウィンドウのサイズを示しています。Excel では、数式バーに対応するために、さらに広い領域が必要になります。</span><span class="sxs-lookup"><span data-stu-id="c0161-p102">The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.</span></span>  

<span data-ttu-id="c0161-121">*図 2. Office 2016 デスクトップ作業ウィンドウのサイズ*</span><span class="sxs-lookup"><span data-stu-id="c0161-121">*Figure 2. Office 2016 desktop task pane sizes*</span></span>

![1366x768 の解像度でデスクトップ作業ウィンドウのサイズを表示する図](../images/office-2016-taskpane-sizes.png)

- <span data-ttu-id="c0161-123">Excel-320x455 ピクセル</span><span class="sxs-lookup"><span data-stu-id="c0161-123">Excel - 320x455 pixels</span></span>
- <span data-ttu-id="c0161-124">PowerPoint-320x531 ピクセル</span><span class="sxs-lookup"><span data-stu-id="c0161-124">PowerPoint - 320x531 pixels</span></span>
- <span data-ttu-id="c0161-125">Word-320x531 ピクセル</span><span class="sxs-lookup"><span data-stu-id="c0161-125">Word - 320x531 pixels</span></span>
- <span data-ttu-id="c0161-126">Outlook-348x535 ピクセル</span><span class="sxs-lookup"><span data-stu-id="c0161-126">Outlook - 348x535 pixels</span></span>

<br/>

<span data-ttu-id="c0161-127">*図3Office の作業ウィンドウのサイズ*</span><span class="sxs-lookup"><span data-stu-id="c0161-127">*Figure 3. Office task pane sizes*</span></span>

![1366x768 の解像度で作業ウィンドウのサイズを表示する図](../images/office-365-taskpane-sizes.png)

- <span data-ttu-id="c0161-129">Excel-350x378 ピクセル</span><span class="sxs-lookup"><span data-stu-id="c0161-129">Excel - 350x378 pixels</span></span>
- <span data-ttu-id="c0161-130">PowerPoint-348x391 ピクセル</span><span class="sxs-lookup"><span data-stu-id="c0161-130">PowerPoint - 348x391 pixels</span></span>
- <span data-ttu-id="c0161-131">ワード329x445 ピクセル</span><span class="sxs-lookup"><span data-stu-id="c0161-131">Word - 329x445 pixels</span></span>
- <span data-ttu-id="c0161-132">Outlook (web 上)-320x570 ピクセル</span><span class="sxs-lookup"><span data-stu-id="c0161-132">Outlook (on the web) - 320x570 pixels</span></span>

## <a name="personality-menu"></a><span data-ttu-id="c0161-133">パーソナル メニュー</span><span class="sxs-lookup"><span data-stu-id="c0161-133">Personality menu</span></span>

<span data-ttu-id="c0161-p103">パーソナル メニューは、アドインの右上付近にあるナビゲーション要素やコマンド要素の妨げになる可能性があります。Windows と Mac でのパーソナル メニューの現在のサイズを次に示します。</span><span class="sxs-lookup"><span data-stu-id="c0161-p103">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="c0161-136">Windows の場合、パーソナル メニューは 12x32 ピクセルを測定します (図を参照)。</span><span class="sxs-lookup"><span data-stu-id="c0161-136">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="c0161-137">*図 4. Windows のパーソナル メニュー*</span><span class="sxs-lookup"><span data-stu-id="c0161-137">*Figure 4. Personality menu on Windows*</span></span>

![Windows デスクトップのパーソナルメニューを示す図](../images/personality-menu-win.png)

<span data-ttu-id="c0161-139">Mac の場合、パーソナル メニューは 26x26 ピクセルを測定しますが、右から 8 ピクセル内側、上から 6 ピクセルの位置にフロートします。これにより、スペースは 34x32 ピクセルに増加します (図を参照)。</span><span class="sxs-lookup"><span data-stu-id="c0161-139">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="c0161-140">*図 5. Mac のパーソナル メニュー*</span><span class="sxs-lookup"><span data-stu-id="c0161-140">*Figure 5. Personality menu on Mac*</span></span>

![Mac デスクトップのパーソナルメニューを示す図](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="c0161-142">実装</span><span class="sxs-lookup"><span data-stu-id="c0161-142">Implementation</span></span>

<span data-ttu-id="c0161-143">作業ウィンドウを実装するサンプルについては、GitHub の「[Excel アドインの JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c0161-143">For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="c0161-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="c0161-144">See also</span></span>

- [<span data-ttu-id="c0161-145">Office アドインでの Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="c0161-145">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
- [<span data-ttu-id="c0161-146">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="c0161-146">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
