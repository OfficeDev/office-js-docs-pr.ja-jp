---
title: Office アドインの作業ウィンドウ
description: 作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: d235d6c437ee124441389e68b54fc6ab8cde8dae
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330151"
---
# <a name="task-panes-in-office-add-ins"></a><span data-ttu-id="1d41c-103">Office アドインの作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1d41c-103">Task panes in Office Add-ins</span></span>

<span data-ttu-id="1d41c-p101">作業ウィンドウは、通常 Word、PowerPoint、Excel、Outlook 内のウィンドウの右側に表示されるインターフェイスのサーフェスです。作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。機能を直接ドキュメントに埋め込む必要がない場合は、作業ウィンドウを使用します。</span><span class="sxs-lookup"><span data-stu-id="1d41c-p101">Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.</span></span>

<span data-ttu-id="1d41c-107">*図 1. 一般的な作業ウィンドウのレイアウト*</span><span class="sxs-lookup"><span data-stu-id="1d41c-107">*Figure 1. Typical task pane layout*</span></span>

![上部にセクション タブ、左下に会社のロゴと会社名、右下に設定アイコンを含む一般的な作業ウィンドウ レイアウトを表示する図](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a><span data-ttu-id="1d41c-109">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="1d41c-109">Best practices</span></span>

|<span data-ttu-id="1d41c-110">するべきこと</span><span class="sxs-lookup"><span data-stu-id="1d41c-110">Do</span></span>|<span data-ttu-id="1d41c-111">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="1d41c-111">Don't</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="1d41c-112">タイトルにアドインの名前を含めます。</span><span class="sxs-lookup"><span data-stu-id="1d41c-112">Include the name of your add-in in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="1d41c-113">タイトルには会社名を追加しません。</span><span class="sxs-lookup"><span data-stu-id="1d41c-113">Don't append your company name to the title.</span></span></li></ul>|
|<ul><li><span data-ttu-id="1d41c-114">タイトルには短くわかりやすい名前を使用します。</span><span class="sxs-lookup"><span data-stu-id="1d41c-114">Use short descriptive names in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="1d41c-115">アドインのタイトルには、"アドイン"、"for Word"、"for Office" などの文字列を追加しません。</span><span class="sxs-lookup"><span data-stu-id="1d41c-115">Don't append strings such as "add-in," "for Word," or "for Office" to the title of your add-in.</span></span></li></ul>|
|<ul><li><span data-ttu-id="1d41c-116">アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。</span><span class="sxs-lookup"><span data-stu-id="1d41c-116">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span></li></ul>||
|<ul><li><span data-ttu-id="1d41c-117">アドインを Outlook 内で使用する場合を除き、アドインの下部に BrandBar などのブランド化の要素を含めます。</span><span class="sxs-lookup"><span data-stu-id="1d41c-117">Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.</span></span></li></ul>||

## <a name="variants"></a><span data-ttu-id="1d41c-118">バリアント</span><span class="sxs-lookup"><span data-stu-id="1d41c-118">Variants</span></span>

<span data-ttu-id="1d41c-119">次の図は、1366x768 解像度のリボンOffice アプリ作業ウィンドウのサイズを示しています。</span><span class="sxs-lookup"><span data-stu-id="1d41c-119">The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution.</span></span> <span data-ttu-id="1d41c-120">Excel では、数式バーを収容するための縦のスペースが必要です。</span><span class="sxs-lookup"><span data-stu-id="1d41c-120">For Excel, additional vertical space is required to accommodate the formula bar.</span></span>  

<span data-ttu-id="1d41c-121">*図 2. Office 2016 デスクトップ作業ウィンドウのサイズ*</span><span class="sxs-lookup"><span data-stu-id="1d41c-121">*Figure 2. Office 2016 desktop task pane sizes*</span></span>

![デスクトップ作業ウィンドウのサイズを 1366x768 解像度で表示する図](../images/office-2016-taskpane-sizes.png)

- <span data-ttu-id="1d41c-123">Excel - 320x455 ピクセル</span><span class="sxs-lookup"><span data-stu-id="1d41c-123">Excel - 320x455 pixels</span></span>
- <span data-ttu-id="1d41c-124">PowerPoint - 320x531 ピクセル</span><span class="sxs-lookup"><span data-stu-id="1d41c-124">PowerPoint - 320x531 pixels</span></span>
- <span data-ttu-id="1d41c-125">Word - 320x531 ピクセル</span><span class="sxs-lookup"><span data-stu-id="1d41c-125">Word - 320x531 pixels</span></span>
- <span data-ttu-id="1d41c-126">Outlook - 348x535 ピクセル</span><span class="sxs-lookup"><span data-stu-id="1d41c-126">Outlook - 348x535 pixels</span></span>

<br/>

<span data-ttu-id="1d41c-127">*図 3.Office作業ウィンドウのサイズ*</span><span class="sxs-lookup"><span data-stu-id="1d41c-127">*Figure 3. Office task pane sizes*</span></span>

![作業ウィンドウのサイズを 1366x768 解像度で表示する図](../images/office-365-taskpane-sizes.png)

- <span data-ttu-id="1d41c-129">Excel - 350x378 ピクセル</span><span class="sxs-lookup"><span data-stu-id="1d41c-129">Excel - 350x378 pixels</span></span>
- <span data-ttu-id="1d41c-130">PowerPoint - 348x391 ピクセル</span><span class="sxs-lookup"><span data-stu-id="1d41c-130">PowerPoint - 348x391 pixels</span></span>
- <span data-ttu-id="1d41c-131">Word - 329x445 ピクセル</span><span class="sxs-lookup"><span data-stu-id="1d41c-131">Word - 329x445 pixels</span></span>
- <span data-ttu-id="1d41c-132">Outlook (web 上) - 320x570 ピクセル</span><span class="sxs-lookup"><span data-stu-id="1d41c-132">Outlook (on the web) - 320x570 pixels</span></span>

## <a name="personality-menu"></a><span data-ttu-id="1d41c-133">パーソナル メニュー</span><span class="sxs-lookup"><span data-stu-id="1d41c-133">Personality menu</span></span>

<span data-ttu-id="1d41c-p103">パーソナル メニューは、アドインの右上付近にあるナビゲーション要素やコマンド要素の妨げになる可能性があります。Windows と Mac でのパーソナル メニューの現在のサイズを次に示します。</span><span class="sxs-lookup"><span data-stu-id="1d41c-p103">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="1d41c-136">Windows の場合、パーソナル メニューは 12x32 ピクセルを測定します (図を参照)。</span><span class="sxs-lookup"><span data-stu-id="1d41c-136">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="1d41c-137">*図 4. Windows のパーソナル メニュー*</span><span class="sxs-lookup"><span data-stu-id="1d41c-137">*Figure 4. Personality menu on Windows*</span></span>

![デスクトップ上のパーソナリティ メニューをWindows図](../images/personality-menu-win.png)

<span data-ttu-id="1d41c-139">Mac の場合、パーソナル メニューは 26x26 ピクセルを測定しますが、右から 8 ピクセル内側、上から 6 ピクセルの位置にフロートします。これにより、スペースは 34x32 ピクセルに増加します (図を参照)。</span><span class="sxs-lookup"><span data-stu-id="1d41c-139">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="1d41c-140">*図 5. Mac のパーソナル メニュー*</span><span class="sxs-lookup"><span data-stu-id="1d41c-140">*Figure 5. Personality menu on Mac*</span></span>

![Mac デスクトップのパーソナリティ メニューを示す図](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="1d41c-142">実装</span><span class="sxs-lookup"><span data-stu-id="1d41c-142">Implementation</span></span>

<span data-ttu-id="1d41c-143">作業ウィンドウを実装するサンプルについては、GitHub の「[Excel アドインの JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1d41c-143">For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="1d41c-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="1d41c-144">See also</span></span>

- [<span data-ttu-id="1d41c-145">ファブリック コア (Office アドイン)</span><span class="sxs-lookup"><span data-stu-id="1d41c-145">Fabric Core in Office Add-ins</span></span>](fabric-core.md)
- [<span data-ttu-id="1d41c-146">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="1d41c-146">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
