---
title: Office アドインの作業ウィンドウ
description: 作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 429042db7e30f5fefe48c9648e6ad5410f6594c4
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608496"
---
# <a name="task-panes-in-office-add-ins"></a><span data-ttu-id="761fc-103">Office アドインの作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="761fc-103">Task panes in Office Add-ins</span></span>
 
<span data-ttu-id="761fc-p101">作業ウィンドウは、通常 Word、PowerPoint、Excel、Outlook 内のウィンドウの右側に表示されるインターフェイスのサーフェスです。作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。機能を直接ドキュメントに埋め込む必要がない場合は、作業ウィンドウを使用します。</span><span class="sxs-lookup"><span data-stu-id="761fc-p101">Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.</span></span>

<span data-ttu-id="761fc-107">*図 1. 一般的な作業ウィンドウのレイアウト*</span><span class="sxs-lookup"><span data-stu-id="761fc-107">*Figure 1. Typical task pane layout*</span></span>

![一般的な作業ウィンドウのレイアウトを表示するイメージ](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a><span data-ttu-id="761fc-109">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="761fc-109">Best practices</span></span>

|<span data-ttu-id="761fc-110">**するべきこと**</span><span class="sxs-lookup"><span data-stu-id="761fc-110">**Do**</span></span>|<span data-ttu-id="761fc-111">**してはいけないこと**</span><span class="sxs-lookup"><span data-stu-id="761fc-111">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="761fc-112">タイトルにアドインの名前を含めます。</span><span class="sxs-lookup"><span data-stu-id="761fc-112">Include the name of your add-in in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="761fc-113">タイトルには会社名を追加しません。</span><span class="sxs-lookup"><span data-stu-id="761fc-113">Don't append your company name to the title.</span></span></li></ul>|
|<ul><li><span data-ttu-id="761fc-114">タイトルには短くわかりやすい名前を使用します。</span><span class="sxs-lookup"><span data-stu-id="761fc-114">Use short descriptive names in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="761fc-115">アドインのタイトルには、「アドイン」、「Word」、「Office」などの文字列を追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="761fc-115">Don't append strings such as "add-in," "for Word," or "for Office" to the title of your add-in.</span></span></li></ul>|
|<ul><li><span data-ttu-id="761fc-116">アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。</span><span class="sxs-lookup"><span data-stu-id="761fc-116">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span></li></ul>||
|<ul><li><span data-ttu-id="761fc-117">アドインを Outlook 内で使用する場合を除き、アドインの下部に BrandBar などのブランド化の要素を含めます。</span><span class="sxs-lookup"><span data-stu-id="761fc-117">Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.</span></span></li></ul>||


## <a name="variants"></a><span data-ttu-id="761fc-118">バリアント</span><span class="sxs-lookup"><span data-stu-id="761fc-118">Variants</span></span>

<span data-ttu-id="761fc-p102">以下の図は、Office リボンの解像度が 1366x768 のさまざまな作業ウィンドウのサイズを示しています。Excel では、数式バーを収容するための縦のスペースが必要です。</span><span class="sxs-lookup"><span data-stu-id="761fc-p102">The following images show the various task pane sizes with the Office ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.</span></span>  

<span data-ttu-id="761fc-121">*図 2. Office 2016 デスクトップ作業ウィンドウのサイズ*</span><span class="sxs-lookup"><span data-stu-id="761fc-121">*Figure 2. Office 2016 desktop task pane sizes*</span></span>

![1366x768 のデスクトップ作業ウィンドウのサイズを示す図](../images/office-2016-taskpane-sizes.png)

- <span data-ttu-id="761fc-123">Excel - 320x455</span><span class="sxs-lookup"><span data-stu-id="761fc-123">Excel - 320x455</span></span>
- <span data-ttu-id="761fc-124">PowerPoint - 320x531</span><span class="sxs-lookup"><span data-stu-id="761fc-124">PowerPoint - 320x531</span></span>
- <span data-ttu-id="761fc-125">Word - 320x531</span><span class="sxs-lookup"><span data-stu-id="761fc-125">Word - 320x531</span></span>
- <span data-ttu-id="761fc-126">Outlook - 348x535</span><span class="sxs-lookup"><span data-stu-id="761fc-126">Outlook - 348x535</span></span>

<br/>

<span data-ttu-id="761fc-127">*図 3. Office 365 の作業ウィンドウのサイズ*</span><span class="sxs-lookup"><span data-stu-id="761fc-127">*Figure 3. Office 365 task pane sizes*</span></span>

![1366x768 のデスクトップ作業ウィンドウのサイズを示す図](../images/office-365-taskpane-sizes.png)

- <span data-ttu-id="761fc-129">Excel - 350x378</span><span class="sxs-lookup"><span data-stu-id="761fc-129">Excel - 350x378</span></span>
- <span data-ttu-id="761fc-130">PowerPoint - 348x391</span><span class="sxs-lookup"><span data-stu-id="761fc-130">PowerPoint - 348x391</span></span>
- <span data-ttu-id="761fc-131">Word - 329x445</span><span class="sxs-lookup"><span data-stu-id="761fc-131">Word - 329x445</span></span>
- <span data-ttu-id="761fc-132">Outlook (on the web) - 320x570</span><span class="sxs-lookup"><span data-stu-id="761fc-132">Outlook (on the web) - 320x570</span></span>

## <a name="personality-menu"></a><span data-ttu-id="761fc-133">パーソナル メニュー</span><span class="sxs-lookup"><span data-stu-id="761fc-133">Personality menu</span></span>

<span data-ttu-id="761fc-p103">パーソナル メニューは、アドインの右上付近にあるナビゲーション要素やコマンド要素の妨げになる可能性があります。Windows と Mac でのパーソナル メニューの現在のサイズを次に示します。</span><span class="sxs-lookup"><span data-stu-id="761fc-p103">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="761fc-136">Windows の場合、パーソナル メニューは 12x32 ピクセルを測定します (図を参照)。</span><span class="sxs-lookup"><span data-stu-id="761fc-136">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="761fc-137">*図 4. Windows のパーソナル メニュー*</span><span class="sxs-lookup"><span data-stu-id="761fc-137">*Figure 4. Personality menu on Windows*</span></span>

![Windows デスクトップのパーソナル メニューを示す図](../images/personality-menu-win.png)

<span data-ttu-id="761fc-139">Mac の場合、パーソナル メニューは 26x26 ピクセルを測定しますが、右から 8 ピクセル内側、上から 6 ピクセルの位置にフロートします。これにより、スペースは 34x32 ピクセルに増加します (図を参照)。</span><span class="sxs-lookup"><span data-stu-id="761fc-139">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="761fc-140">*図 5. Mac のパーソナル メニュー*</span><span class="sxs-lookup"><span data-stu-id="761fc-140">*Figure 5. Personality menu on Mac*</span></span>

![Mac デスクトップのパーソナル メニューを示す図](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="761fc-142">実装</span><span class="sxs-lookup"><span data-stu-id="761fc-142">Implementation</span></span>

<span data-ttu-id="761fc-143">作業ウィンドウを実装するサンプルについては、GitHub の「[Excel アドインの JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="761fc-143">For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub.</span></span> 


## <a name="see-also"></a><span data-ttu-id="761fc-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="761fc-144">See also</span></span>

- [<span data-ttu-id="761fc-145">Office アドインでの Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="761fc-145">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md) 
- [<span data-ttu-id="761fc-146">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="761fc-146">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)

