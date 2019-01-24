---
title: コンテンツ Office アドイン
description: コンテンツ アドインは、Excel または PowerPoint ドキュメントに直接埋め込むことができるサーフェイスです。これでは、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 12/13/2018
localization_priority: Priority
ms.openlocfilehash: 8114fe1d95739ba14fe54c379e44148f451010b2
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389397"
---
# <a name="content-office-add-ins"></a><span data-ttu-id="0c98d-103">コンテンツ Office アドイン</span><span class="sxs-lookup"><span data-stu-id="0c98d-103">Content Office Add-ins</span></span>

<span data-ttu-id="0c98d-104">コンテンツ アドインは、Excel または PowerPoint ドキュメントに直接埋め込むことができるサーフェイスです。</span><span class="sxs-lookup"><span data-stu-id="0c98d-104">Content add-ins are surfaces that can be embedded directly into Excel or PowerPoint documents.</span></span> <span data-ttu-id="0c98d-105">コンテンツ アドインにより、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="0c98d-105">Content add-ins give users access to interface controls that run code to modify documents or display data from a data source.</span></span> <span data-ttu-id="0c98d-106">機能を直接ドキュメントに埋め込む場合は、コンテンツ アドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="0c98d-106">Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="0c98d-107">*図 1. コンテンツ アドインの一般的なレイアウト*</span><span class="sxs-lookup"><span data-stu-id="0c98d-107">*Figure 1. Typical layout for content add-ins*</span></span>

![コンテンツ アドインの一般的なレイアウトを表示する画像の例](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="0c98d-109">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="0c98d-109">Best practices</span></span>

- <span data-ttu-id="0c98d-110">アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。</span><span class="sxs-lookup"><span data-stu-id="0c98d-110">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="0c98d-111">アドインの下部に BrandBar などのブランド化の要素を含めます (Excel、および PowerPoint アドインにのみ適用)。</span><span class="sxs-lookup"><span data-stu-id="0c98d-111">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Excel and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="0c98d-112">バリエーション</span><span class="sxs-lookup"><span data-stu-id="0c98d-112">Variants</span></span>

<span data-ttu-id="0c98d-113">Office デスクトップと Office 365 の Excel、PowerPoint のコンテンツ アドインのサイズはユーザーが指定します。</span><span class="sxs-lookup"><span data-stu-id="0c98d-113">Content add-in sizes for Excel and PowerPoint in Office desktop and Office 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="0c98d-114">パーソナル メニュー</span><span class="sxs-lookup"><span data-stu-id="0c98d-114">Personality menu</span></span>

<span data-ttu-id="0c98d-p102">パーソナル メニューは、アドインの右上付近にあるナビゲーション要素やコマンド要素の妨げになる可能性があります。Windows と Mac でのパーソナル メニューの現在のサイズを次に示します。</span><span class="sxs-lookup"><span data-stu-id="0c98d-p102">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="0c98d-117">Windows の場合、パーソナル メニューは 12x32 ピクセルを測定します (図を参照)。</span><span class="sxs-lookup"><span data-stu-id="0c98d-117">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="0c98d-118">*図 2. Windows のパーソナル メニュー*</span><span class="sxs-lookup"><span data-stu-id="0c98d-118">*Figure 2. Personality menu on Windows*</span></span> 

![Windows デスクトップのパーソナル メニューを示す図](../images/personality-menu-win.png)


<span data-ttu-id="0c98d-120">Mac の場合、パーソナル メニューは 26x26 ピクセルを測定しますが、右から 8 ピクセル内側、上から 6 ピクセルの位置にフロートします。これにより、占有スペースは 34x32 ピクセルに増加します (図を参照)。</span><span class="sxs-lookup"><span data-stu-id="0c98d-120">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="0c98d-121">*図 3. Mac のパーソナル メニュー*</span><span class="sxs-lookup"><span data-stu-id="0c98d-121">*Figure 3. Personality menu on Mac*</span></span>

![Mac デスクトップのパーソナル メニューを示す図](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="0c98d-123">実装</span><span class="sxs-lookup"><span data-stu-id="0c98d-123">Implementation</span></span>

<span data-ttu-id="0c98d-124">コンテンツ アドインの実装サンプルについては、GitHub の「[Excel コンテンツ アドイン Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0c98d-124">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="0c98d-125">サポートに関する考慮事項</span><span class="sxs-lookup"><span data-stu-id="0c98d-125">Support considerations</span></span>
- <span data-ttu-id="0c98d-126">使用している Office アドインが[特定の Office ホスト プラットフォーム](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)で動作するかどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="0c98d-126">Check to see if your Office Add-in will work on a [specific Office host platform](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).</span></span> 
- <span data-ttu-id="0c98d-127">コンテンツ アドインによっては、Excel または PowerPoint の読み取りと書き込みのためにユーザーがアドインを「信頼」する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c98d-127">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="0c98d-128">アドインのマニフェストには、ユーザーに必要とされる[アクセス許可のレベル](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)を宣言することができます。</span><span class="sxs-lookup"><span data-stu-id="0c98d-128">You can declare what [level of permissions](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) you want your user to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="0c98d-129">コンテンツ アドインは Office 2013 以降のバージョンの Excel および PowerPoint でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="0c98d-129">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later.</span></span> <span data-ttu-id="0c98d-130">Office Web アドインをサポートしていない Office のバージョンでアドインを開くと、アドインはイメージとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="0c98d-130">If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="0c98d-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="0c98d-131">See also</span></span>
- [<span data-ttu-id="0c98d-132">Office アドインのホストとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="0c98d-132">Office Add-in host and platform availability</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="0c98d-133">Office アドインの Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="0c98d-133">Office UI Fabric in Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/design/office-ui-fabric) 
- [<span data-ttu-id="0c98d-134">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="0c98d-134">UX design patterns for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/design/ux-design-pattern-templates)
- [<span data-ttu-id="0c98d-135">コンテンツ アドインと作業ウィンドウ アドインでの API 使用についてアクセス許可を要求する</span><span class="sxs-lookup"><span data-stu-id="0c98d-135">Requesting permissions for API use in content and task pane add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
