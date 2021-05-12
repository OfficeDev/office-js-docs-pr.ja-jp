---
title: コンテンツ Office アドイン
description: コンテンツ アドインは、Excel または PowerPoint ドキュメントに直接埋め込むことができるサーフェイスです。これでは、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 89c8c7eb96d6b1cf6484133b52acbe7d0ccb5298
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/12/2021
ms.locfileid: "52329997"
---
# <a name="content-office-add-ins"></a><span data-ttu-id="b7f0a-103">コンテンツ Office アドイン</span><span class="sxs-lookup"><span data-stu-id="b7f0a-103">Content Office Add-ins</span></span>

<span data-ttu-id="b7f0a-104">コンテンツ アドインは、Excel または PowerPoint ドキュメントに直接埋め込むことができるサーフェイスです。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-104">Content add-ins are surfaces that can be embedded directly into Excel or PowerPoint documents.</span></span> <span data-ttu-id="b7f0a-105">コンテンツ アドインにより、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-105">Content add-ins give users access to interface controls that run code to modify documents or display data from a data source.</span></span> <span data-ttu-id="b7f0a-106">機能を直接ドキュメントに埋め込む場合は、コンテンツ アドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-106">Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="b7f0a-107">*図 1. コンテンツ アドインの一般的なレイアウト*</span><span class="sxs-lookup"><span data-stu-id="b7f0a-107">*Figure 1. Typical layout for content add-ins*</span></span>

![アプリケーション内のコンテンツ アドインの一般的Officeレイアウト](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="b7f0a-109">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="b7f0a-109">Best practices</span></span>

- <span data-ttu-id="b7f0a-110">アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-110">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="b7f0a-111">アドインの下部に BrandBar などのブランド化の要素を含めます (Excel、および PowerPoint アドインにのみ適用)。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-111">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Excel and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="b7f0a-112">バリエーション</span><span class="sxs-lookup"><span data-stu-id="b7f0a-112">Variants</span></span>

<span data-ttu-id="b7f0a-113">デスクトップとデスクトップのExcelおよびPowerPointのOfficeのMicrosoft 365サイズは、ユーザーが指定します。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-113">Content add-in sizes for Excel and PowerPoint in Office desktop and Microsoft 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="b7f0a-114">パーソナル メニュー</span><span class="sxs-lookup"><span data-stu-id="b7f0a-114">Personality menu</span></span>

<span data-ttu-id="b7f0a-p102">パーソナル メニューは、アドインの右上付近にあるナビゲーション要素やコマンド要素の妨げになる可能性があります。Windows と Mac でのパーソナル メニューの現在のサイズを次に示します。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-p102">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="b7f0a-117">Windows の場合、パーソナル メニューは 12x32 ピクセルを測定します (図を参照)。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-117">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="b7f0a-118">*図 2. Windows のパーソナル メニュー*</span><span class="sxs-lookup"><span data-stu-id="b7f0a-118">*Figure 2. Personality menu on Windows*</span></span>

![デスクトップ上の 12x32 ピクセルのWindowsメニュー](../images/personality-menu-win.png)

<span data-ttu-id="b7f0a-120">Mac の場合、パーソナル メニューは 26x26 ピクセルを測定しますが、右から 8 ピクセル内側、上から 6 ピクセルの位置にフロートします。これにより、占有スペースは 34x32 ピクセルに増加します (図を参照)。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-120">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="b7f0a-121">*図 3. Mac のパーソナル メニュー*</span><span class="sxs-lookup"><span data-stu-id="b7f0a-121">*Figure 3. Personality menu on Mac*</span></span>

![Mac デスクトップの 34x32 ピクセルのパーソナリティ メニュー](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="b7f0a-123">実装</span><span class="sxs-lookup"><span data-stu-id="b7f0a-123">Implementation</span></span>

<span data-ttu-id="b7f0a-124">コンテンツ アドインの実装サンプルについては、GitHub の「[Excel コンテンツ アドイン Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-124">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="b7f0a-125">サポートに関する考慮事項</span><span class="sxs-lookup"><span data-stu-id="b7f0a-125">Support considerations</span></span>

- <span data-ttu-id="b7f0a-126">特定のアプリケーションまたはプラットフォームでOfficeアドインが動作Office[確認します](../overview/office-add-in-availability.md)。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-126">Check to see if your Office Add-in will work on a [specific Office application or platform](../overview/office-add-in-availability.md).</span></span>
- <span data-ttu-id="b7f0a-127">コンテンツ アドインによっては、Excel または PowerPoint の読み取りと書き込みのためにユーザーがアドインを「信頼」する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-127">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="b7f0a-128">アドインのマニフェストには、ユーザーに必要とされる[アクセス許可のレベル](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)を宣言することができます。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-128">You can declare what [level of permissions](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) you want your user to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="b7f0a-p104">コンテンツ アドインは Office 2013 以降のバージョンの Excel および PowerPoint でサポートされています。 Office Web アドインをサポートしていない Office のバージョンでアドインを開くと、アドインはイメージとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="b7f0a-p104">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later. If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="b7f0a-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="b7f0a-131">See also</span></span>

- [<span data-ttu-id="b7f0a-132">Office アドインの Office クライアント アプリケーションとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="b7f0a-132">Office client application and platform availability for Office Add-ins</span></span>](../overview/office-add-in-availability.md)
- [<span data-ttu-id="b7f0a-133">ファブリック コア (Office アドイン)</span><span class="sxs-lookup"><span data-stu-id="b7f0a-133">Fabric Core in Office Add-ins</span></span>](fabric-core.md)
- [<span data-ttu-id="b7f0a-134">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="b7f0a-134">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
- [<span data-ttu-id="b7f0a-135">アドインでの API 使用についてアクセス許可を要求する</span><span class="sxs-lookup"><span data-stu-id="b7f0a-135">Requesting permissions for API use in add-ins</span></span>](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
