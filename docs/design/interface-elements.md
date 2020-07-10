---
title: Office アドイン用の Office の UI 要素
description: Office アドインのさまざまな種類の UI 要素の概要について説明します。
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: 5b9907924c674ed9db2294621123c394419d0c12
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093764"
---
# <a name="office-ui-elements-for-office-add-ins"></a><span data-ttu-id="6aca1-103">Office アドイン用の Office の UI 要素</span><span class="sxs-lookup"><span data-stu-id="6aca1-103">Office UI elements for Office Add-ins</span></span>

<span data-ttu-id="6aca1-104">You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers.</span><span class="sxs-lookup"><span data-stu-id="6aca1-104">You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers.</span></span> <span data-ttu-id="6aca1-105">These UI elements look like a natural extension of Office and work across platforms.</span><span class="sxs-lookup"><span data-stu-id="6aca1-105">These UI elements look like a natural extension of Office and work across platforms.</span></span> <span data-ttu-id="6aca1-106">You can insert your custom web-based code into any of these elements.</span><span class="sxs-lookup"><span data-stu-id="6aca1-106">You can insert your custom web-based code into any of these elements.</span></span>

<span data-ttu-id="6aca1-107">次の図は、作成できる Office UI 要素の種類を示しています。</span><span class="sxs-lookup"><span data-stu-id="6aca1-107">The following image shows the types of Office UI elements that you can create.</span></span>

![Office ドキュメントのリボン、タスク ウィンドウ、ダイアログ ボックス上のアドイン コマンドを示す図](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a><span data-ttu-id="6aca1-109">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="6aca1-109">Add-in commands</span></span>

<span data-ttu-id="6aca1-110">アドイン[コマンド](add-in-commands.md)を使用して、Office アプリリボンにアドインへのエントリポイントを追加します。</span><span class="sxs-lookup"><span data-stu-id="6aca1-110">Use [add-in commands](add-in-commands.md) to add entry points to your add-in to the Office app ribbon.</span></span> <span data-ttu-id="6aca1-111">コマンドは、JavaScript コードを実行するか、HTML コンテナーを起動することによって、アドインのアクションを開始します。</span><span class="sxs-lookup"><span data-stu-id="6aca1-111">Commands start actions in your add-in either by running JavaScript code, or by launching an HTML container.</span></span> <span data-ttu-id="6aca1-112">2 種類のアドイン コマンドを作成できます。</span><span class="sxs-lookup"><span data-stu-id="6aca1-112">You can create two types of add-in commands.</span></span>

|<span data-ttu-id="6aca1-113">**コマンドの種類**</span><span class="sxs-lookup"><span data-stu-id="6aca1-113">**Command type**</span></span>|<span data-ttu-id="6aca1-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="6aca1-114">**Description**</span></span>|
|:---------------|:--------------|
|<span data-ttu-id="6aca1-115">リボンのボタン、メニュー、およびタブ</span><span class="sxs-lookup"><span data-stu-id="6aca1-115">Ribbon buttons, menus, and tabs</span></span>|<span data-ttu-id="6aca1-116">Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office.</span><span class="sxs-lookup"><span data-stu-id="6aca1-116">Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office.</span></span> <span data-ttu-id="6aca1-117">Use Buttons and menus to trigger an action in Office.</span><span class="sxs-lookup"><span data-stu-id="6aca1-117">Use Buttons and menus to trigger an action in Office.</span></span> <span data-ttu-id="6aca1-118">Use tabs to group and organize buttons and menus.</span><span class="sxs-lookup"><span data-stu-id="6aca1-118">Use tabs to group and organize buttons and menus.</span></span>|
|<span data-ttu-id="6aca1-119">コンテキスト メニュー</span><span class="sxs-lookup"><span data-stu-id="6aca1-119">Context menus</span></span>| <span data-ttu-id="6aca1-120">Use to extend the default context menu.</span><span class="sxs-lookup"><span data-stu-id="6aca1-120">Use to extend the default context menu.</span></span> <span data-ttu-id="6aca1-121">Context menus are displayed when users right-click text in an Office document or a table in Excel.</span><span class="sxs-lookup"><span data-stu-id="6aca1-121">Context menus are displayed when users right-click text in an Office document or a table in Excel.</span></span>| 

## <a name="html-containers"></a><span data-ttu-id="6aca1-122">HTML コンテナー</span><span class="sxs-lookup"><span data-stu-id="6aca1-122">HTML containers</span></span>

<span data-ttu-id="6aca1-123">Use HTML containers to embed HTML-based UI code within Office clients.</span><span class="sxs-lookup"><span data-stu-id="6aca1-123">Use HTML containers to embed HTML-based UI code within Office clients.</span></span> <span data-ttu-id="6aca1-124">These web pages can then reference the Office JavaScript API to interact with content in the document.</span><span class="sxs-lookup"><span data-stu-id="6aca1-124">These web pages can then reference the Office JavaScript API to interact with content in the document.</span></span> <span data-ttu-id="6aca1-125">You can create three types of HTML containers.</span><span class="sxs-lookup"><span data-stu-id="6aca1-125">You can create three types of HTML containers.</span></span>

|<span data-ttu-id="6aca1-126">**HTML コンテナー**</span><span class="sxs-lookup"><span data-stu-id="6aca1-126">**HTML container**</span></span>|<span data-ttu-id="6aca1-127">**説明**</span><span class="sxs-lookup"><span data-stu-id="6aca1-127">**Description**</span></span>|
|:-----------------|:--------------|
|[<span data-ttu-id="6aca1-128">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="6aca1-128">Task panes</span></span>](task-pane-add-ins.md)|<span data-ttu-id="6aca1-129">Display custom UI in the right pane of the Office document.</span><span class="sxs-lookup"><span data-stu-id="6aca1-129">Display custom UI in the right pane of the Office document.</span></span> <span data-ttu-id="6aca1-130">Use task panes to allow users to interact with your add-in side-by-side with the Office document.</span><span class="sxs-lookup"><span data-stu-id="6aca1-130">Use task panes to allow users to interact with your add-in side-by-side with the Office document.</span></span>|
|[<span data-ttu-id="6aca1-131">コンテンツ アドイン</span><span class="sxs-lookup"><span data-stu-id="6aca1-131">Content add-ins</span></span>](content-add-ins.md)|<span data-ttu-id="6aca1-132">Display custom UI embedded within Office documents.</span><span class="sxs-lookup"><span data-stu-id="6aca1-132">Display custom UI embedded within Office documents.</span></span> <span data-ttu-id="6aca1-133">Use content add-ins to allow users to interact with your add-in directly within the Office document.</span><span class="sxs-lookup"><span data-stu-id="6aca1-133">Use content add-ins to allow users to interact with your add-in directly within the Office document.</span></span> <span data-ttu-id="6aca1-134">For example, you might want to show external content such as videos or data visualizations from other sources.</span><span class="sxs-lookup"><span data-stu-id="6aca1-134">For example, you might want to show external content such as videos or data visualizations from other sources.</span></span> |
|[<span data-ttu-id="6aca1-135">ダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="6aca1-135">Dialog boxes</span></span>](dialog-boxes.md)|<span data-ttu-id="6aca1-136">Display custom UI in a dialog box that overlays the Office document.</span><span class="sxs-lookup"><span data-stu-id="6aca1-136">Display custom UI in a dialog box that overlays the Office document.</span></span> <span data-ttu-id="6aca1-137">Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.</span><span class="sxs-lookup"><span data-stu-id="6aca1-137">Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.</span></span>|

## <a name="see-also"></a><span data-ttu-id="6aca1-138">関連項目</span><span class="sxs-lookup"><span data-stu-id="6aca1-138">See also</span></span>

- [<span data-ttu-id="6aca1-139">Excel、Word、PowerPoint のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="6aca1-139">Add-in commands for Excel, Word, and PowerPoint</span></span>](add-in-commands.md)
- [<span data-ttu-id="6aca1-140">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="6aca1-140">Task panes</span></span>](task-pane-add-ins.md)
- [<span data-ttu-id="6aca1-141">コンテンツ アドイン</span><span class="sxs-lookup"><span data-stu-id="6aca1-141">Content add-ins</span></span>](content-add-ins.md)
- [<span data-ttu-id="6aca1-142">ダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="6aca1-142">Dialog boxes</span></span>](dialog-boxes.md)
