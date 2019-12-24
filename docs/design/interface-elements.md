---
title: Office アドイン用の Office の UI 要素
description: ''
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: 1f680d93c995d306717c7b5a81af073f378b0fbd
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851242"
---
# <a name="office-ui-elements-for-office-add-ins"></a><span data-ttu-id="7e1f0-102">Office アドイン用の Office の UI 要素</span><span class="sxs-lookup"><span data-stu-id="7e1f0-102">Office UI elements for Office Add-ins</span></span>

<span data-ttu-id="7e1f0-p101">アドイン コマンドおよび HTML のコンテナーを含むいくつかの種類の UI 要素を使用して Office UI を拡張することができます。これらの UI 要素は、Office の元々の拡張機能のように自然に、あらゆるプラットフォームで使えます。これらのいずれの要素にも、Web ベースのカスタム コードを挿入できます。</span><span class="sxs-lookup"><span data-stu-id="7e1f0-p101">You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers. These UI elements look like a natural extension of Office and work across platforms. You can insert your custom web-based code into any of these elements.</span></span>

<span data-ttu-id="7e1f0-106">次の図は、作成できる Office UI 要素の種類を示しています。</span><span class="sxs-lookup"><span data-stu-id="7e1f0-106">The following image shows the types of Office UI elements that you can create.</span></span>

![Office ドキュメントのリボン、タスク ウィンドウ、ダイアログ ボックス上のアドイン コマンドを示す図](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a><span data-ttu-id="7e1f0-108">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="7e1f0-108">Add-in commands</span></span>

<span data-ttu-id="7e1f0-p102">[アドイン コマンド](add-in-commands.md)を使って、アドインへのエントリ ポイントを Office リボンに追加します。コマンドは、JavaScript コードを実行するか、HTML コンテナーを起動することによって、アドインのアクションを開始します。2 種類のアドイン コマンドを作成できます。</span><span class="sxs-lookup"><span data-stu-id="7e1f0-p102">Use [add-in commands](add-in-commands.md) to add entry points to your add-in to the Office ribbon. Commands start actions in your add-in either by running JavaScript code, or by launching an HTML container. You can create two types of add-in commands.</span></span>

|<span data-ttu-id="7e1f0-112">**コマンドの種類**</span><span class="sxs-lookup"><span data-stu-id="7e1f0-112">**Command type**</span></span>|<span data-ttu-id="7e1f0-113">**説明**</span><span class="sxs-lookup"><span data-stu-id="7e1f0-113">**Description**</span></span>|
|:---------------|:--------------|
|<span data-ttu-id="7e1f0-114">リボンのボタン、メニュー、およびタブ</span><span class="sxs-lookup"><span data-stu-id="7e1f0-114">Ribbon buttons, menus, and tabs</span></span>|<span data-ttu-id="7e1f0-p103">Office の既定のリボンにカスタム ボタン、メニュー (ドロップダウン)、またはタブを追加するのに使用します。ボタンやメニューは、Office でのアクションをトリガーするのに使用します。タブは、ボタンやメニューをグループ化し整理するのに使用します。</span><span class="sxs-lookup"><span data-stu-id="7e1f0-p103">Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office. Use Buttons and menus to trigger an action in Office. Use tabs to group and organize buttons and menus.</span></span>|
|<span data-ttu-id="7e1f0-118">コンテキスト メニュー</span><span class="sxs-lookup"><span data-stu-id="7e1f0-118">Context menus</span></span>| <span data-ttu-id="7e1f0-p104">既定のコンテキスト メニューを拡張するために使用します。Office ドキュメントのテキストまたは Excel のテーブルを右クリックすると、コンテキスト メニューが表示されます。</span><span class="sxs-lookup"><span data-stu-id="7e1f0-p104">Use to extend the default context menu. Context menus are displayed when users right-click text in an Office document or a table in Excel.</span></span>| 

## <a name="html-containers"></a><span data-ttu-id="7e1f0-121">HTML コンテナー</span><span class="sxs-lookup"><span data-stu-id="7e1f0-121">HTML containers</span></span>

<span data-ttu-id="7e1f0-p105">HTML コンテナーは、Office クライアント内に HTML ベースの UI コードを埋め込むのに使用します。その Web ページで、Office の JavaScript API を参照して、ドキュメント内でコンテンツを操作できるようになります。3 種類の HTML コンテナーを作成できます。</span><span class="sxs-lookup"><span data-stu-id="7e1f0-p105">Use HTML containers to embed HTML-based UI code within Office clients. These web pages can then reference the Office JavaScript API to interact with content in the document. You can create three types of HTML containers.</span></span>

|<span data-ttu-id="7e1f0-125">**HTML コンテナー**</span><span class="sxs-lookup"><span data-stu-id="7e1f0-125">**HTML container**</span></span>|<span data-ttu-id="7e1f0-126">**説明**</span><span class="sxs-lookup"><span data-stu-id="7e1f0-126">**Description**</span></span>|
|:-----------------|:--------------|
|[<span data-ttu-id="7e1f0-127">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7e1f0-127">Task panes</span></span>](task-pane-add-ins.md)|<span data-ttu-id="7e1f0-p106">Office ドキュメントの右側のウィンドウに、カスタム UI を表示します。作業ウィンドウを使用して、Office ドキュメントでアドインを横に並べて表示して操作できるようにします。</span><span class="sxs-lookup"><span data-stu-id="7e1f0-p106">Display custom UI in the right pane of the Office document. Use task panes to allow users to interact with your add-in side-by-side with the Office document.</span></span>|
|[<span data-ttu-id="7e1f0-130">コンテンツ アドイン</span><span class="sxs-lookup"><span data-stu-id="7e1f0-130">Content add-ins</span></span>](content-add-ins.md)|<span data-ttu-id="7e1f0-p107">Office ドキュメントに埋め込まれているカスタム UI を表示します。コンテンツ アドインを使用して、Office ドキュメント内でアドインを直接操作できるようにします。たとえば、ビデオや、他のソースからのデータのビジュアル化などの外部コンテンツを表示します。</span><span class="sxs-lookup"><span data-stu-id="7e1f0-p107">Display custom UI embedded within Office documents. Use content add-ins to allow users to interact with your add-in directly within the Office document. For example, you might want to show external content such as videos or data visualizations from other sources.</span></span> |
|[<span data-ttu-id="7e1f0-134">ダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="7e1f0-134">Dialog boxes</span></span>](dialog-boxes.md)|<span data-ttu-id="7e1f0-p108">Office ドキュメントにオーバーレイした形でダイアログ ボックスの中にカスタム UI を表示します。フォーカスする必要がありスペースをより多く取る操作で、ドキュメント内で横並びにする必要がない操作には、ダイアログ ボックスを使用します。</span><span class="sxs-lookup"><span data-stu-id="7e1f0-p108">Display custom UI in a dialog box that overlays the Office document. Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.</span></span>|

## <a name="see-also"></a><span data-ttu-id="7e1f0-137">関連項目</span><span class="sxs-lookup"><span data-stu-id="7e1f0-137">See also</span></span>

- [<span data-ttu-id="7e1f0-138">Excel、Word、PowerPoint のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="7e1f0-138">Add-in commands for Excel, Word, and PowerPoint</span></span>](add-in-commands.md)
- [<span data-ttu-id="7e1f0-139">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7e1f0-139">Task panes</span></span>](task-pane-add-ins.md)
- [<span data-ttu-id="7e1f0-140">コンテンツ アドイン</span><span class="sxs-lookup"><span data-stu-id="7e1f0-140">Content add-ins</span></span>](content-add-ins.md)
- [<span data-ttu-id="7e1f0-141">ダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="7e1f0-141">Dialog boxes</span></span>](dialog-boxes.md)
