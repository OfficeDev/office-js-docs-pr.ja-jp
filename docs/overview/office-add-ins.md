---
title: Office アドイン プラットフォームの概要 | Microsoft Docs
description: HTML、CSS、JavaScript などの一般的な Web テクノロジを使用し、Word、Excel、PowerPoint、OneNote、Project、Outlook を拡張および対話操作できます。
ms.date: 02/13/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 6b162a166bda0c988f5fbbaade3b0bef4b650984
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094072"
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="e3ece-103">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="e3ece-103">Office Add-ins platform overview</span></span>

<span data-ttu-id="e3ece-104">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents.</span><span class="sxs-lookup"><span data-stu-id="e3ece-104">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents.</span></span> <span data-ttu-id="e3ece-105">With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook.</span><span class="sxs-lookup"><span data-stu-id="e3ece-105">With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook.</span></span> <span data-ttu-id="e3ece-106">Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span><span class="sxs-lookup"><span data-stu-id="e3ece-106">Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span></span>

![Office アドインの拡張性の画像](../images/addins-overview.png)

<span data-ttu-id="e3ece-108">Office Add-ins can do almost anything a webpage can do inside a browser.</span><span class="sxs-lookup"><span data-stu-id="e3ece-108">Office Add-ins can do almost anything a webpage can do inside a browser.</span></span> <span data-ttu-id="e3ece-109">Use the Office Add-ins platform to:</span><span class="sxs-lookup"><span data-stu-id="e3ece-109">Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="e3ece-110">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more.</span><span class="sxs-lookup"><span data-stu-id="e3ece-110">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more.</span></span> <span data-ttu-id="e3ece-111">For example, use Microsoft Graph API to connect to data that drives productivity.</span><span class="sxs-lookup"><span data-stu-id="e3ece-111">For example, use Microsoft Graph API to connect to data that drives productivity.</span></span>

-  <span data-ttu-id="e3ece-112">**Office ドキュメントに埋め込み可能な充実した対話型のオブジェクトを新しく作成する** - マップやグラフ、ユーザーが自分の Excel スプレッドシートや PowerPoint プレゼンテーションに追加できる対話型の視覚化などを埋め込みます。</span><span class="sxs-lookup"><span data-stu-id="e3ece-112">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span>

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a><span data-ttu-id="e3ece-113">Office アドインが COM アドインおよび VSTO アドインと異なる点</span><span class="sxs-lookup"><span data-stu-id="e3ece-113">How are Office Add-ins different from COM and VSTO add-ins?</span></span>

<span data-ttu-id="e3ece-114">COM or VSTO add-ins are earlier Office integration solutions that run only on Office on Windows.</span><span class="sxs-lookup"><span data-stu-id="e3ece-114">COM or VSTO add-ins are earlier Office integration solutions that run only on Office on Windows.</span></span> <span data-ttu-id="e3ece-115">Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client.</span><span class="sxs-lookup"><span data-stu-id="e3ece-115">Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client.</span></span> <span data-ttu-id="e3ece-116">For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI.</span><span class="sxs-lookup"><span data-stu-id="e3ece-116">For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI.</span></span> <span data-ttu-id="e3ece-117">When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span><span class="sxs-lookup"><span data-stu-id="e3ece-117">When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span>

![Office アドインを使用する理由の画像](../images/why.png)

<span data-ttu-id="e3ece-119">Office アドインは、VBA、COM、または VSTO を使用して作成されたアドインと比較して、次のような利点があります。</span><span class="sxs-lookup"><span data-stu-id="e3ece-119">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span>

- <span data-ttu-id="e3ece-120">Cross-platform support.</span><span class="sxs-lookup"><span data-stu-id="e3ece-120">Cross-platform support.</span></span> <span data-ttu-id="e3ece-121">Office Add-ins run in Office on the web, Windows, Mac, and iPad.</span><span class="sxs-lookup"><span data-stu-id="e3ece-121">Office Add-ins run in Office on the web, Windows, Mac, and iPad.</span></span>

- <span data-ttu-id="e3ece-122">Centralized deployment and distribution.</span><span class="sxs-lookup"><span data-stu-id="e3ece-122">Centralized deployment and distribution.</span></span> <span data-ttu-id="e3ece-123">Admins can deploy Office Add-ins centrally across an organization.</span><span class="sxs-lookup"><span data-stu-id="e3ece-123">Admins can deploy Office Add-ins centrally across an organization.</span></span>

- <span data-ttu-id="e3ece-124">Easy access via AppSource.</span><span class="sxs-lookup"><span data-stu-id="e3ece-124">Easy access via AppSource.</span></span> <span data-ttu-id="e3ece-125">You can make your solution available to a broad audience by submitting it to AppSource.</span><span class="sxs-lookup"><span data-stu-id="e3ece-125">You can make your solution available to a broad audience by submitting it to AppSource.</span></span>

- <span data-ttu-id="e3ece-126">Based on standard web technology.</span><span class="sxs-lookup"><span data-stu-id="e3ece-126">Based on standard web technology.</span></span> <span data-ttu-id="e3ece-127">You can use any library you like to build Office Add-ins.</span><span class="sxs-lookup"><span data-stu-id="e3ece-127">You can use any library you like to build Office Add-ins.</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="e3ece-128">Office アドインのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="e3ece-128">Components of an Office Add-in</span></span>

<span data-ttu-id="e3ece-129">An Office Add-in includes two basic components: an XML manifest file, and your own web application.</span><span class="sxs-lookup"><span data-stu-id="e3ece-129">An Office Add-in includes two basic components: an XML manifest file, and your own web application.</span></span> <span data-ttu-id="e3ece-130">The manifest defines various settings, including how your add-in integrates with Office clients.</span><span class="sxs-lookup"><span data-stu-id="e3ece-130">The manifest defines various settings, including how your add-in integrates with Office clients.</span></span> <span data-ttu-id="e3ece-131">Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="e3ece-131">Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

### <a name="manifest"></a><span data-ttu-id="e3ece-132">マニフェスト</span><span class="sxs-lookup"><span data-stu-id="e3ece-132">Manifest</span></span>

<span data-ttu-id="e3ece-133">マニフェストは、次のようなアドインの設定と機能を指定する XML ファイルです。</span><span class="sxs-lookup"><span data-stu-id="e3ece-133">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span>

- <span data-ttu-id="e3ece-134">アドインの表示名、説明、ID、バージョン、および既定のロケール。</span><span class="sxs-lookup"><span data-stu-id="e3ece-134">The add-in's display name, description, ID, version, and default locale.</span></span>

- <span data-ttu-id="e3ece-135">Office とアドインを統合する方法。</span><span class="sxs-lookup"><span data-stu-id="e3ece-135">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="e3ece-136">アドインのアクセス許可レベルとデータ アクセスの要件。</span><span class="sxs-lookup"><span data-stu-id="e3ece-136">The permission level and data access requirements for the add-in.</span></span>

### <a name="web-app"></a><span data-ttu-id="e3ece-137">Web アプリケーション</span><span class="sxs-lookup"><span data-stu-id="e3ece-137">Web app</span></span>

<span data-ttu-id="e3ece-138">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource.</span><span class="sxs-lookup"><span data-stu-id="e3ece-138">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource.</span></span> <span data-ttu-id="e3ece-139">However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js).</span><span class="sxs-lookup"><span data-stu-id="e3ece-139">However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js).</span></span> <span data-ttu-id="e3ece-140">To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span><span class="sxs-lookup"><span data-stu-id="e3ece-140">To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span>

<span data-ttu-id="e3ece-141">*図 2. Hello World Office アドインのコンポーネント*</span><span class="sxs-lookup"><span data-stu-id="e3ece-141">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Hello World アドインのコンポーネント](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="e3ece-143">Office クライアントの拡張と、Office クライアントとの対話</span><span class="sxs-lookup"><span data-stu-id="e3ece-143">Extending and interacting with Office clients</span></span>

<span data-ttu-id="e3ece-144">Office アドインは、Office ホスト アプリケーション内で次を実行できます。</span><span class="sxs-lookup"><span data-stu-id="e3ece-144">Office Add-ins can do the following within an Office host application:</span></span>

-  <span data-ttu-id="e3ece-145">機能の拡張 (任意の Office アプリケーション)</span><span class="sxs-lookup"><span data-stu-id="e3ece-145">Extend functionality (any Office application)</span></span>

-  <span data-ttu-id="e3ece-146">新しいオブジェクトの作成 (Excel または PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="e3ece-146">Create new objects (Excel or PowerPoint)</span></span>
 
### <a name="extend-office-functionality"></a><span data-ttu-id="e3ece-147">Office 機能の拡張</span><span class="sxs-lookup"><span data-stu-id="e3ece-147">Extend Office functionality</span></span>

<span data-ttu-id="e3ece-148">次の方法で、Office アプリケーションに新しい機能を追加できます。</span><span class="sxs-lookup"><span data-stu-id="e3ece-148">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="e3ece-149">カスタム リボン ボタンとメニュー コマンド ("アドイン コマンド" と総称されます)</span><span class="sxs-lookup"><span data-stu-id="e3ece-149">Custom ribbon buttons and menu commands (collectively called “add-in commands”)</span></span>

-  <span data-ttu-id="e3ece-150">挿入可能な作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e3ece-150">Insertable task panes</span></span>

<span data-ttu-id="e3ece-151">カスタムの UI と作業ウィンドウは、アドイン マニフェストで指定されます。</span><span class="sxs-lookup"><span data-stu-id="e3ece-151">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="e3ece-152">カスタム ボタンとメニュー コマンド</span><span class="sxs-lookup"><span data-stu-id="e3ece-152">Custom buttons and menu commands</span></span>  

<span data-ttu-id="e3ece-153">You can add custom ribbon buttons and menu items to the ribbon in Office on the web and Windows.</span><span class="sxs-lookup"><span data-stu-id="e3ece-153">You can add custom ribbon buttons and menu items to the ribbon in Office on the web and Windows.</span></span> <span data-ttu-id="e3ece-154">This makes it easy for users to access your add-in directly from their Office application.</span><span class="sxs-lookup"><span data-stu-id="e3ece-154">This makes it easy for users to access your add-in directly from their Office application.</span></span> <span data-ttu-id="e3ece-155">Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span><span class="sxs-lookup"><span data-stu-id="e3ece-155">Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="e3ece-156">*図 3. リボンにあるアドイン コマンド*</span><span class="sxs-lookup"><span data-stu-id="e3ece-156">*Figure 3. Add-in commands in the ribbon*</span></span>

![カスタム ボタンとメニュー コマンド](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a><span data-ttu-id="e3ece-158">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e3ece-158">Task panes</span></span>  

<span data-ttu-id="e3ece-159">You can use task panes in addition to add-in commands to enable users to interact with your solution.</span><span class="sxs-lookup"><span data-stu-id="e3ece-159">You can use task panes in addition to add-in commands to enable users to interact with your solution.</span></span> <span data-ttu-id="e3ece-160">Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane.</span><span class="sxs-lookup"><span data-stu-id="e3ece-160">Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane.</span></span> <span data-ttu-id="e3ece-161">Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span><span class="sxs-lookup"><span data-stu-id="e3ece-161">Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span>

<span data-ttu-id="e3ece-162">*図 4. 作業ウィンドウ*</span><span class="sxs-lookup"><span data-stu-id="e3ece-162">*Figure 4. Task pane*</span></span>

![作業ウィンドウとアドイン コマンドを使用する](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="e3ece-164">Outlook の機能を拡張する</span><span class="sxs-lookup"><span data-stu-id="e3ece-164">Extend Outlook functionality</span></span>

<span data-ttu-id="e3ece-165">Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it.</span><span class="sxs-lookup"><span data-stu-id="e3ece-165">Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it.</span></span> <span data-ttu-id="e3ece-166">They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span><span class="sxs-lookup"><span data-stu-id="e3ece-166">They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="e3ece-167">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences.</span><span class="sxs-lookup"><span data-stu-id="e3ece-167">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences.</span></span> <span data-ttu-id="e3ece-168">In most cases, an Outlook add-in runs without modification in the Outlook host application to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span><span class="sxs-lookup"><span data-stu-id="e3ece-168">In most cases, an Outlook add-in runs without modification in the Outlook host application to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span>

<span data-ttu-id="e3ece-169">Outlook アドインの概要については、「[Outlook アドインの概要](../outlook/outlook-add-ins-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3ece-169">For an overview of Outlook add-ins, see [Outlook add-ins overview](../outlook/outlook-add-ins-overview.md).</span></span>

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="e3ece-170">Office ドキュメント内に新しいオブジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="e3ece-170">Create new objects in Office documents</span></span>

<span data-ttu-id="e3ece-171">You can embed web-based objects called content add-ins within Excel and PowerPoint documents.</span><span class="sxs-lookup"><span data-stu-id="e3ece-171">You can embed web-based objects called content add-ins within Excel and PowerPoint documents.</span></span> <span data-ttu-id="e3ece-172">With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span><span class="sxs-lookup"><span data-stu-id="e3ece-172">With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="e3ece-173">*図 5. コンテンツ アドイン*</span><span class="sxs-lookup"><span data-stu-id="e3ece-173">*Figure 5. Content add-in*</span></span>

![コンテンツ アドインと呼ばれる Web ベースのオブジェクトを埋め込む](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="e3ece-175">Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="e3ece-175">Office JavaScript APIs</span></span>

<span data-ttu-id="e3ece-176">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services.</span><span class="sxs-lookup"><span data-stu-id="e3ece-176">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services.</span></span> <span data-ttu-id="e3ece-177">There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project.</span><span class="sxs-lookup"><span data-stu-id="e3ece-177">There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project.</span></span> <span data-ttu-id="e3ece-178">There are also more extensive host-specific object models for Excel and Word.</span><span class="sxs-lookup"><span data-stu-id="e3ece-178">There are also more extensive host-specific object models for Excel and Word.</span></span> <span data-ttu-id="e3ece-179">These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span><span class="sxs-lookup"><span data-stu-id="e3ece-179">These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span></span>  

## <a name="next-steps"></a><span data-ttu-id="e3ece-180">次のステップ</span><span class="sxs-lookup"><span data-stu-id="e3ece-180">Next steps</span></span>

<span data-ttu-id="e3ece-181">Office アドインの開発の詳細については、「[Office アドインを構築する](../overview/office-add-ins-fundamentals.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3ece-181">For a more detailed introduction to developing Office Add-ins, see [Building Office Add-ins](../overview/office-add-ins-fundamentals.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e3ece-182">関連項目</span><span class="sxs-lookup"><span data-stu-id="e3ece-182">See also</span></span>

- [<span data-ttu-id="e3ece-183">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="e3ece-183">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="e3ece-184">Office アドインの中心概念</span><span class="sxs-lookup"><span data-stu-id="e3ece-184">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="e3ece-185">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="e3ece-185">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="e3ece-186">Office アドインを設計する</span><span class="sxs-lookup"><span data-stu-id="e3ece-186">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="e3ece-187">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="e3ece-187">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="e3ece-188">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="e3ece-188">Publish Office Add-ins</span></span>](../publish/publish.md)
