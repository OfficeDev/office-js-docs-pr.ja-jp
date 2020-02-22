---
title: Office アドイン プラットフォームの概要 | Microsoft Docs
description: HTML、CSS、JavaScript などの一般的な Web テクノロジを使用し、Word、Excel、PowerPoint、OneNote、Project、Outlook を拡張および対話操作できます。
ms.date: 02/13/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: fd69f3e93b18d2d17ef7ba1b213bfb39b86b74f4
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165574"
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="1cf96-103">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="1cf96-103">Office Add-ins platform overview</span></span>

<span data-ttu-id="1cf96-p101">Office アドイン プラットフォームを使用すると、Office アプリケーションを拡張して、Office ドキュメント内のコンテンツを対話方式で操作するソリューションを構築できます。Office アドインでは、HTML、CSS、および JavaScript などの一般的な Web テクノロジを使用することで、Word、Excel、PowerPoint、OneNote、Project、および Outlook を拡張して対話できるようになります。Windows、Mac、iPad、およびブラウザーを含む複数のプラットフォームにわたって Office ソリューションを実行できます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p101">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook. Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span></span>

![Office アドインの拡張性の画像](../images/addins-overview.png)

<span data-ttu-id="1cf96-p102">Office アドインでは、ブラウザー内で Web ページが実行できる操作のほとんどすべてを実行できます。Office アドイン プラットフォームを使用して、次のことができます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p102">Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="1cf96-p103">**Office クライアントに新しい機能を追加する** - Office に外部データを取り込む、Office ドキュメントを自動化する、サード パーティの機能を Office クライアントで公開する、などがあります。たとえば、Microsoft Graph API を使用して、生産性の向上につながるデータに接続します。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p103">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.</span></span>

-  <span data-ttu-id="1cf96-112">**Office ドキュメントに埋め込み可能な充実した対話型のオブジェクトを新しく作成する** - マップやグラフ、ユーザーが自分の Excel スプレッドシートや PowerPoint プレゼンテーションに追加できる対話型の視覚化などを埋め込みます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-112">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span>

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a><span data-ttu-id="1cf96-113">Office アドインが COM アドインおよび VSTO アドインと異なる点</span><span class="sxs-lookup"><span data-stu-id="1cf96-113">How are Office Add-ins different from COM and VSTO add-ins?</span></span>

<span data-ttu-id="1cf96-p104">COM または VSTO アドインは、Office for Windows 上でのみ実行する以前の Office 統合ソリューションです。COM アドインとは異なり、Office アドインにはユーザーのデバイスまたは Office クライアントで実行されるコードは含まれません。Office アドインの場合、ホスト アプリケーション (たとえば Excel) がアドインのマニフェストを読み取り、アドインのカスタム リボン ボタンと UI のメニュー コマンドをフックします。これは必要に応じて、サンド ボックスのブラウザーのコンテキストで実行されるアドインの JavaScript と HTML を読み込みます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p104">COM or VSTO add-ins are earlier Office integration solutions that run only on Office on Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span>

![Office アドインを使用する理由の画像](../images/why.png)

<span data-ttu-id="1cf96-119">Office アドインは、VBA、COM、または VSTO を使用して作成されたアドインと比較して、次のような利点があります。</span><span class="sxs-lookup"><span data-stu-id="1cf96-119">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span>

- <span data-ttu-id="1cf96-p105">クロスプラットフォーム サポート。Office アドインは、Office on the web、Windows、Mac、および iPad で実行できます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p105">Cross-platform support. Office Add-ins run in Office on the web, Windows, Mac, and iPad.</span></span>

- <span data-ttu-id="1cf96-p106">一元展開と配布。管理者は、組織全体に Office アドインを一元的に展開できます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p106">Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.</span></span>

- <span data-ttu-id="1cf96-p107">AppSource を経由した簡単なアクセス。AppSource に提出することで、広範な対象ユーザーにソリューションを公開できます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p107">Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.</span></span>

- <span data-ttu-id="1cf96-p108">標準の Web テクノロジに基づいている。任意のライブラリを使用して、Office アドインを構築することができます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p108">Based on standard web technology. You can use any library you like to build Office Add-ins.</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="1cf96-128">Office アドインのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="1cf96-128">Components of an Office Add-in</span></span>

<span data-ttu-id="1cf96-p109">Office アドインには、2 つの基本的なコンポーネントが含まれています。XML マニフェスト ファイルと独自の Web アプリケーションです。マニフェストは、アドインを Office クライアントと統合する方法など、さまざまな設定を定義します。Web アプリケーションは Web サーバーか、Microsoft Azure などの Web ホスティング サービスでホストされる必要があります。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p109">An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

### <a name="manifest"></a><span data-ttu-id="1cf96-132">マニフェスト</span><span class="sxs-lookup"><span data-stu-id="1cf96-132">Manifest</span></span>

<span data-ttu-id="1cf96-133">マニフェストは、次のようなアドインの設定と機能を指定する XML ファイルです。</span><span class="sxs-lookup"><span data-stu-id="1cf96-133">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span>

- <span data-ttu-id="1cf96-134">アドインの表示名、説明、ID、バージョン、および既定のロケール。</span><span class="sxs-lookup"><span data-stu-id="1cf96-134">The add-in's display name, description, ID, version, and default locale.</span></span>

- <span data-ttu-id="1cf96-135">Office とアドインを統合する方法。</span><span class="sxs-lookup"><span data-stu-id="1cf96-135">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="1cf96-136">アドインのアクセス許可レベルとデータ アクセスの要件。</span><span class="sxs-lookup"><span data-stu-id="1cf96-136">The permission level and data access requirements for the add-in.</span></span>

### <a name="web-app"></a><span data-ttu-id="1cf96-137">Web アプリケーション</span><span class="sxs-lookup"><span data-stu-id="1cf96-137">Web app</span></span>

<span data-ttu-id="1cf96-p110">最も基本的な Office アドインは、Office アプリケーション内に表示される静的な HTML ページで構成されますが、Office ドキュメントやその他のどんなインターネット リソースとも対話を行いません。ただし、Office ドキュメントと対話するエクスペリエンスを作成する、または、ユーザーが Office ホスト アプリケーションからオンライン リソースと対話できるようにするには、ホスティング プロバイダーがサポートする任意のクライアント側とサーバー側のテクノロジ (ASP.NET、PHP、または Node.js など) を使用できます。Office クライアントとドキュメントとの対話を行うには、Office.js JavaScript API を使用します。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p110">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span>

<span data-ttu-id="1cf96-141">*図 2. Hello World Office アドインのコンポーネント*</span><span class="sxs-lookup"><span data-stu-id="1cf96-141">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Hello World アドインのコンポーネント](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="1cf96-143">Office クライアントの拡張と、Office クライアントとの対話</span><span class="sxs-lookup"><span data-stu-id="1cf96-143">Extending and interacting with Office clients</span></span>

<span data-ttu-id="1cf96-144">Office アドインは、Office ホスト アプリケーション内で次を実行できます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-144">Office Add-ins can do the following within an Office host application:</span></span>

-  <span data-ttu-id="1cf96-145">機能の拡張 (任意の Office アプリケーション)</span><span class="sxs-lookup"><span data-stu-id="1cf96-145">Extend functionality (any Office application)</span></span>

-  <span data-ttu-id="1cf96-146">新しいオブジェクトの作成 (Excel または PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="1cf96-146">Create new objects (Excel or PowerPoint)</span></span>
 
### <a name="extend-office-functionality"></a><span data-ttu-id="1cf96-147">Office 機能の拡張</span><span class="sxs-lookup"><span data-stu-id="1cf96-147">Extend Office functionality</span></span>

<span data-ttu-id="1cf96-148">次の方法で、Office アプリケーションに新しい機能を追加できます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-148">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="1cf96-149">カスタム リボン ボタンとメニュー コマンド ("アドイン コマンド" と総称されます)</span><span class="sxs-lookup"><span data-stu-id="1cf96-149">Custom ribbon buttons and menu commands (collectively called “add-in commands”)</span></span>

-  <span data-ttu-id="1cf96-150">挿入可能な作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1cf96-150">Insertable task panes</span></span>

<span data-ttu-id="1cf96-151">カスタムの UI と作業ウィンドウは、アドイン マニフェストで指定されます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-151">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="1cf96-152">カスタム ボタンとメニュー コマンド</span><span class="sxs-lookup"><span data-stu-id="1cf96-152">Custom buttons and menu commands</span></span>  

<span data-ttu-id="1cf96-p111">カスタムのリボン ボタンとメニュー項目は、Office on the web および Windows のリボンに追加できます。これにより、ユーザーは、Office アプリケーションからアドインに簡単に直接アクセスできます。コマンド ボタンでは、カスタム HTML による作業ウィンドウを表示したり、JavaScript 関数を実行したりするなど、さまざまなアクションを起動できます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p111">You can add custom ribbon buttons and menu items to the ribbon in Office on the web and Windows. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="1cf96-156">*図 3. リボンにあるアドイン コマンド*</span><span class="sxs-lookup"><span data-stu-id="1cf96-156">*Figure 3. Add-in commands in the ribbon*</span></span>

![カスタム ボタンとメニュー コマンド](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a><span data-ttu-id="1cf96-158">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1cf96-158">Task panes</span></span>  

<span data-ttu-id="1cf96-p112">ユーザーはアドイン コマンドに加えて、作業ウィンドウを使用してソリューションの対話式操作を実行できます。アドイン コマンド (Office 2013 および Office on iPad) をサポートしていないクライアントは、アドインを作業ウィンドウとして実行します。ユーザーは **[挿入]** タブの **[アドイン]** ボタンを使用して、作業ウィンドウのアドインを起動します。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p112">You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span>

<span data-ttu-id="1cf96-162">*図 4. 作業ウィンドウ*</span><span class="sxs-lookup"><span data-stu-id="1cf96-162">*Figure 4. Task pane*</span></span>

![作業ウィンドウとアドイン コマンドを使用する](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="1cf96-164">Outlook の機能を拡張する</span><span class="sxs-lookup"><span data-stu-id="1cf96-164">Extend Outlook functionality</span></span>

<span data-ttu-id="1cf96-p113">Outlook アドインは Office のリボンを拡張したり、コンテキストに応じて表示または作成時に Outlook アイテムの隣に表示したりすることもできます。ユーザーが受信した項目を表示するか、返信または新しい項目を作成している場合には、電子メールメッセージ、会議出席依頼、会議の返信、会議の取り消し、または予定を操作できます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p113">Outlook add-ins can extend the Office ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="1cf96-p114">Outlook アドインでは、アイテムのコンテキスト情報 (アドレスや追跡 ID など) にアクセスし、そのデータを使用してサーバー上または Web サービスの追加情報にアクセスして、魅力的なユーザー エクスペリエンスを作成できます。ほとんどの場合、Outlook アドインはデスクトップ、Web、タブレットおよびモバイル デバイスに変更なしでシームレスなエクスペリエンスを提供する Outlook ホスト アプリケーションで実行できます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p114">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification in the Outlook host application to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span>

<span data-ttu-id="1cf96-169">Outlook アドインの概要については、「[Outlook アドインの概要](../outlook/outlook-add-ins-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1cf96-169">For an overview of Outlook add-ins, see [Outlook add-ins overview](../outlook/outlook-add-ins-overview.md).</span></span>

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="1cf96-170">Office ドキュメント内に新しいオブジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="1cf96-170">Create new objects in Office documents</span></span>

<span data-ttu-id="1cf96-p115">Excel および PowerPoint のドキュメント内に、コンテンツ アドインと呼ばれる Web ベースのオブジェクトを埋め込むことができます。コンテンツ アドインにより、ユーザーは充実した Web ベースのデータの可視化、埋め込まれたメディア (YouTube ビデオ プレーヤーや画像ギャラリーなど)、およびその他の外部コンテンツを統合できます。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p115">You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="1cf96-173">*図 5. コンテンツ アドイン*</span><span class="sxs-lookup"><span data-stu-id="1cf96-173">*Figure 5. Content add-in*</span></span>

![コンテンツ アドインと呼ばれる Web ベースのオブジェクトを埋め込む](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="1cf96-175">Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="1cf96-175">Office JavaScript APIs</span></span>

<span data-ttu-id="1cf96-p116">Office JavaScript API には、アドインを構築したり、Office のコンテンツおよび Web サービスと対話したりするためのオブジェクトとメンバーが含まれています。Excel、Outlook、Word、PowerPoint、OneNote、Project には、共通のオブジェクト モデルがあり、共有されています。Excel および Word には、さらに多くのホスト固有のオブジェクト モデルが用意されています。これらの API では、特定のホストのアドイン作成を容易にする段落やブックなど、既知のオブジェクトへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1cf96-p116">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services. There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project. There are also more extensive host-specific object models for Excel and Word. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span></span>  

## <a name="next-steps"></a><span data-ttu-id="1cf96-180">次のステップ</span><span class="sxs-lookup"><span data-stu-id="1cf96-180">Next steps</span></span>

<span data-ttu-id="1cf96-181">Office アドインの開発の詳細については、「[Office アドインを構築する](../overview/office-add-ins-fundamentals.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1cf96-181">For a more detailed introduction to developing Office Add-ins, see [Building Office Add-ins](../overview/office-add-ins-fundamentals.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="1cf96-182">関連項目</span><span class="sxs-lookup"><span data-stu-id="1cf96-182">See also</span></span>

- [<span data-ttu-id="1cf96-183">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="1cf96-183">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="1cf96-184">Office アドインの中心概念</span><span class="sxs-lookup"><span data-stu-id="1cf96-184">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="1cf96-185">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="1cf96-185">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="1cf96-186">Office アドインを設計する</span><span class="sxs-lookup"><span data-stu-id="1cf96-186">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="1cf96-187">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="1cf96-187">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="1cf96-188">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="1cf96-188">Publish Office Add-ins</span></span>](../publish/publish.md)
