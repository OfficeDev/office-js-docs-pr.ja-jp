---
title: Excel アドインの概要
description: ''
ms.date: 01/23/2018
localization_priority: Priority
ms.openlocfilehash: 747b9b28f8e15de71a7af7e72bced61f8063139e
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386737"
---
# <a name="excel-add-ins-overview"></a><span data-ttu-id="cf79e-102">Excel アドインの概要</span><span class="sxs-lookup"><span data-stu-id="cf79e-102">Excel add-ins overview</span></span>

<span data-ttu-id="cf79e-p101">Excel アドインを使用すると、Office for Windows、Office Online、Office for the Mac、Office for the iPad など、複数のプラットフォームにわたって Excel アプリケーションの機能を拡張できます。ブック内で Excel アドインを使用すると、次の操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-p101">An Excel add-in allows you to extend Excel application functionality across multiple platforms including Office for Windows, Office Online, Office for the Mac, and Office for the iPad. Use Excel add-ins within a workbook to:</span></span>

- <span data-ttu-id="cf79e-105">Excel オブジェクトを操作して Excel データを読み書きします。</span><span class="sxs-lookup"><span data-stu-id="cf79e-105">Interact with Excel objects, read and write Excel data.</span></span> 
- <span data-ttu-id="cf79e-106">Web ベースの作業ウィンドウまたはコンテンツ ウィンドウを使用して機能を拡張します</span><span class="sxs-lookup"><span data-stu-id="cf79e-106">Extend functionality using web based task pane or content pane</span></span> 
- <span data-ttu-id="cf79e-107">カスタム リボン ボタンやコンテキスト メニューの項目を追加します</span><span class="sxs-lookup"><span data-stu-id="cf79e-107">Add custom ribbon buttons or contextual menu items</span></span>
- <span data-ttu-id="cf79e-108">ダイアログ ウィンドウを使用して充実した操作を提供します</span><span class="sxs-lookup"><span data-stu-id="cf79e-108">Provide richer interaction using dialog window</span></span> 

<span data-ttu-id="cf79e-109">Office アドインのプラットフォームには、Excel アドインの作成と実行を可能にするフレームワークと Office.js JavaScript API が用意されています。Office アドインのプラットフォームを使用した Excel アドインの作成には、次の利点があります。</span><span class="sxs-lookup"><span data-stu-id="cf79e-109">The Office Add-ins platform provides the framework and Office.js JavaScript APIs that enable you to create and run Excel add-ins. By using the Office Add-ins platform to create your Excel add-in, you'll get the following benefits:</span></span>

* <span data-ttu-id="cf79e-110">**クロスプラットフォーム サポート**:Excel アドインは、Windows 版、Mac 版、iOS 版の Office と、Office Online で実行できます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-110">**Cross-platform support**: Excel add-ins run in Office for Windows, Mac, iOS, and Office Online.</span></span>
* <span data-ttu-id="cf79e-111">**一元展開**: 管理者は、組織全体のユーザーに Excel アドインをすばやく簡単に展開できます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-111">**Centralized deployment**: Admins can quickly and easily deploy Excel add-ins to users throughout an organization.</span></span>
* <span data-ttu-id="cf79e-112">**標準の Web テクノロジの使用**: HTML、CSS、JavaScript などの一般的な Web テクノロジを使用する Excel アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="cf79e-112">**Use of standard web technology**: Create your Excel add-in using familiar web technologies such as HTML, CSS, and JavaScript.</span></span>
* <span data-ttu-id="cf79e-113">**AppSource を経由した配布**: Excel アドインを [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d) に公開することで、幅広いユーザーと共有します。</span><span class="sxs-lookup"><span data-stu-id="cf79e-113">**Distribution via AppSource**: Share your Excel add-in with a broad audience by publishing it to [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d).</span></span>

> [!NOTE]
> <span data-ttu-id="cf79e-114">Excel アドインは、Office for Windows 上でのみ実行する以前の Office 統合ソリューションである COM アドインや VSTO アドインとは異なります。</span><span class="sxs-lookup"><span data-stu-id="cf79e-114">Excel add-ins are different from COM and VSTO add-ins, which are earlier Office integration solutions that run only on Office for Windows.</span></span> <span data-ttu-id="cf79e-115">COM アドインとは異なり、Excel アドインではユーザーのデバイスや Excel 内にコードをインストールする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="cf79e-115">Unlike COM add-ins, Excel add-ins do not require you to install any code on a user's device, or within Excel.</span></span> 

## <a name="components-of-an-excel-add-in"></a><span data-ttu-id="cf79e-116">Excel アドインのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="cf79e-116">Components of an Excel add-in</span></span> 

<span data-ttu-id="cf79e-117">Excel アドインには 2 つの基本コンポーネントが含まれています。Web アプリケーションと、マニフェスト ファイルと呼ばれる構成ファイルです。</span><span class="sxs-lookup"><span data-stu-id="cf79e-117">An Excel add-in includes two basic components: a web application and a configuration file, called a manifest file.</span></span> 

<span data-ttu-id="cf79e-118">Web アプリケーションは、[JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) を使用して Excel のオブジェクトを操作します。また、オンライン リソースとの相互操作を簡単にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-118">The web application uses the [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) to interact with objects in Excel, and can also facilitate interaction with online resources.</span></span> <span data-ttu-id="cf79e-119">たとえば、アドインでは次の操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-119">For example, an add-in can perform any of the following tasks:</span></span>

* <span data-ttu-id="cf79e-120">ブック内のデータ (ワークシート、範囲、表、グラフ、名前付きの項目など) を作成、読み込み、更新、および削除します。</span><span class="sxs-lookup"><span data-stu-id="cf79e-120">Create, read, update, and delete data in the workbook (worksheets, ranges, tables, charts, named items, and more).</span></span>
* <span data-ttu-id="cf79e-121">標準の OAuth 2.0 のフローを使用して、オンライン サービスでユーザー認証を実行します。</span><span class="sxs-lookup"><span data-stu-id="cf79e-121">Perform user authorization with an online service by using the standard OAuth 2.0 flow.</span></span>
* <span data-ttu-id="cf79e-122">Microsoft Graph やその他の API に、API 要求を発行します。</span><span class="sxs-lookup"><span data-stu-id="cf79e-122">Issue API requests to Microsoft Graph or any other API.</span></span>

<span data-ttu-id="cf79e-123">Web アプリケーションは、任意の Web サーバー上でホストできます。また、クライアント側のフレームワーク (Angular、React、jQuery など) や、サーバー側のテクノロジ (ASP.NET、Node.js、PHP など) を使用して構築できます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-123">The web application can be hosted on any web server, and can be built using client-side frameworks (such as Angular, React, jQuery) or server-side technologies (such as ASP.NET, Node.js, PHP).</span></span>

<span data-ttu-id="cf79e-124">[マニフェスト](../develop/add-in-manifests.md)は XML 構成ファイルであり、次のような設定と機能を指定することによって、アドインと Office クライアントを統合する方法を定義します。</span><span class="sxs-lookup"><span data-stu-id="cf79e-124">The [manifest](../develop/add-in-manifests.md) is an XML configuration file that defines how the add-in integrates with Office clients by specifying settings and capabilities such as:</span></span> 

* <span data-ttu-id="cf79e-125">アドインの Web アプリケーションの URL。</span><span class="sxs-lookup"><span data-stu-id="cf79e-125">The URL of the add-in's web application.</span></span>
* <span data-ttu-id="cf79e-126">アドインの表示名、説明、ID、バージョン、および既定のロケール。</span><span class="sxs-lookup"><span data-stu-id="cf79e-126">The add-in's display name, description, ID, version, and default locale.</span></span>
* <span data-ttu-id="cf79e-127">アドインと Excel を統合する方法。アドインが作成する任意のカスタム UI (リボンのボタン、コンテキスト メニューなど) の統合を含む。</span><span class="sxs-lookup"><span data-stu-id="cf79e-127">How the add-in integrates with Excel, including any custom UI that the add-in creates (ribbon buttons, context menus, and so on).</span></span>
* <span data-ttu-id="cf79e-128">ドキュメントの読み取り、書き込みなど、アドインに必要なアクセス許可。</span><span class="sxs-lookup"><span data-stu-id="cf79e-128">Permissions that the add-in requires, such as reading and writing to the document.</span></span>

<span data-ttu-id="cf79e-129">エンドユーザーが Excel アドインをインストールして使用できるようにするには、そのマニフェストを AppSource かアドイン カタログに公開する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf79e-129">To enable end users to install and use an Excel add-in, you must publish its manifest either to AppSource or to an add-ins catalog.</span></span> 

## <a name="capabilities-of-an-excel-add-in"></a><span data-ttu-id="cf79e-130">Excel アドインの機能</span><span class="sxs-lookup"><span data-stu-id="cf79e-130">Capabilities of an Excel add-in</span></span>

<span data-ttu-id="cf79e-131">ブック内のコンテンツの操作の他に、Excel アドインでは、カスタム リボンのボタンやメニュー コマンドを追加したり、作業ウィンドウを挿入したり、ダイアログ ボックスを開いたり、グラフや対話型のビジュアル化などの豊富な Web ベースのオブジェクトをワークシート内に埋め込むことができます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-131">In addition to interacting with the content in the workbook, Excel add-ins can add custom ribbon buttons or menu commands, insert task panes, open dialog boxes, and even embed rich, web-based objects such as charts or interactive visualizations within a worksheet.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="cf79e-132">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="cf79e-132">Add-in commands</span></span>

<span data-ttu-id="cf79e-p104">アドイン コマンドは、Excel UI を拡張する UI 要素であり、アドインのアクションを開始します。アドイン コマンドを使って、Excel のリボンにボタンを追加したり、コンテキスト メニューに項目を追加したりできます。ユーザーがアドイン コマンドを選択するときは、JavaScript コードの実行や、作業ウィンドウでのアドインのページの表示といったアクションを開始します。</span><span class="sxs-lookup"><span data-stu-id="cf79e-p104">Add-in commands are UI elements that extend the Excel UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu in Excel. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span></span> 

<span data-ttu-id="cf79e-136">**アドイン コマンド**</span><span class="sxs-lookup"><span data-stu-id="cf79e-136">**Add-in commands**</span></span>

![Excel のアドイン コマンド](../images/excel-add-in-commands-script-lab.png)

<span data-ttu-id="cf79e-138">コマンドの機能、サポートされているプラットフォーム、およびアドイン コマンド開発のベスト プラクティスについては、「[Excel、Word、および PowerPoint のアドイン コマンド](../design/add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf79e-138">For more information about command capabilities, supported platforms, and best practices for developing add-in commands, see [Add-in commands for Excel, Word, and PowerPoint](../design/add-in-commands.md).</span></span>

### <a name="task-panes"></a><span data-ttu-id="cf79e-139">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="cf79e-139">Task panes</span></span>

<span data-ttu-id="cf79e-p105">作業ウィンドウは、通常 Excel 内のウィンドウの右側に表示されるインターフェイスのサーフェスです。作業ウィンドウにより、ユーザーはコードを実行して Excel ドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-p105">Task panes are interface surfaces that typically appear on the right side of the window within Excel. Task panes give users access to interface controls that run code to modify the Excel document or display data from a data source.</span></span> 

<span data-ttu-id="cf79e-142">**作業ウィンドウ**</span><span class="sxs-lookup"><span data-stu-id="cf79e-142">**Task pane**</span></span>

![Excel の作業ウィンドウ アドイン](../images/excel-add-in-task-pane-insights.png)

<span data-ttu-id="cf79e-144">作業ウィンドウの詳細については、「[Office アドインの作業ウィンドウ](../design/task-pane-add-ins.md)」を参照してください。Excel の作業ウィンドウを実装するサンプルについては、「[Excel アドインの JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf79e-144">For more information about task panes, see [Task panes in Office Add-ins](../design/task-pane-add-ins.md). For a sample that implements a task pane in Excel, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends).</span></span>

### <a name="dialog-boxes"></a><span data-ttu-id="cf79e-145">ダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="cf79e-145">Dialog boxes</span></span>

<span data-ttu-id="cf79e-146">ダイアログ ボックスは、作業中の Excel アプリケーション ウィンドウの手前に浮動するサーフェスです。</span><span class="sxs-lookup"><span data-stu-id="cf79e-146">Dialog boxes are surfaces that float above the active Excel application window.</span></span> <span data-ttu-id="cf79e-147">ダイアログ ボックスは、作業ウィンドウに直接開くことができないサインイン ページの表示、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオのホストなどの作業に使用できます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-147">You can use dialog boxes for tasks such as displaying sign-in pages that can't be opened directly in a task pane, requesting that the user confirm an action, or hosting videos that might be too small if confined to a task pane.</span></span> <span data-ttu-id="cf79e-148">Excel アドインでダイアログ ボックスを開くには、[ダイアログ API](https://docs.microsoft.com/javascript/api/office/office.ui) を使用します。</span><span class="sxs-lookup"><span data-stu-id="cf79e-148">To open dialog boxes in your Excel add-in, use the [Dialog API](https://docs.microsoft.com/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="cf79e-149">**ダイアログ ボックス**</span><span class="sxs-lookup"><span data-stu-id="cf79e-149">**Dialog box**</span></span>

![Excel のアドイン ダイアログ ボックス](../images/excel-add-in-dialog-choose-number.png)

<span data-ttu-id="cf79e-151">ダイアログ ボックスとダイアログ API の詳細については、「[Office アドインのダイアログ ボックス](../design/dialog-boxes.md)」と「[Office アドインでダイアログ API を使用する](../develop/dialog-api-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf79e-151">For more information about dialog boxes and the Dialog API, see [Dialog boxes in Office Add-ins](../design/dialog-boxes.md) and [Use the Dialog API in your Office Add-ins](../develop/dialog-api-in-office-add-ins.md).</span></span>

### <a name="content-add-ins"></a><span data-ttu-id="cf79e-152">コンテンツ アドイン</span><span class="sxs-lookup"><span data-stu-id="cf79e-152">Content add-ins</span></span>

<span data-ttu-id="cf79e-153">コンテンツ アドインは、Excel ドキュメントに直接埋め込むことができるサーフェスです。</span><span class="sxs-lookup"><span data-stu-id="cf79e-153">Content add-ins are surfaces that you can embed directly into Excel documents.</span></span> <span data-ttu-id="cf79e-154">コンテンツ アドインを使用すると、グラフ、データのビジュアル化、メディアなど豊富な Web ベース オブジェクトをワークシートに埋め込んだり、Excel ドキュメントの変更またはデータ ソースのデータの表示のためのコードを実行するインターフェイス コントロールへのアクセスをユーザーに提供したりできます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-154">You can use content add-ins to embed rich, web-based objects such as charts, data visualizations, or media into a worksheet or to give users access to interface controls that run code to modify the Excel document or display data from a data source.</span></span> <span data-ttu-id="cf79e-155">機能を直接ドキュメントに埋め込む場合は、コンテンツ アドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="cf79e-155">Use content add-ins when you want to embed functionality directly into the document.</span></span>

<span data-ttu-id="cf79e-156">**コンテンツ アドイン**</span><span class="sxs-lookup"><span data-stu-id="cf79e-156">**Content add-in**</span></span>

![Excel のコンテンツ アドイン](../images/excel-add-in-content-map.png)

<span data-ttu-id="cf79e-158">コンテンツ アドインの詳細については、「[コンテンツ Office アドイン](../design/content-add-ins.md)」を参照してください。Excel のコンテンツ アドインの実装サンプルについては、GitHub の「[Excel コンテンツ アドイン Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf79e-158">For more information about content add-ins, see [Content Office Add-ins](../design/content-add-ins.md). For a sample that implements a content add-in in Excel, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="javascript-apis-to-interact-with-workbook-content"></a><span data-ttu-id="cf79e-159">ブックのコンテンツを操作する JavaScript API</span><span class="sxs-lookup"><span data-stu-id="cf79e-159">JavaScript APIs to interact with workbook content</span></span>

<span data-ttu-id="cf79e-160">Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) を使用して、Excel のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="cf79e-160">An Excel add-in interacts with objects in Excel by using the [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office), which includes two JavaScript object models:</span></span>

* <span data-ttu-id="cf79e-161">**Excel JavaScript API**:Office 2016 で導入された [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定された Excel オブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="cf79e-161">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) provides strongly-typed Excel objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="cf79e-162">**共通 API**: Office 2013 で導入された共通 API を使用すると、Word、Excel、PowerPoint など複数の種類のホスト アプリケーションに共通する UI、ダイアログ、クライアント設定などの機能にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-162">**Common API**: Introduced with Office 2013, the Common API enables you to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span> <span data-ttu-id="cf79e-163">共通 API は Excel の操作に限られた機能を提供します。そのため、アドインを Excel 2013 で実行する必要がある場合に使用できます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-163">Because the Common API does provide limited functionality for Excel interaction, you can use it if your add-in needs to run on Excel 2013.</span></span>

## <a name="next-steps"></a><span data-ttu-id="cf79e-164">次の手順</span><span class="sxs-lookup"><span data-stu-id="cf79e-164">Next steps</span></span>

<span data-ttu-id="cf79e-165">[最初の Excel アドインを作成する](excel-add-ins-get-started-overview.md)ことから始めます。</span><span class="sxs-lookup"><span data-stu-id="cf79e-165">Get started by [creating your first Excel add-in](excel-add-ins-get-started-overview.md).</span></span> <span data-ttu-id="cf79e-166">次に、Excel アドイン構築の[中心概念](excel-add-ins-core-concepts.md)について説明します。</span><span class="sxs-lookup"><span data-stu-id="cf79e-166">Then, learn about the [core concepts](excel-add-ins-core-concepts.md) of building Excel add-ins.</span></span>

## <a name="see-also"></a><span data-ttu-id="cf79e-167">関連項目</span><span class="sxs-lookup"><span data-stu-id="cf79e-167">See also</span></span>

- [<span data-ttu-id="cf79e-168">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="cf79e-168">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="cf79e-169">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="cf79e-169">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="cf79e-170">Office アドインの設計ガイドライン</span><span class="sxs-lookup"><span data-stu-id="cf79e-170">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="cf79e-171">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="cf79e-171">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="cf79e-172">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="cf79e-172">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
