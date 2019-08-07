---
title: Excel アドインの概要
description: ''
ms.date: 07/05/2019
localization_priority: Priority
ms.openlocfilehash: fbb0f69e7c32776fdd0bce6e5c10f39c562a5cbe
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575570"
---
# <a name="excel-add-ins-overview"></a><span data-ttu-id="c06a6-102">Excel アドインの概要</span><span class="sxs-lookup"><span data-stu-id="c06a6-102">Excel add-ins overview</span></span>

<span data-ttu-id="c06a6-p101">Excel アドインを使用すると、Windows、Mac、iPad、ブラウザーなどの複数のプラットフォームで Excel アプリケーションの機能を拡張できます。ブック内で Excel アドインを使用すると、次の操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-p101">An Excel add-in allows you to extend Excel application functionality across multiple platforms including Office on Windows, Office Online, Office for Mac, and Office for iPad. Use Excel add-ins within a workbook to:</span></span>

- <span data-ttu-id="c06a6-105">Excel オブジェクトを操作して Excel データを読み書きします。</span><span class="sxs-lookup"><span data-stu-id="c06a6-105">Interact with Excel objects, read and write Excel data.</span></span>
- <span data-ttu-id="c06a6-106">Web ベースの作業ウィンドウまたはコンテンツ ウィンドウを使用して機能を拡張します</span><span class="sxs-lookup"><span data-stu-id="c06a6-106">Extend functionality using web based task pane or content pane</span></span>
- <span data-ttu-id="c06a6-107">カスタム リボン ボタンやコンテキスト メニューの項目を追加します</span><span class="sxs-lookup"><span data-stu-id="c06a6-107">Add custom ribbon buttons or contextual menu items</span></span>
- <span data-ttu-id="c06a6-108">カスタム関数を追加します</span><span class="sxs-lookup"><span data-stu-id="c06a6-108">Add custom functions</span></span>
- <span data-ttu-id="c06a6-109">ダイアログ ウィンドウを使用して充実した操作を提供します</span><span class="sxs-lookup"><span data-stu-id="c06a6-109">Provide richer interaction using dialog window</span></span>

<span data-ttu-id="c06a6-110">Office アドインのプラットフォームには、Excel アドインの作成と実行を可能にするフレームワークと Office.js JavaScript API が用意されています。Office アドインのプラットフォームを使用した Excel アドインの作成には、次の利点があります。</span><span class="sxs-lookup"><span data-stu-id="c06a6-110">The Office Add-ins platform provides the framework and Office.js JavaScript APIs that enable you to create and run Excel add-ins. By using the Office Add-ins platform to create your Excel add-in, you'll get the following benefits:</span></span>

* <span data-ttu-id="c06a6-111">**クロスプラットフォーム サポート**: Excel アドインは、Office on the web、Windows、Mac、iPad で実行できます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-111">**Cross-platform support**: Excel add-ins run in Office on Windows, Mac, iOS, and Office Online.</span></span>
* <span data-ttu-id="c06a6-112">**一元展開**: 管理者は、組織全体のユーザーに Excel アドインをすばやく簡単に展開できます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-112">**Centralized deployment**: Admins can quickly and easily deploy Excel add-ins to users throughout an organization.</span></span>
* <span data-ttu-id="c06a6-113">**標準の Web テクノロジの使用**: HTML、CSS、JavaScript などの一般的な Web テクノロジを使用する Excel アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="c06a6-113">**Use of standard web technology**: Create your Excel add-in using familiar web technologies such as HTML, CSS, and JavaScript.</span></span>
* <span data-ttu-id="c06a6-114">**AppSource を経由した配布**: Excel アドインを [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d) に公開することで、幅広いユーザーと共有します。</span><span class="sxs-lookup"><span data-stu-id="c06a6-114">**Distribution via AppSource**: Share your Excel add-in with a broad audience by publishing it to [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d).</span></span>

> [!NOTE]
> <span data-ttu-id="c06a6-p102">Excel アドインは、Windows での Office でのみ実行する、以前の Office 統合ソリューションである COM アドインや VSTO アドインとは異なります。COM アドインとは異なり、Excel アドインではユーザーのデバイスや Excel 内にコードをインストールする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="c06a6-p102">Excel add-ins are different from COM and VSTO add-ins, which are earlier Office integration solutions that run only in Office on Windows. Unlike COM add-ins, Excel add-ins do not require you to install any code on a user's device, or within Excel.</span></span>

## <a name="components-of-an-excel-add-in"></a><span data-ttu-id="c06a6-117">Excel アドインのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="c06a6-117">Components of an Excel add-in</span></span>

<span data-ttu-id="c06a6-118">Excel アドインには 2 つの基本コンポーネントが含まれています。Web アプリケーションと、マニフェスト ファイルと呼ばれる構成ファイルです。</span><span class="sxs-lookup"><span data-stu-id="c06a6-118">An Excel add-in includes two basic components: a web application and a configuration file, called a manifest file.</span></span> 

<span data-ttu-id="c06a6-p103">Web アプリケーションは、[JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office) を使用して Excel のオブジェクトを操作します。また、オンライン リソースとの相互操作を簡単にすることもできます。 たとえば、アドインでは次の操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-p103">The web application uses the [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office) to interact with objects in Excel, and can also facilitate interaction with online resources. For example, an add-in can perform any of the following tasks:</span></span>

* <span data-ttu-id="c06a6-121">ブック内のデータ (ワークシート、範囲、表、グラフ、名前付きの項目など) を作成、読み込み、更新、および削除します。</span><span class="sxs-lookup"><span data-stu-id="c06a6-121">Create, read, update, and delete data in the workbook (worksheets, ranges, tables, charts, named items, and more).</span></span>
* <span data-ttu-id="c06a6-122">標準の OAuth 2.0 のフローを使用して、オンライン サービスでユーザー認証を実行します。</span><span class="sxs-lookup"><span data-stu-id="c06a6-122">Perform user authorization with an online service by using the standard OAuth 2.0 flow.</span></span>
* <span data-ttu-id="c06a6-123">Microsoft Graph やその他の API に、API 要求を発行します。</span><span class="sxs-lookup"><span data-stu-id="c06a6-123">Issue API requests to Microsoft Graph or any other API.</span></span>

<span data-ttu-id="c06a6-124">Web アプリケーションは、任意の Web サーバー上でホストできます。また、クライアント側のフレームワーク (Angular、React、jQuery など) や、サーバー側のテクノロジ (ASP.NET、Node.js、PHP など) を使用して構築できます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-124">The web application can be hosted on any web server, and can be built using client-side frameworks (such as Angular, React, jQuery) or server-side technologies (such as ASP.NET, Node.js, PHP).</span></span>

<span data-ttu-id="c06a6-125">[マニフェスト](../develop/add-in-manifests.md)は XML 構成ファイルであり、次のような設定と機能を指定することによって、アドインと Office クライアントを統合する方法を定義します。</span><span class="sxs-lookup"><span data-stu-id="c06a6-125">The [manifest](../develop/add-in-manifests.md) is an XML configuration file that defines how the add-in integrates with Office clients by specifying settings and capabilities such as:</span></span>

* <span data-ttu-id="c06a6-126">アドインの Web アプリケーションの URL。</span><span class="sxs-lookup"><span data-stu-id="c06a6-126">The URL of the add-in's web application.</span></span>
* <span data-ttu-id="c06a6-127">アドインの表示名、説明、ID、バージョン、および既定のロケール。</span><span class="sxs-lookup"><span data-stu-id="c06a6-127">The add-in's display name, description, ID, version, and default locale.</span></span>
* <span data-ttu-id="c06a6-128">アドインと Excel を統合する方法。アドインが作成する任意のカスタム UI (リボンのボタン、コンテキスト メニューなど) の統合を含む。</span><span class="sxs-lookup"><span data-stu-id="c06a6-128">How the add-in integrates with Excel, including any custom UI that the add-in creates (ribbon buttons, context menus, and so on).</span></span>
* <span data-ttu-id="c06a6-129">ドキュメントの読み取り、書き込みなど、アドインに必要なアクセス許可。</span><span class="sxs-lookup"><span data-stu-id="c06a6-129">Permissions that the add-in requires, such as reading and writing to the document.</span></span>

<span data-ttu-id="c06a6-130">エンドユーザーが Excel アドインをインストールして使用できるようにするには、そのマニフェストを AppSource かアドイン カタログに公開する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c06a6-130">To enable end users to install and use an Excel add-in, you must publish its manifest either to AppSource or to an add-ins catalog.</span></span> 

## <a name="capabilities-of-an-excel-add-in"></a><span data-ttu-id="c06a6-131">Excel アドインの機能</span><span class="sxs-lookup"><span data-stu-id="c06a6-131">Capabilities of an Excel add-in</span></span>

<span data-ttu-id="c06a6-132">ブック内のコンテンツの操作の他に、Excel アドインでは、カスタム リボンのボタンやメニュー コマンドを追加したり、作業ウィンドウを挿入したり、カスタム関数を追加したり、ダイアログ ボックスを開いたり、グラフや対話型のビジュアル化などの豊富な Web ベースのオブジェクトをワークシート内に埋め込むことができます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-132">In addition to interacting with the content in the workbook, Excel add-ins can add custom ribbon buttons or menu commands, insert task panes, open dialog boxes, and even embed rich, web-based objects such as charts or interactive visualizations within a worksheet.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="c06a6-133">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="c06a6-133">Add-in commands</span></span>

<span data-ttu-id="c06a6-p104">アドイン コマンドは、Excel UI を拡張する UI 要素であり、アドインのアクションを開始します。アドイン コマンドを使って、Excel のリボンにボタンを追加したり、コンテキスト メニューに項目を追加したりできます。ユーザーがアドイン コマンドを選択するときは、JavaScript コードの実行や、作業ウィンドウでのアドインのページの表示といったアクションを開始します。</span><span class="sxs-lookup"><span data-stu-id="c06a6-p104">Add-in commands are UI elements that extend the Excel UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu in Excel. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span></span> 

<span data-ttu-id="c06a6-137">**アドイン コマンド**</span><span class="sxs-lookup"><span data-stu-id="c06a6-137">**Add-in commands**</span></span>

![Excel のアドイン コマンド](../images/excel-add-in-commands-script-lab.png)

<span data-ttu-id="c06a6-139">コマンドの機能、サポートされているプラットフォーム、およびアドイン コマンド開発のベスト プラクティスについては、「[Excel、Word、および PowerPoint のアドイン コマンド](../design/add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c06a6-139">For more information about command capabilities, supported platforms, and best practices for developing add-in commands, see [Add-in commands for Excel, Word, and PowerPoint](../design/add-in-commands.md).</span></span>

### <a name="task-panes"></a><span data-ttu-id="c06a6-140">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c06a6-140">Task panes</span></span>

<span data-ttu-id="c06a6-p105">作業ウィンドウは、通常 Excel 内のウィンドウの右側に表示されるインターフェイスのサーフェスです。作業ウィンドウにより、ユーザーはコードを実行して Excel ドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-p105">Task panes are interface surfaces that typically appear on the right side of the window within Excel. Task panes give users access to interface controls that run code to modify the Excel document or display data from a data source.</span></span> 

<span data-ttu-id="c06a6-143">**作業ウィンドウ**</span><span class="sxs-lookup"><span data-stu-id="c06a6-143">**Task pane**</span></span>

![Excel の作業ウィンドウ アドイン](../images/excel-add-in-task-pane-insights.png)

<span data-ttu-id="c06a6-145">作業ウィンドウの詳細については、「[Office アドインの作業ウィンドウ](../design/task-pane-add-ins.md)」を参照してください。Excel の作業ウィンドウを実装するサンプルについては、「[Excel アドインの JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c06a6-145">For more information about task panes, see [Task panes in Office Add-ins](../design/task-pane-add-ins.md). For a sample that implements a task pane in Excel, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends).</span></span>

### <a name="custom-functions"></a><span data-ttu-id="c06a6-146">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="c06a6-146">Custom Functions</span></span>

<span data-ttu-id="c06a6-147">開発者は、カスタム関数を使用して関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-147">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="c06a6-148">ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-148">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> 

<span data-ttu-id="c06a6-149">**カスタム関数**</span><span class="sxs-lookup"><span data-stu-id="c06a6-149">**Custom function.**</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="c06a6-150">カスタム関数の詳細については、「[Excel でカスタム関数を作成する](custom-functions-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c06a6-150">For information about how to create custom functions in Excel using the Excel JavaScript API, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

### <a name="dialog-boxes"></a><span data-ttu-id="c06a6-151">ダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="c06a6-151">Dialog boxes</span></span>

<span data-ttu-id="c06a6-p107">ダイアログ ボックスは、作業中の Excel アプリケーション ウィンドウの手前に浮動するサーフェスです。 ダイアログ ボックスは、作業ウィンドウに直接開くことができないサインイン ページの表示、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオのホストなどの作業に使用できます。 Excel アドインでダイアログ ボックスを開くには、[ダイアログ API](/javascript/api/office/office.ui) を使用します。</span><span class="sxs-lookup"><span data-stu-id="c06a6-p107">Dialog boxes are surfaces that float above the active Excel application window. You can use dialog boxes for tasks such as displaying sign-in pages that can't be opened directly in a task pane, requesting that the user confirm an action, or hosting videos that might be too small if confined to a task pane. To open dialog boxes in your Excel add-in, use the [Dialog API](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="c06a6-155">**ダイアログ ボックス**</span><span class="sxs-lookup"><span data-stu-id="c06a6-155">**Dialog box**</span></span>

![Excel のアドイン ダイアログ ボックス](../images/excel-add-in-dialog-choose-number.png)

<span data-ttu-id="c06a6-157">ダイアログ ボックスとダイアログ API の詳細については、「[Office アドインのダイアログ ボックス](../design/dialog-boxes.md)」と「[Office アドインでダイアログ API を使用する](../develop/dialog-api-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c06a6-157">For more information about dialog boxes and the Dialog API, see [Dialog boxes in Office Add-ins](../design/dialog-boxes.md) and [Use the Dialog API in your Office Add-ins](../develop/dialog-api-in-office-add-ins.md).</span></span>

### <a name="content-add-ins"></a><span data-ttu-id="c06a6-158">コンテンツ アドイン</span><span class="sxs-lookup"><span data-stu-id="c06a6-158">Content add-ins</span></span>

<span data-ttu-id="c06a6-p108">コンテンツ アドインは、Excel ドキュメントに直接埋め込むことができるサーフェスです。 コンテンツ アドインを使用すると、グラフ、データのビジュアル化、メディアなど豊富な Web ベース オブジェクトをワークシートに埋め込んだり、Excel ドキュメントの変更またはデータ ソースのデータの表示のためのコードを実行するインターフェイス コントロールへのアクセスをユーザーに提供したりできます。 機能を直接ドキュメントに埋め込む場合は、コンテンツ アドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="c06a6-p108">Content add-ins are surfaces that you can embed directly into Excel documents. You can use content add-ins to embed rich, web-based objects such as charts, data visualizations, or media into a worksheet or to give users access to interface controls that run code to modify the Excel document or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.</span></span>

<span data-ttu-id="c06a6-162">**コンテンツ アドイン**</span><span class="sxs-lookup"><span data-stu-id="c06a6-162">**Content add-in**</span></span>

![Excel のコンテンツ アドイン](../images/excel-add-in-content-map.png)

<span data-ttu-id="c06a6-164">コンテンツ アドインの詳細については、「[コンテンツ Office アドイン](../design/content-add-ins.md)」を参照してください。Excel のコンテンツ アドインの実装サンプルについては、GitHub の「[Excel コンテンツ アドイン Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c06a6-164">For more information about content add-ins, see [Content Office Add-ins](../design/content-add-ins.md). For a sample that implements a content add-in in Excel, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="javascript-apis-to-interact-with-workbook-content"></a><span data-ttu-id="c06a6-165">ブックのコンテンツを操作する JavaScript API</span><span class="sxs-lookup"><span data-stu-id="c06a6-165">JavaScript APIs to interact with workbook content</span></span>

<span data-ttu-id="c06a6-166">Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office) を使用して、Excel のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="c06a6-166">An Excel add-in interacts with objects in Excel by using the [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office), which includes two JavaScript object models:</span></span>

* <span data-ttu-id="c06a6-167">**Excel JavaScript API**:Office 2016 で導入された [Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定された Excel オブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="c06a6-167">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) provides strongly-typed Excel objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="c06a6-168">**共通 API**: Office 2013 で導入された共通 API を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-168">**Common API**: Introduced with Office 2013, the Common API enables you to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span> <span data-ttu-id="c06a6-169">共通 API は Excel の操作に限られた機能を提供します。そのため、アドインを Excel 2013 で実行する必要がある場合に使用できます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-169">Because the Common API does provide limited functionality for Excel interaction, you can use it if your add-in needs to run on Excel 2013.</span></span>

## <a name="next-steps"></a><span data-ttu-id="c06a6-170">次の手順</span><span class="sxs-lookup"><span data-stu-id="c06a6-170">Next steps</span></span>

<span data-ttu-id="c06a6-171">[最初の Excel アドインを作成する](../quickstarts/excel-quickstart-jquery.md)ことから始めます。</span><span class="sxs-lookup"><span data-stu-id="c06a6-171">Get started by [creating your first Excel add-in](../quickstarts/excel-quickstart-jquery.md).</span></span> <span data-ttu-id="c06a6-172">次に、Excel アドイン構築の[中心概念](excel-add-ins-core-concepts.md)について説明します。</span><span class="sxs-lookup"><span data-stu-id="c06a6-172">Then, learn about the [core concepts](excel-add-ins-core-concepts.md) of building Excel add-ins.</span></span>

## <a name="see-also"></a><span data-ttu-id="c06a6-173">関連項目</span><span class="sxs-lookup"><span data-stu-id="c06a6-173">See also</span></span>

- [<span data-ttu-id="c06a6-174">Excel アドイン ドキュメント</span><span class="sxs-lookup"><span data-stu-id="c06a6-174">Excel add-ins documentation</span></span>](index.md)
- [<span data-ttu-id="c06a6-175">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="c06a6-175">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="c06a6-176">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="c06a6-176">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="c06a6-177">Office アドインの設計ガイドライン</span><span class="sxs-lookup"><span data-stu-id="c06a6-177">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="c06a6-178">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="c06a6-178">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c06a6-179">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="c06a6-179">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
