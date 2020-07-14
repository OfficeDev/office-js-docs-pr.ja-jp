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
# <a name="office-add-ins-platform-overview"></a>Office アドイン プラットフォームの概要

You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook. Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.

![Office アドインの拡張性の画像](../images/addins-overview.png)

Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:

-  **Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.

-  **Office ドキュメントに埋め込み可能な充実した対話型のオブジェクトを新しく作成する** - マップやグラフ、ユーザーが自分の Excel スプレッドシートや PowerPoint プレゼンテーションに追加できる対話型の視覚化などを埋め込みます。

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a>Office アドインが COM アドインおよび VSTO アドインと異なる点

COM or VSTO add-ins are earlier Office integration solutions that run only on Office on Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.

![Office アドインを使用する理由の画像](../images/why.png)

Office アドインは、VBA、COM、または VSTO を使用して作成されたアドインと比較して、次のような利点があります。

- Cross-platform support. Office Add-ins run in Office on the web, Windows, Mac, and iPad.

- Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.

- Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.

- Based on standard web technology. You can use any library you like to build Office Add-ins.

## <a name="components-of-an-office-add-in"></a>Office アドインのコンポーネント

An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.

### <a name="manifest"></a>マニフェスト

マニフェストは、次のようなアドインの設定と機能を指定する XML ファイルです。

- アドインの表示名、説明、ID、バージョン、および既定のロケール。

- Office とアドインを統合する方法。  

- アドインのアクセス許可レベルとデータ アクセスの要件。

### <a name="web-app"></a>Web アプリケーション

The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.

*図 2. Hello World Office アドインのコンポーネント*

![Hello World アドインのコンポーネント](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a>Office クライアントの拡張と、Office クライアントとの対話

Office アドインは、Office ホスト アプリケーション内で次を実行できます。

-  機能の拡張 (任意の Office アプリケーション)

-  新しいオブジェクトの作成 (Excel または PowerPoint)
 
### <a name="extend-office-functionality"></a>Office 機能の拡張

次の方法で、Office アプリケーションに新しい機能を追加できます。  

-  カスタム リボン ボタンとメニュー コマンド ("アドイン コマンド" と総称されます)

-  挿入可能な作業ウィンドウ

カスタムの UI と作業ウィンドウは、アドイン マニフェストで指定されます。  

#### <a name="custom-buttons-and-menu-commands"></a>カスタム ボタンとメニュー コマンド  

You can add custom ribbon buttons and menu items to the ribbon in Office on the web and Windows. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.  

*図 3. リボンにあるアドイン コマンド*

![カスタム ボタンとメニュー コマンド](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a>作業ウィンドウ  

You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.

*図 4. 作業ウィンドウ*

![作業ウィンドウとアドイン コマンドを使用する](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a>Outlook の機能を拡張する

Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item. 

Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification in the Outlook host application to provide a seamless experience on the desktop, web, and tablet and mobile devices.

Outlook アドインの概要については、「[Outlook アドインの概要](../outlook/outlook-add-ins-overview.md)」を参照してください。

### <a name="create-new-objects-in-office-documents"></a>Office ドキュメント内に新しいオブジェクトを作成する

You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.

*図 5. コンテンツ アドイン*

![コンテンツ アドインと呼ばれる Web ベースのオブジェクトを埋め込む](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a>Office JavaScript API

The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services. There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project. There are also more extensive host-specific object models for Excel and Word. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.  

## <a name="next-steps"></a>次のステップ

Office アドインの開発の詳細については、「[Office アドインを構築する](../overview/office-add-ins-fundamentals.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインを構築する](../overview/office-add-ins-fundamentals.md)
- [Office アドインの中心概念](../overview/core-concepts-office-add-ins.md)
- [Office アドインを開発する](../develop/develop-overview.md)
- [Office アドインを設計する](../design/add-in-design.md)
- [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
- [Office アドインを発行する](../publish/publish.md)
