---
title: Office アドイン プラットフォームの概要
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: f7b1f4add776f1971e9762c5cb80dabed45b0a1c
ms.sourcegitcommit: a0e0416289b293863b8b4d3f9a12581a9e681b27
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2018
ms.locfileid: "20023166"
---
# <a name="office-add-ins-platform-overview"></a>Office アドイン プラットフォームの概要

Office アドインのプラットフォームを使用すると、Office アプリケーションを拡張し、Office ドキュメント内のコンテンツと対話するソリューションを構築できます。Office アドインで、HTML、CSS、および JavaScript などの一般的な Web テクノロジを使用し、Word、Excel、PowerPoint、OneNote、Project、および Outlook を拡張して対話操作することができます。Office for Windows、Office Online、Office for Mac、および Office for iPad を含む複数のプラットフォームにわたって Office ソリューションを実行できます。

Office アドインでは、ブラウザー内で Web ページが実行できる操作のほとんどすべてを実行できます。Office アドイン プラットフォームを使用して、次のことができます。

-  **Office クライアントに新しい機能を追加する** - Office に外部データを取り込む、Office ドキュメントを自動化する、サード パーティの機能を Office クライアントで公開する、などがあります。たとえば、Microsoft Graph API を使用して、生産性の向上につながるデータに接続します。 
    
-  **Office ドキュメントに埋め込み可能な充実した対話型のオブジェクトを新しく作成する** - マップやグラフ、ユーザーが自分の Excel スプレッドシートや PowerPoint プレゼンテーションに追加できる対話型の視覚化などを埋め込みます。 
    
## <a name="how-are-office-add-ins-different-than-com-and-vsto-add-ins"></a>Office アドインが COM および VSTO アドインと異なる点 

COM または VSTO アドインは、Office for Windows 上でのみ実行する以前の Office 統合ソリューションです。COM アドインとは異なり、Office アドインにはユーザーのデバイスまたは Office クライアントで実行されるコードは含まれません。Office アドインの場合、ホスト アプリケーション (たとえば Excel) がアドインのマニフェストを読み取り、アドインのカスタム リボン ボタンと UI のメニュー コマンドをフックします。これは必要に応じて、サンド ボックスのブラウザーのコンテキストで実行されるアドインの JavaScript と HTML を読み込みます。 

Office アドインは、VBA、COM、または VSTO を使用して作成されたアドインと比較して、次のような利点があります。 

- クロスプラットフォーム サポート。Office アドインは、Windows 用、Mac 用、iOS 用の Office と、Office Online で実行できます。 

- シングル サインオン (SSO)。Office アドインは、ユーザーの Office 365 のアカウントと簡単に統合できます。 

- 一元展開と配布。管理者は、組織全体で Office アドインを一元的に展開できます。 

- AppSource を経由した簡単なアクセス。AppSource に提出することで、広範な対象ユーザーにソリューションを公開できます。 

- 標準の Web テクノロジに基づいている。任意のライブラリを使用して、Office アドインを構築することができます。 

## <a name="components-of-an-office-add-in"></a>Office アドインのコンポーネント 

Office アドインには、2 つの基本的なコンポーネントが含まれています。XML マニフェスト ファイルと独自の Web アプリケーションです。マニフェストは、アドインを Office クライアントと統合する方法など、さまざまな設定を定義します。Web アプリケーションは Web サーバーか、Microsoft Azure などの Web ホスティング サービスでホストされる必要があります。

*図1. アドイン マニフェスト (XML) + Web ページ (HTML、JS) = Office アドイン*

![マニフェストと Web ページで構成される Office アドイン](../images/about-addins-manifestwebpage.png)

### <a name="manifest"></a>マニフェスト 

マニフェストは、次のようなアドインの設定と機能を指定する XML ファイルです。 

- アドインの表示名、説明、ID、バージョン、および既定のロケール。 

- Office とアドインを統合する方法。  

- アドインのアクセス許可レベルとデータ アクセスの要件。 

### <a name="web-app"></a>Web アプリケーション 

最も基本的な Office アドインは、Office アプリケーション内に表示される静的な HTML ページで構成されますが、Office ドキュメントやその他のどんなインターネット リソースとも対話を行いません。ただし、Office ドキュメントと対話するエクスペリエンスを作成する、または、ユーザーが Office ホスト アプリケーションからオンライン リソースと対話できるようにするには、ホスティング プロバイダーがサポートする任意のクライアント側とサーバー側のテクノロジ (ASP.NET、PHP、または Node.js など) を使用できます。Office クライアントとドキュメントとの対話を行うには、Office.js JavaScript API を使用します。 

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

デスクトップ版 Office for Windows と Office Online のリボンにカスタム リボン ボタンおよびメニュー項目を追加できます。これにより、ユーザーは、Office アプリケーションから直接アドインに簡単にアクセスできます。コマンド ボタンは、カスタム HTML を使用して作業ウィンドウを表示したり、JavaScript 関数を実行したりするなど、さまざまなアクションを起動できます。  

*図3. リボンのアドイン コマンド*

![カスタム ボタンとメニュー コマンド](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a>作業ウィンドウ  

ユーザーがソリューションと対話できるようにするために、アドイン コマンドに加えて、作業ウィンドウを使用できます。アドイン コマンド (Office 2013 および Office for iPad) をサポートしていないクライアントは、アドインを作業ウィンドウとして実行します。ユーザーは **[挿入]** タブの **[アドイン]** ボタンを使用して、作業ウィンドウのアドインを起動します。 

*図 4. 作業ウィンドウ*

![作業ウィンドウ](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a>Outlook の機能を拡張する 

Outlook アドインは Office のリボンを拡張したり、コンテキストに応じて表示または作成時に Outlook アイテムの隣に表示したりすることもできます。ユーザーが受信した項目を表示するか、返信または新しい項目を作成している場合には、電子メールメッセージ、会議出席依頼、会議の返信、会議の取り消し、または予定を操作できます。 

Outlook アドインでは、アドレスや追跡 ID などのアイテムからコンテキスト情報にアクセスし、そのデータを使用してサーバー上の追加情報や Web サービスから魅力的なユーザー エクスペリエンスを作成することができます。Outlook アドインはほとんどの場合、Outlook、Outlook for Mac、Outlook Web App、デバイス用 Outlook Web App などのさまざまなサポートしているホスト アプリケーションで変更なしで実行でき、デスクトップ、Web、およびタブレットとモバイル デバイスでシームレスな操作を提供します。 

Outlook アドインの概要については、「[Outlook アドインの概要](https://docs.microsoft.com/en-us/outlook/add-ins/)」を参照してください。 

### <a name="create-new-objects-in-office-documents"></a>Office ドキュメント内に新しいオブジェクトを作成する 

Excel および PowerPoint のドキュメント内に、コンテンツ アドインと呼ばれる Web ベースのオブジェクトを埋め込むことができます。コンテンツ アドインにより、ユーザーは充実した Web ベースのデータの可視化、埋め込まれたメディア (YouTube ビデオ プレーヤーや画像ギャラリーなど)、およびその他の外部コンテンツを統合できます。

*図 5. コンテンツ アドイン*

![コンテンツ アドイン](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a>Office JavaScript API 

Office JavaScript API には、アドインを構築したり、Office のコンテンツおよび Web サービスと対話したりするためのオブジェクトとメンバーが含まれています。Excel、Outlook、Word、PowerPoint、OneNote、Project には、共通のオブジェクト モデルがあり、共有されています。Excel および Word には、さらに多くのホスト固有のオブジェクト モデルが用意されています。これらの API では、特定のホストのアドイン作成を容易にする段落やブックなど、既知のオブジェクトへのアクセスを提供します。  

## <a name="next-steps"></a>次の手順 

Officeアドインの構築を開始する詳細な方法は、[5 分間クイック スタート](https://docs.microsoft.com/en-us/office/dev/add-ins/)を参照してください。 Visual Studio やその他のエディタを使用してすぐにアドインを作成することができます。 

効果的で魅力的なユーザー エクスペリエンスを生み出すソリューションの計画を開始するには、Office アドインの[設計ガイドライン](../design/add-in-design.md)と[ベスト プラクティス](../concepts/add-in-development-best-practices.md)の理解を深めてください。    
   
## <a name="see-also"></a>関連項目

- [Office アドインのサンプル](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples)
- [JavaScript API for Office について](../develop/understanding-the-javascript-api-for-office.md)
- [Office アドインのホストとプラットフォームの可用性](../overview/office-add-in-availability.md)


    
