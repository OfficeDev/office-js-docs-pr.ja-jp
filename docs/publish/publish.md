---
title: Office アドインを展開し、発行する
description: テスト目的またはユーザーに配布する目的で Office アドインを展開するための方法とオプション。
ms.date: 06/02/2020
localization_priority: Priority
ms.openlocfilehash: 8a3de7ae6f507ac21dce89d13417e87d5c89a428
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839685"
---
# <a name="deploy-and-publish-office-add-ins"></a>Office アドインを展開し、発行する

さまざまな方法を利用し、テスト目的またはユーザーに配布する目的で、Office アドインを展開できます。

|**メソッド**|**Use...**|
|:---------|:------------|
|[サイドロード](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)|開発プロセスの一環として、Windows、iPad、Mac、またはブラウザーで実行するアドインをテストします。 (製品版アドインではありません)。|
|[ネットワーク共有](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|開発プロセスの一環として、ローカルホスト以外のサーバーにアドインを公開した後、Windows で実行されているアドインをテストします。 (運用環境用のアドインや、iPad、Mac、Web でのテスト用ではありません。)|
|[一元展開](centralized-deployment.md)|クラウド環境で、Microsoft 365 管理センターを使用して組織内のユーザーにアドインを配布します。|
|[SharePoint カタログ](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|オンプレミス環境で、組織内のユーザーにアドインを配布します。|
|[AppSource](/office/dev/store/submit-to-appsource-via-partner-center)|ユーザーに配布する目的でアドインを公開します。|
|[Exchange サーバー](#outlook-add-in-deployment)|オンプレミス環境またはオンライン環境で、ユーザーに Outlook アドインを配布します。|

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="deployment-options-by-office-application-and-add-in-type"></a>Office アプリケーションとアドインの種類による展開オプション

選択可能な展開オプションは、対象の Office アプリケーションや作成するアドインの種類によって異なります。

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Word、Excel、PowerPoint のアドインの展開オプション

| 拡張点 | サイドロード | ネットワーク共有 | Microsoft 365 管理センター |AppSource   | SharePoint カタログ\* |
|:----------------|:-----------:|:-------------:|:-----------------------:|:----------:|:--------------------:|
| コンテンツ         | X           | X             | X                       | X          | X                    |
| 作業ウィンドウ       | X           | X             | X                       | X          | X                    |
| コマンド         | X           | X             | X                       | X          |                      |

&#42; SharePoint カタログは Office on Mac をサポートしません。

### <a name="deployment-options-for-outlook-add-ins"></a>Outlook アドインの展開オプション

| 拡張点 | サイドロード | Exchange サーバー | AppSource    |
|:----------------|:-----------:|:---------------:|:------------:|
| メール アプリ        | X           | X               | X            |
| コマンド         | X           | X               | X            |

## <a name="production-deployment-methods"></a>運用環境での展開方法

次からの各セクションでは、組織内のユーザーに運用環境の Office アドインを配布する際に最も一般的に使用される展開方法についての追加情報を示します。

エンド ユーザーがアドインを取得、挿入、実行する方法については、「[Office アドインの使用を開始する](https://support.office.com/article/start-using-your-office-add-in-82e665c4-6700-4b56-a3f3-ef5441996862)」を参照してください。

### <a name="centralized-deployment-via-the-microsoft-365-admin-center"></a>Microsoft 365 管理センターからの一元展開

Microsoft 365 管理センターを使用すると、管理者は組織内のユーザーとグループに Office アドインを簡単に展開できるようになります。 管理センター経由で展開されたアドインは、ユーザーがすぐに Office アプリケーションで利用できるようになります。クライアントの構成は必要ありません。 一元展開は、内部アドインの展開に使用することも、ISV が提供するアドインの展開に使用することもできます。

詳細については、「[Microsoft 365 管理センターからの一元展開を使用した Office アドインの発行](centralized-deployment.md)」を参照してください。

### <a name="sharepoint-app-catalog-deployment"></a>SharePoint アプリ カタログの展開

SharePoint アプリ カタログは、Word、Excel、PowerPoint のアドインをホストするために作成できる特別なサイト コレクションです。SharePoint カタログは、マニフェストの `VersionOverrides` ノードに実装されている新しいアドイン機能 (アドイン コマンドを含む) をサポートしていないため、可能な場合は管理センター経由の一元展開を実行することをお勧めします。SharePoint カタログによって展開したアドイン コマンドは、既定では作業ウィンドウで開かれます。

オンプレミス環境でアドインを展開する場合は、SharePoint カタログを使用します。詳細については、「[作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)」を参照してください。

> [!NOTE]
> SharePoint カタログは Office on Mac をサポートしません。 Office アドインを Mac クライアントに展開するには、そのアドインを [AppSource](/office/dev/store/submit-to-the-office-store) に提出する必要があります。

### <a name="outlook-add-in-deployment"></a>Outlook アドインの展開

Azure AD の ID サービスを使用しないオンプレミス環境およびオンライン環境では、Exchange サーバー経由で Outlook アドインを展開することができます。

Outlook アドインの展開には以下が必要です。

- Microsoft 365、Exchange Online、または Exchange Server 2013 以降
- Outlook 2013 以降

アドインをテナントに割り当てるには、Exchange 管理センターを使用して、ファイルまたは URL から直接マニフェストをアップロードするか、または AppSource からアドインを追加します。アドインを個々のユーザーに割り当てるには、Exchange PowerShell を使用する必要があります。詳細については、TechNet の「[組織の Outlook アドインをインストールまたは削除する](/exchange/clients-and-mobile-in-exchange-online/add-ins-for-outlook/install-or-remove-outlook-add-ins)」を参照してください。

## <a name="see-also"></a>関連項目

- [テスト用に Outlook アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [AppSource に提出する][AppSource]
- [Office アドインの設計ガイドライン](../design/add-in-design.md)
- [効果的な AppSource 登録リストを作成する](/office/dev/store/create-effective-office-store-listings)
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)

[AppSource]: /office/dev/store/submit-to-appsource-via-partner-center