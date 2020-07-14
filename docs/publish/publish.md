---
title: Office アドインを展開し、発行する
description: テスト目的またはユーザーに配布する目的で Office アドインを展開するための方法とオプション。
ms.date: 06/02/2020
localization_priority: Priority
ms.openlocfilehash: 797abbde43e6172ba26f3dd4b128fb06f1e70bec
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094184"
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

## <a name="deployment-options-by-office-host-and-add-in-type"></a>Office ホストとアドインの種類による展開オプション

選択可能な展開オプションは、対象の Office ホストや作成するアドインの種類によって異なります。

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

A SharePoint app catalog is a special site collection that you can create to host Word, Excel, and PowerPoint add-ins. Because SharePoint catalogs don't support new add-in features implemented in the `VersionOverrides` node of the manifest, including add-in commands, we recommend that you use Centralized Deployment via the admin center if possible. Add-in commands deployed via a SharePoint catalog open in a task pane by default.

If you are deploying add-ins in an on-premises environment, use a SharePoint catalog. For details, see [Publish task pane and content add-ins to a SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> [!NOTE]
> SharePoint カタログは Office on Mac をサポートしません。 Office アドインを Mac クライアントに展開するには、そのアドインを [AppSource](/office/dev/store/submit-to-the-office-store) に提出する必要があります。

### <a name="outlook-add-in-deployment"></a>Outlook アドインの展開

Azure AD の ID サービスを使用しないオンプレミス環境およびオンライン環境では、Exchange サーバー経由で Outlook アドインを展開することができます。

Outlook アドインの展開には以下が必要です。

- Microsoft 365、Exchange Online、または Exchange Server 2013 以降
- Outlook 2013 以降

To assign add-ins to tenants, you use the Exchange admin center to upload a manifest directly, either from a file or a URL, or add an add-in from AppSource. To assign add-ins to individual users, you must use Exchange PowerShell. For details, see [Install or remove Outlook add-ins for your organization](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx) on TechNet.

## <a name="see-also"></a>関連項目

- [テスト用に Outlook アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [AppSource に提出する][AppSource]
- [Office アドインの設計ガイドライン](../design/add-in-design.md)
- [効果的な AppSource 登録リストを作成する](/office/dev/store/create-effective-office-store-listings)
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)

[AppSource]: /office/dev/store/submit-to-appsource-via-partner-center
[Office Add-in host and platform availability]: ../overview/office-add-in-availability
