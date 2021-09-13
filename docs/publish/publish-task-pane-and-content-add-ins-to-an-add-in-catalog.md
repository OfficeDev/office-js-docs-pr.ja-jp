---
title: 作業ウィンドウ アドインとコンテンツ アドインを SharePoint アプリ カタログに発行する
description: 組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアプリ カタログに Office アドインのマニフェスト ファイルをアップロードできます。
ms.date: 07/27/2021
ms.localizationpriority: medium
ms.openlocfilehash: 786fbd24790a1b8205fc3b0e8a15ce591cf66ca4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154561"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a>作業ウィンドウ アドインとコンテンツ アドインを SharePoint アプリ カタログに発行する

アプリ カタログは、Office アドインと SharePoint アドインのドキュメント ライブラリをホストする SharePoint Web アプリケーションまたは SharePoint Online テナンシーの専用サイト コレクションです。組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアプリ カタログに Office アドインのマニフェスト ファイルをアップロードできます。管理者がアプリ カタログを信頼できるカタログとして登録すると、ユーザーは Office クライアント アプリケーションで挿入 UI からアドインを挿入できます。

> [!IMPORTANT]
>
> - SharePoint のアプリ カタログでは、アドイン コマンドなど、[アドイン マニフェスト](../develop/add-in-manifests.md)の `VersionOverrides` ノードで実装されるアドイン機能がサポートされていません。
> - クラウド環境またはハイブリッド環境をターゲットにしている場合は、アドインを発行[](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)するために、Microsoft 365 管理センター経由で統合アプリを使用することをお勧めします。
> - SharePoint のアプリ カタログは Office on Mac ではサポートされていません。 Office アドインを Mac クライアントに展開するには、そのアドインを [AppSource](/office/dev/store/submit-to-the-office-store) に提出する必要があります。

## <a name="create-an-app-catalog"></a>アプリ カタログを作成する

以下のいずれかのセクションの手順を実行して、オンプレミスのサーバーまたはサーバー上でアプリ カタログSharePoint作成Microsoft 365。

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a>オンプレミス SharePoint サーバーでアプリ カタログを作成する

SharePoint アプリ カタログを作成するには、[web アプリケーションのアプリ カタログサイトを作成する](/sharepoint/administration/manage-the-app-catalog)を参照してください。

アプリ カタログを作成したら [Office アドインを発行する](#publish-an-office-add-in) 手順に従います。

### <a name="to-create-an-app-catalog-on-microsoft-365"></a>アプリ カタログを作成するには、Microsoft 365

アプリ カタログをSharePointするには、「アプリ カタログ サイト コレクションの作成[」の手順に従います](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection)。 アプリ カタログを作成したら、次のセクションの手順に従って、アドインにOffice発行します。

## <a name="publish-an-office-add-in"></a>Office アドインの発行

次のいずれかのセクションの手順を実行して、Office またはオンプレミスの Microsoft 365 サーバー上のアプリ カタログにSharePointします。

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-microsoft-365"></a>アプリ カタログにOfficeアドインを発行するには、SharePointアプリ カタログMicrosoft 365

1. [新しい SharePoint 管理センターの [アクティブなサイト] ページ](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true)に移動し、組織の[管理者権限](/sharepoint/sharepoint-admin-role)が付与されているアカウントでサインインします。

    > [!NOTE]
    > ドイツ語をMicrosoft 365場合は、Microsoft 365 管理センター[](https://go.microsoft.com/fwlink/p/?linkid=848041)にサインインし、SharePoint管理センターを参照し、[その他の機能] ページを開きます。 <br>21Vianet (中国Microsoft 365操作している場合は、Microsoft 365 管理センター にサインインし[、SharePoint](https://go.microsoft.com/fwlink/p/?linkid=850627)管理センターを参照し、[その他の機能] ページを開きます。

1. [URL] 列で URL を選択して、アプリ カタログ サイトを開きます。

    > [!NOTE]
    > 前のセクションでアプリ カタログ サイトを作成し終えたばかりの場合、サイトのセットアップが完了するには数分かかる場合があります。

1. [**Office 用アプリを配信する**] を選択します。
1. [**Office 用アプリ**] ページで、[**新規**] を選択します。
1. [**ドキュメントの追加**] ダイアログで、[**ファイルの選択**] ボタンをクリックします。
1. アップロードする [マニフェスト](../develop/add-in-manifests.md) ファイルを見つけて指定し、[**開く**] を選択します。
1. [**ドキュメントの追加**] ダイアログで、[**OK**] を選択します。

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a>オンプレミスの SharePoint サーバーでアプリ カタログにアドインを発行する

1. **[サーバーの全体管理]** ページを開きます。
1. 左側の作業ウィンドウで、[**アプリ**] を選択します。
1. **[アプリ]** ページの **[アプリの管理]** で **[アプリ カタログの管理]** を選択します。
1. **[アプリ カタログの管理]** ページの **Web アプリケーション** セレクターで正しい Web アプリケーションが選択されていることを確認します。
1. **サイト URL** の下にある URL を選び、アプリ カタログサイトを開きます。
1. [**Office 用アプリを配信する**] を選択します。
1. [**Office 用アプリ**] ページで、[**新規**] を選択します。
1. [**ドキュメントの追加**] ダイアログで、[**ファイルの選択**] ボタンをクリックします。
1. アップロードする [マニフェスト](../develop/add-in-manifests.md) ファイルを見つけて指定し、[**開く**] を選択します。
1. [**ドキュメントの追加**] ダイアログで、[**OK**] を選択します。

## <a name="insert-office-add-ins-from-the-app-catalog"></a>アプリ カタログから Office アドインを挿入する

オンライン Office アプリケーションの場合は、次の手順を実行してアプリ カタログから Office アドインを見つけることができます。

1. オンライン Office アプリケーション (Excel、PowerPoint、または Word) を開きます。
1. 文書を作成または開く。
1. **[挿入]** > **[アドイン]** を選択します。
1. [Office アドイン] ダイアログの **[自分の所属組織]** タブを選択します。Office アドインのリストが表示されます。
1. Office アドインを選択し、 **追加** を選択します。

デスクトップの Office アプリケーションの場合は、次の手順を実行してアプリ カタログから Office アドインを見つけることができます。

1. デスクトップ Office アプリケーション (Excel、Word、または PowerPoint) を開きます。
1. **[ファイル]**  >  **[オプション]**  >  **[セキュリティ センター]**  >  **[セキュリティ センターの設定]**  >  **[信頼できるアドイン カタログ]** の順に選択します。
1. [**カタログ URL** ] ボックスに SharePoint アプリ カタログの URL を入力し、**[カタログの追加]** を選択します。
    短い URL を使用します。 たとえば、SharePoint アプリ カタログの URL が次のような場合:
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`

    親サイト コレクションの URL のみを指定します。
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
1. Office アプリケーションを閉じてから、もう一度開きます。
1. **[挿入]** > **[アドインの取得]** の順に選択します。
1. [Office アドイン] ダイアログの **[自分の所属組織]** タブを選択します。Office アドインのリストが表示されます。
1. Office アドインを選択し、 **追加** を選択します。

または、管理者はグループ ポリシーを使用して SharePoint のアプリ カタログを指定できます。 関連するポリシー設定は [、Microsoft 365 Apps、Office 2019、および Office 2016](https://www.microsoft.com/download/details.aspx?id=49030)の管理用テンプレート ファイル (ADMX/ADML) で使用できます。「ユーザー構成\ポリシー\管理用テンプレート **\Microsoft Office 2016\Security 設定\Trust Center\Trusted Catalogs」** の下にあります。
