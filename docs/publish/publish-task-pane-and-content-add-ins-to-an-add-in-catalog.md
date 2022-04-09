---
title: 作業ウィンドウ アドインとコンテンツ アドインを SharePoint アプリ カタログに発行する
description: 組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアプリ カタログに Office アドインのマニフェスト ファイルをアップロードできます。
ms.date: 11/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: b545d81c8e728a5d6c9129bfa166e23a76ce2b04
ms.sourcegitcommit: 3d37c42f5e465dac52d231d31717bdbb3bfa0e30
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/10/2021
ms.locfileid: "60889994"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a>作業ウィンドウ アドインとコンテンツ アドインを SharePoint アプリ カタログに発行する

アプリ カタログは、Office アドインと SharePoint アドインのドキュメント ライブラリをホストする SharePoint Web アプリケーションまたは SharePoint Online テナンシーの専用サイト コレクションです。組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアプリ カタログに Office アドインのマニフェスト ファイルをアップロードできます。管理者がアプリ カタログを信頼できるカタログとして登録すると、ユーザーは Office クライアント アプリケーションで挿入 UI からアドインを挿入できます。

> [!IMPORTANT]
>
> - SharePoint のアプリ カタログでは、アドイン コマンドなど、[アドイン マニフェスト](../develop/add-in-manifests.md)の `VersionOverrides` ノードで実装されるアドイン機能がサポートされていません。
> - クラウド環境またはハイブリッド環境をターゲットにしている場合は、アドインの発行に [Office 365 管理センターからの一元展開を使用](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) することをお勧めします。
> - SharePoint のアプリ カタログは Mac 用 Office ではサポートされていません。Office アドインを Mac クライアントに展開するには、[AppSource](/office/dev/store/submit-to-the-office-store) に提出する必要があります。

## <a name="create-an-app-catalog"></a>アプリ カタログを作成する

次のいずれかのセクションの手順を完了し、オンプレミスの SharePoint サーバーまたは Microsoft 365 を使用して、アプリ カタログを作成します。

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a>オンプレミス SharePoint サーバーでアプリ カタログを作成する

SharePoint アプリ カタログを作成するには、[web アプリケーションのアプリ カタログサイトを作成する](/sharepoint/administration/manage-the-app-catalog)を参照してください。

アプリ カタログを作成したら [Office アドインを発行する](#publish-an-office-add-in) 手順に従います。

### <a name="to-create-an-app-catalog-on-microsoft-365"></a>Microsoft 365でアプリカタログを作成する場合

SharePoint アプリ カタログを作成するには、「[アプリ カタログ サイト コレクションの作成](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection)」 の手順に従います。 アプリカタログを作成したら、次のセクションの手順に従って、Office アドインを発行することができます。

## <a name="publish-an-office-add-in"></a>Office アドインの発行

次のいずれかのセクションの手順を完了し、Microsoft 365 またはオンプレミスの SharePoint Server 上のアプリカタログに Office アドインを発行します。

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-microsoft-365"></a>Microsoft 365 の SharePoint アプリカタログに Office アドインを発行する場合

1. [新しい SharePoint 管理センターの [アクティブなサイト] ページ](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true)に移動し、組織の[管理者権限](/sharepoint/sharepoint-admin-role)が付与されているアカウントでサインインします。

    > [!NOTE]
    > 21Vianet (中国) によって運用されている Microsoft 365 を使用している場合は、[Microsoft 365 管理センターにサインイン](https://go.microsoft.com/fwlink/p/?linkid=850627) し、次に SharePoint 管理センターに移動して [その他の機能] ページを開きます。

1. [URL] 列で URL を選択して、アプリ カタログ サイトを開きます。

    > [!NOTE]
    > 前のセクションでアプリ カタログ サイトを作成したばかりの場合、サイトの設定が完了するまで数分かかることがあります。

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

または、管理者はグループ ポリシーを使用して SharePoint のアプリ カタログを指定できます。 関連するポリシー設定を 「[Microsoft 365 アプリ、Office LTSC 2021、Office 2019、および Office 2016 の管理用テンプレート ファイル (ADMX/ADML)](https://www.microsoft.com/download/details.aspx?id=49030) 」 で利用できますし、「**ユーザー構成\ポリシー\管理用テンプレート\Microsoft Office 2016\セキュリティ設定\トラスト センター\信頼できるカタログ**」 の下にも見つけられます。
