---
title: 作業ウィンドウ アドインとコンテンツ アドインを SharePoint アプリ カタログに発行する
description: 組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアプリ カタログに Office アドインのマニフェスト ファイルをアップロードできます。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 5557dd31e829fac2c2dbd421200da46a5c3b9b99
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185590"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a>作業ウィンドウ アドインとコンテンツ アドインを SharePoint アプリ カタログに発行する

アプリ カタログは、Office アドインと SharePoint アドインのドキュメント ライブラリをホストする SharePoint Web アプリケーションまたは SharePoint Online テナンシーの専用サイト コレクションです。組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアプリ カタログに Office アドインのマニフェスト ファイルをアップロードできます。管理者がアプリ カタログを信頼できるカタログとして登録すると、ユーザーは Office クライアント アプリケーションで挿入 UI からアドインを挿入できます。

> [!IMPORTANT]
> - SharePoint のアプリ カタログでは、アドイン コマンドなど、[アドイン マニフェスト](../develop/add-in-manifests.md)の `VersionOverrides` ノードで実装されるアドイン機能がサポートされていません。
> - クラウド環境またはハイブリッド環境をターゲットにしている場合は、アドインの発行に [Office 365 管理センターからの一元展開を使用する](../publish/centralized-deployment.md)ことをお勧めします。
> - SharePoint のアプリ カタログは Office on Mac ではサポートされていません。 Office アドインを Mac クライアントに展開するには、そのアドインを [AppSource](/office/dev/store/submit-to-the-office-store) に提出する必要があります。

## <a name="create-an-app-catalog"></a>アプリ カタログを作成する

次のいずれかのセクションの手順を完了し、オンプレミス SharePoint サーバーまたは Office 365 を使用して、アプリ カタログを作成します。

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a>オンプレミス SharePoint サーバーでアプリ カタログを作成する

SharePoint アプリ カタログを作成するには、[web アプリケーションのアプリ カタログサイトを作成する](/sharepoint/administration/manage-the-app-catalog)を参照してください。

アプリ カタログを作成したら [Office アドインを発行する](#publish-an-office-add-in) 手順に従います。

### <a name="to-create-an-app-catalog-on-office-365"></a>Office 365 でアプリ カタログを作成する

SharePoint アプリカタログを作成するには、「[アプリカタログサイトコレクションを作成](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection)する」の手順に従ってください。 アプリカタログを作成したら、次のセクションの手順に従って、Office アドインを発行します。

## <a name="publish-an-office-add-in"></a>Office アドインの発行

次のいずれかのセクションの手順を完了し、Office アドインを Office 365 またはオンプレミス SharePoint サーバーのアプリ カタログに発行します。 

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-office-365"></a>Office アドインを Office 365 の SharePoint アプリ カタログに発行する

1. [新しい SharePoint 管理センターの [アクティブなサイト] ページ](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true)に移動し、組織の[管理者権限](/sharepoint/sharepoint-admin-role)が付与されているアカウントでサインインします。

>[!NOTE]
>Office 365 Germany を使用している場合は、[Microsoft 365 管理センターにサインイン](https://go.microsoft.com/fwlink/p/?linkid=848041)し、SharePoint 管理センターを参照してから [その他の機能] ページを開きます。 <br>21Vianet (中国) によって運用されている Office 365 を使用している場合は、[Microsoft 365 管理センターにサインイン](https://go.microsoft.com/fwlink/p/?linkid=850627)し、次に SharePoint 管理センターに移動して [その他の機能] ページを開きます。
 
2. [URL] 列で URL を選択して、アプリカタログサイトを開きます。 

>[!NOTE]
>前のセクションでアプリカタログサイトを作成したばかりの場合は、サイトのセットアップが完了するまでに数分かかることがあります。

3. [**Office 用アプリを配信する**] を選択します。
4. [**Office 用アプリ**] ページで、[**新規**] を選択します。
5. [**ドキュメントの追加**] ダイアログで、[**ファイルの選択**] ボタンをクリックします。
6. アップロードする[マニフェスト](../develop/add-in-manifests.md) ファイルを見つけて指定し、[**開く**] を選択します。
7. [**ドキュメントの追加**] ダイアログで、[**OK**] を選択します。

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a>オンプレミスの SharePoint サーバーでアプリ カタログにアドインを発行する

1. **[サーバーの全体管理]** ページを開きます。
2. 左側の作業ウィンドウで、[**アプリ**] を選択します。
3. **[アプリ]** ページの **[アプリの管理]** で **[アプリ カタログの管理]** を選択します。
4. **[アプリ カタログの管理]** ページの ** Web アプリケーション ** セレクターで正しい Web アプリケーションが選択されていることを確認します。
5. **サイト URL**の下にある URL を選び、アプリ カタログサイトを開きます。
6. [**Office 用アプリを配信する**] を選択します。
7. [**Office 用アプリ**] ページで、[**新規**] を選択します。
8. [**ドキュメントの追加**] ダイアログで、[**ファイルの選択**] ボタンをクリックします。
9. アップロードする[マニフェスト](../develop/add-in-manifests.md) ファイルを見つけて指定し、[**開く**] を選択します。
10. [**ドキュメントの追加**] ダイアログで、[**OK**] を選択します。

## <a name="insert-office-add-ins-from-the-app-catalog"></a>アプリ カタログから Office アドインを挿入する

オンライン Office アプリケーションの場合は、次の手順を実行してアプリ カタログから Office アドインを見つけることができます。

1. オンライン Office アプリケーション (Excel、PowerPoint、または Word) を開きます。
2. 文書を作成または開く。
3. **[挿入]** > **[アドイン]** を選択します。
4. [Office アドイン] ダイアログの **[自分の所属組織]** タブを選択します。Office アドインのリストが表示されます。
5. Office アドインを選択し、 **追加** を選択します。

デスクトップの Office アプリケーションの場合は、次の手順を実行してアプリ カタログから Office アドインを見つけることができます。

1. デスクトップ Office アプリケーション (Excel、Word、または PowerPoint) を開きます。
2. **[ファイル]**  >  **[オプション]**  >  **[セキュリティ センター]**  >  **[セキュリティ センターの設定]**  >  **[信頼できるアドイン カタログ]** の順に選択します。
3. [**カタログ URL** ] ボックスに SharePoint アプリ カタログの URL を入力し、**[カタログの追加]** を選択します。
    短い URL を使用します。 たとえば、SharePoint アプリ カタログの URL が次のような場合:
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`
    
    親サイト コレクションの URL のみを指定します。
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
4. Office アプリケーションを閉じてから、もう一度開きます。 
5. **[挿入]** > **[アドインの取得]** の順に選択します。
4. [Office アドイン] ダイアログの **[自分の所属組織]** タブを選択します。Office アドインのリストが表示されます。
5. Office アドインを選択し、 **追加**を選択します。

または、管理者はグループ ポリシーを使用して SharePoint のアプリ カタログを指定できます。 関連するポリシー設定を「[365 ProPlus、Office 2019、Office 2016 の管理用テンプレート ファイル (ADMX/ADML)](https://www.microsoft.com/download/details.aspx?id=49030) 」で利用できますし、**ユーザー構成\ポリシー\管理用テンプレート\Microsoft Office 2016\セキュリティ設定\セキュリティ センター\信頼できるカタログ**の下にも見つけられます。
