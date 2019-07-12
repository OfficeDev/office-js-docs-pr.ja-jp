---
title: 作業ウィンドウ アドインとコンテンツ アドインを SharePoint アプリ カタログに発行する
description: 組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアプリ カタログに Office アドインのマニフェスト ファイルをアップロードできます。
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 106dfd2b1610be92f1b53dc1644ff3f8c60c0543
ms.sourcegitcommit: 9c5a836d4464e49846c9795bf44cfe23e9fc8fbe
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2019
ms.locfileid: "35617031"
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

1. Microsoft 365 管理センターに移動します。 管理センターの検索方法については、「[Microsoft 365 管理センターについて](/office365/admin/admin-overview/about-the-admin-center)」を参照してください。

2. Microsoft 365 管理センターのページで [**管理センター**] のリストを展開し、[**SharePoint**] を選択します。

    > [!NOTE]
    > カタログを作成するには、従来の SharePoint 管理センターを使用する必要があります。 新しい SharePoint 管理センターにいる場合は、左側のウィンドウで「**従来の SharePoint 管理センター**」を選択します。

3. 左側の作業ウィンドウで、[**アプリ**] を選択します。

4. [**アプリ**] ページで、[**アプリ カタログ**] を選択します。
    > [!NOTE]
    > アプリ カタログが既に作成されてこのページに表示されている場合は、残りの手順をスキップしてこの記事の次のセクションに移動し、カタログにアドインを発行できます。

5. [**アプリ カタログ サイト**] ページで、[**OK**] を選択して既定のオプションを受け入れ、新しいアプリ カタログ サイトを作成します。

6. [**アプリ カタログ サイト コレクションを作成する**] ページで、アプリ カタログ サイトのタイトルを指定します。

7. [**Web サイトのアドレス**]を指定します。

8. [**管理者**]を指定します。

9. [**サーバー リソース クォータ**] に 0 (ゼロ) を設定します。 (サーバー リソース クォータは、パフォーマンスが低いサンドボックス ソリューションのスロットルに関連していますが、このアプリ カタログ サイトにはサンドボックス ソリューションをインストールしません。)

10. [**OK**] を選択します。

## <a name="publish-an-office-add-in"></a>Office アドインの発行

次のいずれかのセクションの手順を完了し、Office アドインを Office 365 またはオンプレミス SharePoint サーバーのアプリ カタログに発行します。 

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-office-365"></a>Office アドインを Office 365 の SharePoint アプリ カタログに発行する

1. Microsoft 365 管理センターに移動します。 管理センターの検索方法については、「[Microsoft 365 管理センターについて](/office365/admin/admin-overview/about-the-admin-center)」を参照してください。
2. Microsoft 365 管理センターのページで [**管理センター**] のリストを展開し、[**SharePoint**] を選択します。
    > [!NOTE]
    > カタログを作成するには、従来の SharePoint 管理センターを使用する必要があります。 新しい SharePoint 管理センターにいる場合は、左側のウィンドウで「**従来の SharePoint 管理センター**」を選択します。
3. 左側の作業ウィンドウで、[**アプリ**] を選択します。
4. [**アプリ**] ページで、[**アプリ カタログ**] を選択します。
5. [**Office 用アプリを配信する**] を選択します。
6. [**Office 用アプリ**] ページで、[**新規**] を選択します。
7. [**ドキュメントの追加**] ダイアログで、[**ファイルの選択**] ボタンをクリックします。
8. アップロードする[マニフェスト](../develop/add-in-manifests.md) ファイルを見つけて指定し、[**開く**] を選択します。
9. [**ドキュメントの追加**] ダイアログで、[**OK**] を選択します。

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
