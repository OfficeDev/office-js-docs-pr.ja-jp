---
title: Microsoft Azure で Office アドインをホストする | Microsoft Docs
description: アドイン Web アプリを Azure に展開して、Office クライアント アプリケーションでテストのためにアドインをサイドロードする方法について説明します。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 5db98ca65aac019a027592a442f427ee3b6126f1
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451095"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a>Microsoft Azure で Office アドインをホストする

最も簡単な Office アドインは、XML マニフェスト ファイルと HTML ページで成り立っています。XML マニフェスト ファイルには、アドインの特性 (アドインの名前や実行可能な Office クライアント アプリケーションの種類、アドインの HTML ページの URL など) を記述します。HTML ページには、ユーザーが Office クライアント アプリケーションにアドインをインストールして実行したときに操作する、Web アプリが含まれています。Office アドインの Web アプリは、Azure を含む、あらゆる Web ホスティング プラットフォームでホストできます。

この記事では、アドイン Web アプリを Azure に展開して、Office クライアント アプリケーションでテストのために[アドインをサイドロード](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)する方法について説明します。

## <a name="prerequisites"></a>前提条件 

1. [Visual Studio 2017](https://www.visualstudio.com/downloads) をインストールします。このとき、**Azure 開発**ワークロードを含めるようにします。

    > [!NOTE]
    > 既に Visual Studio 2017 がインストールされている場合は、[Visual Studio インストーラー](/visualstudio/install/modify-visual-studio)を使用して、**Azure 開発**ワークロードがインストールされていることを確認してください。 

2. Office をインストールする。

    > [!NOTE]
    > まだ Office を所持していない場合は、[1 か月間無料試用版の登録](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735)が可能です。

3. Azure サブスクリプションを取得します。

    > [!NOTE]
    > まだ Azure サブスクリプションを所持していない場合、このサブスクリプションは [Visual Studio サブスクリプションの一部として取得](https://azure.microsoft.com/ja-JP/pricing/member-offers/visual-studio-subscriptions/)できます。また、[無料試用版の登録](https://azure.microsoft.com/pricing/free-trial)も可能です。 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a>手順 1: アドインの XML マニフェスト ファイルをホストするための共有フォルダーを作成する

1. 開発用のコンピューターでエクスプローラーを開きます。

2. C:\ ドライブを右クリックして、**[新規]** > **[フォルダー]** をクリックします。

3. 新規フォルダーに「AddinManifests」という名前を付けます。

4. [AddinManifests] フォルダーを右クリックして、**[共有相手]** > **[特定の人]** をクリックします。

5. **[ファイル共有]** で、ドロップダウンの矢印をクリックして、**[すべてのユーザー]** > **[追加]** > **[共有]** をクリックします。

> [!NOTE]
> このチュートリアルでは、信頼できるカタログとしてローカルのファイル共有を使用します。アドインの XML マニフェスト ファイルは、この場所に保存することになります。現実のシナリオでは、[SharePoint カタログに XML マニフェスト ファイルを展開](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)するか、[AppSource にアドインを発行](/office/dev/store/submit-to-the-office-store)することもできます。

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a>手順 2: 信頼できるアドイン カタログにファイル共有を追加する

1. Word を起動してドキュメントを作成します。

    > [!NOTE]
    > この例では Word を使用していますが、Office アドインをサポートしている任意の Office アプリケーションを使用できます (Excel、Outlook、PowerPoint、Project など)。

2. **[ファイル]**  >  **[オプション]** を選択します。

3. **[Word オプション]** ダイアログ ボックスで、**[セキュリティ センター]** をクリックして、**[セキュリティ センターの設定]** をクリックします。

4. **[セキュリティ センター]** ダイアログ ボックスで、**[信頼できるアドイン カタログ]** をクリックします。**[カタログの URL]** として、前の手順で作成したファイル共有の汎用名前付け規則 (UNC) パス (たとえば、\\\YourMachineName\AddinManifests) を入力して、**[カタログの追加]** をクリックします。 

5. **[メニューに表示する]** チェック ボックスをオンにします。

    > [!NOTE]
    > 信頼できるアドイン カタログとして指定されている共有にアドインの XML マニフェスト ファイルを保存すると、そのアドインは、ユーザーがリボンの **[挿入]** タブから **[個人用アドイン]** をクリックしたときに、**[Office アドイン]** ダイアログ ボックスの **[共有フォルダー]** に表示されるようになります。

6. Word を終了します。

## <a name="step-3-create-a-web-app-in-azure"></a>手順 3: Azure に Web アプリを作成する

Azure に空の Web アプリを作成するには、[Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) または [Azure ポータル](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal)のどちらかを使用します。

### <a name="using-visual-studio-2017"></a>Visual Studio 2017 を使用する場合

Visual Studio 2017 を使用して Web アプリを作成するには、次の手順を実行します。

1. Visual Studio の **[表示]** メニューで、**[サーバー エクスプローラー]** をクリックします。**[Azure]** を右クリックして、**[Microsoft Azure サブスクリプションへの接続]** をクリックします。Azure サブスクリプションに接続するための指示に従います。

2. Visual Studio の **[サーバー エクスプローラー]** で、**[Azure]** を展開し、**[App Service]** を右クリックして **[新しい App Service の作成]** をクリックします。

3. **[App Service の作成]** ダイアログ ボックスで、次の情報を指定します。

      - サイトの一意の **[Web アプリの名前]** を入力します。Azure は、サイト名が azurewebsites.net ドメイン全体で一意であることを確認します。

      - このサイトの作成に使用する **[サブスクリプション]** を選択します。

      - サイトの **[リソース グループ]** を選択します。新しいグループを作成する場合は、そのグループに名前を指定する必要もあります。

      - このサイトの作成に使用する **[App Service プラン]** を選択します。新しいプランを作成する場合は、そのプランに名前を指定する必要もあります。

      - **[作成]** を選択します。

    新しい Web アプリが、**[サーバー エクスプローラー]** の **[Azure]** >> **[App Service]** >> (選択したリソース グループ) に表示されます。

4. 新しい Web アプリを右クリックして、**[ブラウザーで表示]** をクリックします。ブラウザーが開いて、「App Service アプリが作成されました」というメッセージを示す Web ページが表示されます。

5. ブラウザーのアドレス バーで、HTTPS を使用するように Web アプリの URL を変更してから **Enter** キーを押して、HTTPS プロトコルが有効であることを確認します。 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。

### <a name="using-the-azure-portal"></a>Azure ポータルを使用する場合

Azure ポータルを使用して Web アプリケーションを作成するには、次の手順を実行します。

1. Azure の資格情報を使用して、[Azure ポータル](https://portal.azure.com/)にログオンします。

2. **[新規]** > **[Web + モバイル]** > **[Web アプリ]** をクリックします。

3. **[Web アプリの作成]** ダイアログ ボックスで、次の情報を指定します。

      - サイトの一意の **[アプリ名]** を入力します。Azure は、サイト名が azureweb apps.net ドメイン全体で一意であることを確認します。

      - このサイトの作成に使用する **[サブスクリプション]** を選択します。

      - サイトの **[リソース グループ]** を選択します。新しいグループを作成する場合は、そのグループに名前を指定する必要もあります。

      - サイトの **[OS]** を選択します。

      - このサイトの作成用に使用する **[App Service プラン]** を選択します。新しいプランを作成する場合は、そのプランに名前を指定する必要もあります。

      - **[作成]** を選択します。

4. **[通知]** (Azure ポータルの上辺に配置されているベル アイコン) をクリックし、**[デプロイメントが成功しました]** の通知をクリックして Azure ポータルでサイトの **[概要]** ページを開きます。

    > [!NOTE]
    > この通知は、サイトのデプロイが完了した時点で **[デプロイは進行中です]** から **[デプロイメントが成功しました]** に変化します。

5. Azure ポータルのサイトの **[概要]** ページにある **[要点]** セクションで、**[URL]** の下に表示されている URL を選択します。ブラウザーが開いて、「App Service アプリが作成されました」というメッセージを示す Web ページが表示されます。 

6. ブラウザーのアドレス バーで、HTTPS を使用するように Web アプリの URL を変更してから **Enter** キーを押して、HTTPS プロトコルが有効であることを確認します。 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a>手順 4: Visual Studio で Office アドインを作成する

1. 管理者として Visual Studio を起動します。

2. **[ファイル]** > **[新規作成]** > **[プロジェクト]** をクリックします。

3. **[テンプレート]** の **[Visual C#]** (または **[Visual Basic]**) を展開し、**[Office/SharePoint]** を展開して **[アドイン]** をクリックします。

4. **[Word Web アドイン]** を選択してから、**[OK]** をクリックして、既定の設定を使用します。

Visual Studio により、基本的な Word アドインが作成されます。このアドインは、Web プロジェクトに一切変更を加えることなく、そのままの状態で発行できます。

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a>手順 5: Azure に Office アドイン Web アプリを発行する

1. Visual Studio でアドイン プロジェクトを開き、**[ソリューション エクスプローラー]** のソリューション ノードを展開して、ソリューションの両方のプロジェクトが表示されるようにします。

2. Web プロジェクトを右クリックして、**[発行]** をクリックします。Web プロジェクトには Office アドイン Web アプリのファイルが含まれているため、このプロジェクトを Azure に発行することになります。

3. **[発行]** タブで、次の操作を実行します。

      - **[Microsoft Azure App Service]** をクリックします。

      - **[既存のものを選択]** をクリックします。

      - **[発行]** をクリックします。

4. **[App Service]** ダイアログ ボックスで、「[手順 3: Azure に Web アプリを作成する](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure)」で作成した Web アプリを見つけて、そのアプリを選択してから **[OK]** をクリックします。 

    Visual Studio により、Office アドインの Web プロジェクトが Azure Web アプリに発行されます。Visual Studio による Web プロジェクトの発行が完了すると、ブラウザーが開いて、「App Service アプリが作成されました」というテキストを示す Web ページが表示されます。これは、Web アプリの現在の既定のページです。

 アドインの Web ページを確認するには、そのアドインが HTTPS を使用するように URL を変更して、アドインの HTML ページのパスを指定します (たとえば、https://YourDomain.azurewebsites.net/Home.html)。こうすることで、アドインの Web アプリが Azure でホストされるようになったことを確認します。ルート URL (たとえば、https://YourDomain.azurewebsites.net) をコピーします。この URL は、アドイン マニフェスト ファイルの編集時に必要になります。これについては、この記事で説明します。

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a>手順 6: アドインの XML マニフェスト ファイルを編集して展開する

1. Visual Studio の **[ソリューション エクスプローラー]** でサンプルの Office アドインを開いて、ソリューションを展開し、両方のプロジェクトが表示されるようにします。

2. Office アドイン プロジェクト (たとえば、WordWebAddIn) を展開し、マニフェスト フォルダーを右クリックして **[開く]** をクリックします。アドインの XML マニフェスト ファイルが開きます。

3. XML マニフェスト ファイルで、"~remoteAppUrl" というインスタンスをすべて検索して、Azure のアドイン Web アプリのルート URL に置換します。この URL は、前の手順で Azure にアドイン Web アプリを発行した後にコピーしたものです (たとえば、https://YourDomain.azurewebsites.net)。 

4. **[ファイル]** をクリックして、**[すべてを保存]** をクリックします。アドインの XML マニフェスト ファイルを閉じます。

5. **[ソリューション エクスプローラー]** に戻って、マニフェストのフォルダーを右クリックして、**[エクスプローラーでフォルダーを開く]** をクリックします。

6. アドインの XML マニフェスト ファイル (たとえば、WordWebAddIn.xml) をコピーします。 

7. 「[手順 1: 共有フォルダーを作成する](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file)」で作成したネットワーク ファイル共有を参照して、そのフォルダー内にマニフェスト ファイルを貼り付けます。

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a>手順 7: Office クライアント アプリケーションにアプリを挿入し、実行する

1. Word を起動してドキュメントを作成します。

2. リボンで、**[挿入]** > **[個人用アドイン]** をクリックします。

3. **[Office アドイン]** ダイアログ ボックスで、**[共有フォルダー]** をクリックします。Word により、信頼できるアドイン カタログとしてリストしたフォルダー (「[手順 2: 信頼できるアドイン カタログにファイル共有を追加する](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)」で指定したもの) がスキャンされ、アドインがダイアログ ボックスに表示されます。サンプル アドインのアイコンが表示されます。

4. アドインを選択して、**[追加]** をクリックします。リボンに、そのアドインの **[作業ウィンドウの表示]** ボタンが追加されます。

5. **[ホーム]** タブのリボンで、**[作業ウィンドウの表示]** ボタンをクリックします。現在のドキュメントの右側の作業ウィンドウ内でアドインが開きます。

6. アドインが動作していることを確認するために、ドキュメント内のテキストを選択して、作業ウィンド内の **[Highlight!]** ボタンをクリックします。

## <a name="see-also"></a>関連項目

- [Office アドインを発行する](../publish/publish.md)
- [発行のための準備として Visual Studio を使用してアドインをパッケージ化する](../publish/package-your-add-in-using-visual-studio.md)
