---
title: Microsoft Azure で Office アドインをホストする | Microsoft Docs
description: アドイン Web アプリを Azure に展開して、Office クライアント アプリケーションでテストのためにアドインをサイドロードする方法について説明します。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: a30f1a8219501a68e6f46f013ef46640a59fe4e9
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094233"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a>Microsoft Azure で Office アドインをホストする

The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office desktop applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.

この記事では、アドイン Web アプリを Azure に展開して、Office クライアント アプリケーションでテストのために[アドインをサイドロード](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)する方法について説明します。

## <a name="prerequisites"></a>前提条件 

1. [Visual Studio 2019](https://www.visualstudio.com/downloads) をインストールし、**Azure 開発**ワークロードを含めるよう選択します。

    > [!NOTE]
    > 既に Visual Studio 2019 がインストールされている場合は、[Visual Studio インストーラーを使用](/visualstudio/install/modify-visual-studio)して、**Azure 開発**ワークロードがインストールされていることを確認してください。 

2. Office をインストールする。

    > [!NOTE]
    > まだ Office を所持していない場合は、[1 か月間無料試用版の登録](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735)が可能です。

3. Azure サブスクリプションを取得します。

    > [!NOTE]
    > まだ Azure サブスクリプションを所持していない場合、このサブスクリプションは [Visual Studio サブスクリプションの一部として取得](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/)できます。また、[無料試用版の登録](https://azure.microsoft.com/pricing/free-trial)も可能です。 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a>手順 1: アドインの XML マニフェスト ファイルをホストするための共有フォルダーを作成する

1. 開発用のコンピューターでエクスプローラーを開きます。

2. C:\ ドライブを右クリックして、**[新規]** > **[フォルダー]** をクリックします。

3. 新規フォルダーに「AddinManifests」という名前を付けます。

4. [AddinManifests] フォルダーを右クリックして、**[共有相手]** > **[特定の人]** をクリックします。

5. **[ファイル共有]** で、ドロップダウンの矢印をクリックして、**[すべてのユーザー]** > **[追加]** > **[共有]** をクリックします。

> [!NOTE]
> In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a>手順 2: 信頼できるアドイン カタログにファイル共有を追加する

1. Word を起動してドキュメントを作成します。

    > [!NOTE]
    > この例では Word を使用していますが、Office アドインをサポートしている任意の Office アプリケーションを使用できます (Excel、Outlook、PowerPoint、Project など)。

2. **[ファイル]**  >  **[オプション]** を選択します。

3. **[Word オプション]** ダイアログ ボックスで、**[セキュリティ センター]** をクリックして、**[セキュリティ センターの設定]** をクリックします。

4. In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**. 

5. **[メニューに表示する]** チェック ボックスをオンにします。

    > [!NOTE]
    > 信頼できるアドイン カタログとして指定されている共有にアドインの XML マニフェスト ファイルを保存すると、そのアドインは、ユーザーがリボンの **[挿入]** タブから **[個人用アドイン]** をクリックしたときに、**[Office アドイン]** ダイアログ ボックスの **[共有フォルダー]** に表示されるようになります。

6. Word を終了します。

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a>手順 3: Azure ポータルを使用して Azure で Web アプリを作成する

Azure ポータルを使用して Web アプリケーションを作成するには、次の手順を実行します。

1. Azure の資格情報を使用して、[Azure ポータル](https://portal.azure.com/)にログオンします。

2. [**Azure サービス**] で [**Web アプリ**] を選択します。

3. [**App Service**] ページで、[**追加**] を選択します。 この情報を提供してください:

      - このサイトの作成に使用する **[サブスクリプション]** を選択します。
      
      - Choose the **Resource Group** for your site. If you create a new group, you also need to name it.
      
      - サイトの一意の **[アプリ名]** を入力します。 Azure は、サイト名が azureweb apps.net ドメイン全体で一意であることを確認します。

      - コードを使用して発行するか、Docker コンテナを使用して発行するかを選択します。

      - **ランタイム スタック**を指定します。

      - サイトの **OS** を選択します。

      - **地域**を選択します。

      - このサイトの作成に使用する [**App Service プラン**] を選択します。

      - [**作成**] を選択します。

4. 次のページでは、展開が進行中であること、完了したことが通知されます。 完了したら、[**リソースに移動**] を選択します。  

5. [**概要**] セクションで、[**URL**] の下に表示される URL を選択します。 ブラウザが開き、"App Service アプリが起動し、実行中です" というメッセージを含む Web ページが表示されます。

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a>手順 4: Visual Studio で Office アドインを作成する

1. 管理者として Visual Studio を起動します。

2. [**新規プロジェクトの作成**] を選択します。

3. 検索ボックスを使用して、**アドイン**と入力します。

4. プロジェクト タイプとして **Word Web アドイン**を選択し、[**次へ**] を選択して規定の設定を使用します。

Visual Studio は、Web プロジェクトに変更を加えることなくそのまま発行できる、基本的な Word アドインを作成します。 Excel などの異なる Office ホスト タイプのアドインを作成するには、この手順を繰り返して、目的の Office ホストのプロジェクト タイプを選択します。

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a>手順 5: Azure に Office アドイン Web アプリを発行する

1. アドイン プロジェクトを Visual Studio で開いた状態で、**ソリューション エクスプローラー**でソリューション ノードを展開し、**App Service** を選択します。

2. Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.

3. **[発行]** タブで、次の操作を実行します。

      - **[Microsoft Azure App Service]** をクリックします。

      - **[既存のものを選択]** をクリックします。

      - **[発行]** をクリックします。

4. Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.

5. ルート URL (たとえば https://YourDomain.azurewebsites.net)) をコピーします。この URL は、アドイン マニフェスト ファイルの編集時に必要になります。これについては、この記事で説明します。

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a>手順 6: アドインの XML マニフェスト ファイルを編集して展開する

1. Visual Studio の **[ソリューション エクスプローラー]** でサンプルの Office アドインを開いて、ソリューションを展開し、両方のプロジェクトが表示されるようにします。

2. Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.

3. In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net). 

4. [**ファイル**] をクリックして、[**すべてを保存**] をクリックします。 次に、アドイン XML マニフェスト ファイル (WordWebAddIn.xml など) をコピーします。

5. **ファイル エクスプローラー** プログラムを使用して、「[手順 1: 共有フォルダーを作成する](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file)」で作成したネットワーク ファイル共有を参照し、マニフェスト ファイルをそのフォルダー内に貼り付けます。

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a>手順 7: Office クライアント アプリケーションにアプリを挿入し、実行する

1. Word を起動してドキュメントを作成します。

2. リボンで、**[挿入]** > **[個人用アドイン]** をクリックします。

3. In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.

4. Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.

5. On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.

6. Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.

## <a name="see-also"></a>関連項目

- [Office アドインを発行する](../publish/publish.md)
- [Visual Studio を使用してアドインを発行する](../publish/package-your-add-in-using-visual-studio.md)
