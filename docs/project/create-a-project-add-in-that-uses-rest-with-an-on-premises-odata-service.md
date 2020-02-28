---
title: 社内の Project Server OData サービスで REST を使用する Project アドインを作成する
description: ''
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: 73099f244ef68fc1633adc9b842f64830761805f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324913"
---
# <a name="create-a-project-add-in-that-uses-rest-with-an-on-premises-project-server-odata-service"></a>社内の Project Server OData サービスで REST を使用する Project アドインを作成する

この記事では、作業中のプロジェクトのコストと作業データを現在の Project Web App インスタンスのすべてのプロジェクトの平均値と比較する、Project Professional 2013 用の作業ウィンドウアドインをビルドする方法について説明します。アドインは、jQuery ライブラリで REST を使用して、Project Server 2013 の**Projectdata** OData レポートサービスにアクセスします。

この記事のコードは、Microsoft Corporation の Saurabh Sanghvi と Arvind Iyer が開発したサンプルに基づいています。

## <a name="prerequisites-for-creating-a-task-pane-add-in-that-reads-project-server-reporting-data"></a>Project Server のレポート データを読み取る作業ウィンドウ アドインを作成するための前提条件

Project Server 2013 の社内インストールにおける Project Web App インスタンスの**Projectdata**サービスを読み取る project 作業ウィンドウアドインを作成するための前提条件を以下に示します。

- 使用するローカルの開発用コンピューターに最新のサービス パックと Windows 更新プログラムをインストールしてあることを確認します。オペレーティング システムは、Windows 7、Windows 8、Windows Server 2008、Windows Server 2012 のいずれでもかまいません。

- Project Professional 2013 は、Project Web App に接続するために必要です。Visual Studio での**F5**デバッグを有効にするには、開発コンピューターに Project Professional 2013 がインストールされている必要があります。

    > [!NOTE]
    > Project Standard 2013 でも作業ウィンドウ アドインをホストできますが、Project Web App にはログオンできません。

- Office Developer Tools for Visual Studio を備えた Visual Studio 2015 には、Office アドインと SharePoint アドインの作成用のテンプレートが含まれています。最新バージョンの Office Developer Tools がインストールされていることを確認してください。 _Office アドインと SharePoint のダウンロード_の「 [ツール](https://developer.microsoft.com/office/docs) 」セクションを参照してください。

- この記事の手順とコード例では、ローカルドメインの Project Server 2013 の**Projectdata**サービスにアクセスします。この記事に記載されている jQuery メソッドは、web 上の Project では機能しません。

    開発用コンピューターから**Projectdata**サービスにアクセスできることを確認します。

### <a name="procedure-1-to-verify-that-the-projectdata-service-is-accessible"></a>手順 1. ProjectData サービスにアクセスできることを確認するには

1. ブラウザーで REST クエリからの XML データの直接表示を可能にするには、フィードの読み取りビューをオフにします。Internet Explorer でこれを行う方法については、「 [Project Server 2013 レポート データの OData フィードにクエリを実行する](/previous-versions/office/project-odata/jj163048(v=office.15))」の手順 1. のステップ 4. を参照してください。

2. ブラウザーで次の URL を使用して**projectdata**サービスに対してクエリを実行します。 ** http://ServerName /projectdata/_api/projectdata**。 たとえば、Project Web App インスタンスが `http://MyServer/pwa` である場合、ブラウザーは次の結果を示します。

    ```xml
    <?xml version="1.0" encoding="utf-8"?>
        <service xml:base="http://myserver/pwa/_api/ProjectData/"
        xmlns="https://www.w3.org/2007/app"
        xmlns:atom="https://www.w3.org/2005/Atom">
        <workspace>
            <atom:title>Default</atom:title>
            <collection href="Projects">
                <atom:title>Projects</atom:title>
            </collection>
            <collection href="ProjectBaselines">
                <atom:title>ProjectBaselines</atom:title>
            </collection>
            <!-- ... and 33 more collection elements -->
        </workspace>
        </service>
    ```

3. 結果を確認するためにネットワーク資格情報の入力が必要になることもあります。ブラウザーでアクセスが拒否されことを示すメッセージ (エラー 403) が表示された場合は、その Project Web App インスタンスに対するログオンのアクセス許可を与えられていないか、管理者のサポートを必要とするネットワーク上の問題が発生しています。

## <a name="using-visual-studio-to-create-a-task-pane-add-in-for-project"></a>Visual Studio を使用して Project 用の作業ウィンドウ アドインを作成する

Office Developer Tools for Visual Studio には、Project 2013 用の作業ウィンドウアドインのテンプレートが含まれています。**HelloProjectOData**という名前のソリューションを作成すると、ソリューションには次の2つの Visual Studio プロジェクトが含まれます。

- アドインプロジェクトには、ソリューションの名前が指定されます。このファイルには、アドインの XML マニフェストファイルが含まれており、.NET Framework 4.5 を対象としています。手順3は、 **HelloProjectOData**アドインのマニフェストを変更する手順を示しています。

- Web プロジェクトには、 **HelloProjectODataWeb**という名前が付けられます。これには、作業ウィンドウの web ページ、JavaScript ファイル、CSS ファイル、画像、参照、および構成ファイルが含まれます。Web プロジェクトは .NET Framework 4 を対象としています。手順4および手順5では、web プロジェクト内のファイルを変更して、 **HelloProjectOData**アドインの機能を作成する方法を示します。

### <a name="procedure-2-to-create-the-helloprojectodata-add-in-for-project"></a>手順 2. Project 用の HelloProjectOData アドインを作成するには

1. 管理者として Visual Studio 2015 を実行し、スタートページで [**新しいプロジェクト**] を選択します。

2. [**新しいプロジェクト**] ダイアログボックスで、[**テンプレート**]、[ **Visual C#**]、[ **office/SharePoint** ] の各ノードを展開し、[* * Office アドイン * *] を選択します。中央のウィンドウの上部にある [ターゲットフレームワーク] ドロップダウンリストで [ **.Net framework 4.5.2** ] を選択し、[ **Office アドイン**] を選択します (次のスクリーンショットを参照)。

3. これらの Visual Studio プロジェクトを両方とも同じディレクトリに配置するには、[**ソリューションのディレクトリを作成**] を選択し、目的の場所を参照します。

4. [**名前**] フィールドに typeHelloProjectOData と入力し、[ **OK]** を選択します。

    *図 1. Office アドインの作成*

    ![Office アドインの作成](../images/pj15-hello-project-o-data-creating-app.png)

5. [**アドインの種類の選択**] ダイアログ ボックスで、[**作業ウィンドウ**] を選択して、[**次へ**] を選択します (次のスクリーンショットを参照)。

    *図 2. 作成するアドインの種類の選択*

    ![作成するアドインの種類の選択](../images/pj15-hello-project-o-data-choose-project.png)

6. [**ホスト アプリケーションの選択**] ダイアログ ボックスで、[**Project**] 以外のすべてのチェック ボックスをオフにし (次のスクリーンショットを参照)、[**完了**] をクリックします。

    *図 3. ホスト アプリケーションの選択*

    ![Project を唯一のホスト アプリケーションとして選択する](../images/create-office-add-in.png)

    Visual Studio によって、 **HelloProjectOdata**プロジェクトと**HelloProjectODataWeb**プロジェクトが作成されます。

[**AddIn**] フォルダー (次のスクリーンショットを参照) には、カスタム CSS スタイルの App.css ファイルが含まれています。 **ホーム** サブフォルダー内の、Home.html ファイルには、CSS ファイル、アドインを使用している JavaScript ファイル、およびアドインの HTML5 コンテンツへの参照が含まれています。 また、Home.js ファイルは、カスタムの JavaScript コード用です。 **Scripts** フォルダーには、jQuery ライブラリのファイルが含まれています。 **Office** サブフォルダーには、office.js や project-15.js などの JavaScript ライブラリ、および Office アドインでの標準の文字列用の言語ライブラリが含まれています。**コンテンツ** フォルダーで Office.css ファイルには、Office のアドインのすべてに使用する既定のスタイルが含まれます。

*図 4. ソリューション エクスプローラーでの既定の Web プロジェクト ファイルの表示*

![ソリューション エクスプローラーで Web プロジェクト ファイルを表示する](../images/pj15-hello-project-o-data-initial-solution-explorer.png)

**HelloProjectOData**プロジェクトのマニフェストは、HelloProjectOData ファイルです。必要に応じて、マニフェストを変更して、アドインの説明、アイコンへの参照、追加の言語の情報、その他の設定を追加できます。手順3は単にアドインの表示名と説明を変更し、アイコンを追加します。

マニフェストについて詳しくは、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」と「[Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md#see-also)」をご覧ください。

### <a name="procedure-3-to-modify-the-add-in-manifest"></a>手順 3. アドインのマニフェストを変更するには

1. Visual Studio で、HelloProjectOData.xml ファイルを開きます。

2. 既定の表示名は、Visual Studio プロジェクトの名前 ("HelloProjectOData") です。たとえば、 **DisplayName**要素の既定値を "Hello projectdata" に変更します。

3. 既定の説明も "HelloProjectOData" です。たとえば、Description 要素の既定値を "Test REST queries of the ProjectData service" に変更します。

4. リボンの [**プロジェクト**] タブにある [ **Office アドイン**] ドロップダウンリストに表示するアイコンを追加します。アイコンファイルは、Visual Studio ソリューションに追加することも、アイコンの URL を使用することもできます。 

以下の手順は、Visual Studio ソリューションにアイコン ファイルを追加するための方法を示しています。

1. **ソリューションエクスプローラー**で、Images という名前のフォルダーに移動します。

2. [ **Office アドイン**] ドロップダウンリストに表示するには、アイコンを 32 x 32 ピクセルにする必要があります。たとえば、Project 2013 SDK をインストールしてから、[ **Images** ] フォルダーを選択し、次のファイルを SDK から追加します。`\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`

    または、独自の 32 x 32 アイコンを使用するか、NewIcon.png という名前のファイルに次の画像をコピーして、`HelloProjectODataWeb\Images` フォルダーにそのファイルを追加します。

    ![HelloProjectOData アプリのアイコン](../images/pj15-hello-project-data-new-icon.jpg)

3. HelloProjectOData マニフェストで、アイコンの URL の値が32x32 のアイコンファイルへの相対パスである**Description**要素の下に**iconurl**要素を追加します。たとえば、次の行**<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />** を追加します。現在、HelloProjectOData マニフェストファイルには次のものが含まれています ( **Id**の値は異なります)。

    ```XML
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
        <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
        <Id>c512df8d-a1c5-4d74-8a34-d30f6bbcbd82</Id>
        <Version>1.0</Version>
        <ProviderName> [Provider name]</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Hello ProjectData" />
        <Description DefaultValue="Test REST queries of the ProjectData service"/>
        <IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />
        <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
        <Hosts>
            <Host Name="Project" />
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="~remoteAppUrl/AddIn/Home/Home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

## <a name="creating-the-html-content-for-the-helloprojectodata-add-in"></a>HelloProjectOData アドインの HTML コンテンツを作成する

**HelloProjectOData**アドインは、デバッグとエラー出力を含むサンプルです。これは、運用環境での使用を目的としたものではありません。HTML コンテンツのコーディングを開始する前に、アドインの UI とユーザーの操作手順を設計し、HTML コードを操作する JavaScript 関数の概要を説明します。詳細については、「[Office アドインの設計ガイドライン](../design/add-in-design.md)」を参照してください。 

作業ウィンドウには、上部にあるアドインの表示名が表示されます。これはマニフェスト内の**DisplayName**要素の値です。HelloProjectOData ファイルの**body**要素には、その他の UI 要素が次のように含まれています。

- サブタイトルは、**ODATA REST QUERY** など、操作の一般的な機能または種類を表します。

- [ **Projectdata エンドポイントの取得**] `setOdataUrl`ボタンをクリックすると、関数が呼び出されて**projectdata**サービスのエンドポイントが取得され、テキストボックスに表示されます。Project が Project Web App に接続されていない場合、アドインはエラーハンドラーを呼び出してポップアップエラーメッセージを表示します。

- アドインが有効な OData エンドポイントを取得するまで、[**すべてのプロジェクトを比較**] ボタンは無効になります。ボタンを選択すると、関数が呼び出さ`retrieveOData`れます。この関数は、REST クエリを使用して**projectdata**サービスからプロジェクトのコストと作業データを取得します。

- テーブルには、プロジェクト コスト、実績コスト、作業、および達成率の平均値が表示されます。このテーブルでは、現在アクティブなプロジェクトの値と平均値との比較も行われます。現在の値が全プロジェクトの平均値より大きい場合は、その値が赤で表示されます。現在の値が平均値より小さい場合は、その値が緑で表示されます。現在の値がない場合は、青の **NA** が表示されます。

    関数`retrieveOData`は、テーブル`parseODataResult`の値を計算して表示する関数を呼び出します。

    > [!NOTE]
    > この例では、作業中のプロジェクトのコストと作業時間データは、発行された値から派生します。Project で値を変更すると、プロジェクトが発行されるまで**Projectdata**サービスは変更されません。

### <a name="procedure-4-to-create-the-html-content"></a>手順 4. HTML コンテンツを作成するには

1. .Html ファイルの**head**要素に、アドインが使用する CSS ファイルの**リンク**要素を追加します。Visual Studio プロジェクトテンプレートには、カスタム CSS スタイルに使用できる App.xaml ファイルへのリンクが含まれています。

2. アドインが使用する JavaScript ライブラリの**スクリプト**要素を追加します。プロジェクトテンプレートには、 **Scripts**フォルダーにある jQuery- _[version]_.js、microsoftajax.js、およびファイルへのリンクが含まれています。

    > [!NOTE]
    > アドインを展開する前に、office.js の参照と jQuery の参照をコンテンツ配信ネットワーク (CDN) の参照に変更してください。CDN の参照は最新のバージョンと高いパフォーマンスを提供します。

    **HelloProjectOData**アドインも surfaceerrors.js ファイルを使用します。このファイルには、ポップアップメッセージにエラーが表示されます。[テキストエディターを使用して、「Project 2013 用の作業ウィンドウアドインを初めて作成](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)する」の「堅牢な_プログラミング_」セクションからコードをコピーし、 **HelloProjectODataWeb**プロジェクトの**スクリプト \ Office**フォルダーに surfaceerrors.js ファイルを追加できます。

    次に、Surfaceerrors.js ファイルの追加行を含む**head**要素の更新された HTML コードを示します。

    ```HTML
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Test ProjectData Service</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

    <!-- Add your CSS styles to the following file -->
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
    <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
    <script src="../Scripts/jquery-1.7.1.js"></script>

    <!-- Use the CDN reference to office.js when deploying your add-in. -->
    <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

    <!-- Use the local script references for Office.js to enable offline debugging -->
    <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/1.0/Office.js"></script>

    <!-- Add your JavaScript to the following files -->
    <script src="../Scripts/HelloProjectOData.js"></script>
    <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
    <!-- See the code in Step 3. -->
    </body>
    </html>
    ```

3. **Body**要素で、テンプレートから既存のコードを削除し、ユーザーインターフェイス用のコードを追加します。要素にデータを格納する場合、または jQuery ステートメントで操作する場合、要素には一意の**id**属性が含まれている必要があります。次のコードでは、jQuery 関数が使用する**ボタン**、 **span**、および**td** (table セル定義) 要素の**id**属性は、太字のフォントで表示されます。

   次の HTML は、会社のロゴなどのグラフィックスイメージを追加します。任意のロゴを使用するか、プロジェクト 2013 SDK のダウンロードから NewLogo .png ファイルをコピーしてから、**ソリューションエクスプローラー**を使用してそのファイルを`HelloProjectODataWeb\Images`フォルダーに追加できます。

    ```HTML
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br /><br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
                <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
            <table class="infoTable" aria-readonly="True" style="width: 100%;">
                <tr>
                    <td class="heading_leftCol"></td>
                    <td class="heading_midCol"><strong>Average</strong></td>
                    <td class="heading_rightCol"><strong>Current</strong></td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Work</strong></td>
                    <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project % Complete</strong></td>
                    <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
                </tr>
            </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
    ```

## <a name="creating-the-javascript-code-for-the-add-in"></a>このアドインの JavaScript コードを作成する

Project の作業ウィンドウアドインのテンプレートには、一般的な Office 2013 アドインのドキュメント内のデータに対する基本的な get および set アクションをデモンストレーションするために設計された既定の初期化コードが含まれています。Project 2013 では、作業中のプロジェクトに書き込む操作がサポートされておらず、 **HelloProjectOData**アドインはこの`getSelectedDataAsync`メソッドを使用していないため、 `Office.initialize`関数内のスクリプトを`setData`削除し`getData`て、既定の HelloProjectOData ファイル内の関数と関数を削除することができます。

JavaScript には、REST クエリのグローバル定数と、いくつかの関数で使用されるグローバル変数が含まれています。[ **ProjectData エンドポイントの取得**] `setOdataUrl`ボタンを呼び出して、グローバル変数を初期化し、Project が project Web App で接続されているかどうかを判断します。

HelloProjectOData ファイルの残りの部分には、次の2つ`retrieveOData`の関数が含まれています。この関数は、ユーザーが [**すべてのプロジェクトを比較**するを選択すると呼び出されます。関数は`parseODataResult`平均を計算し、次に、比較表に色と単位を書式設定した値を設定します。

### <a name="procedure-5-to-create-the-javascript-code"></a>手順 5. JavaScript コードを作成するには

1. 既定の HelloProjectOData ファイル内のすべてのコードを削除してから、グローバル変数と`**`Office initialize ' 関数を追加します。すべて大文字の変数名は定数であることを意味します。この例では、後で REST クエリを作成するために **_pwa**変数を使用しています。

    ```js
    var PROJDATA = "/_api/ProjectData";
    var PROJQUERY = "/Projects?";
    var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
    var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
    var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
    var _pwa;           // URL of Project Web App.
    var _projectUid;    // GUID of the active project.
    var _docUrl;        // Path of the project document.
    var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData

    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
        });
    }
    ```

2. 追加`setOdataUrl`および関連する関数。関数`setOdataUrl`は、 `getProjectGuid`グローバル`getDocumentUrl`変数を呼び出し、初期化します。[Getprojectfieldasync メソッド](/javascript/api/office/office.document)の場合、 _callback_ライブラリ`removeAttr`のメソッドを使用して [**すべてのプロジェクトを比較**] ボタンを有効にし、 **projectdata**サービスの URL を表示します。Project が Project Web App に接続されていない場合、この関数はエラーをスローします。これにより、ポップアップエラーメッセージが表示されます。Surfaceerrors.js ファイルには、 `throwError`メソッドが含まれています。

   > [!NOTE]
   > Project Server コンピュータで Visual Studio を実行している場合、 **F5**デバッグを使用するには、 **_pwa**グローバル変数を初期化する行の後にあるコードをコメント解除します。Project Server コンピュータでデバッグ`ajax`するときに jQuery メソッドを使用できるようにするには`localhost` 、PWA URL の値を設定する必要があります。Visual Studio をリモートコンピューターで実行する場合、 `localhost` URL は必須ではありません。アドインを展開する前に、そのコードをコメントアウトします。

    ```js
    function setOdataUrl() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _pwa = String(asyncResult.value.fieldValue);

                    // If you debug with Visual Studio on a local Project Server computer, 
                    // uncomment the following lines to use the localhost URL.
                    //var localhost = location.host.split(":", 1);
                    //var pwaStartPosition = _pwa.lastIndexOf("/");
                    //var pwaLength = _pwa.length - pwaStartPosition;
                    //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                    //_pwa = location.protocol + "//" + localhost + pwaName;

                    if (_pwa.substring(0, 4) == "http") {
                        _odataUrl = _pwa + PROJDATA;
                        $("#compareProjects").removeAttr("disabled");
                        getProjectGuid();
                    }
                    else {
                        _odataUrl = "No connection!";
                        throwError(_odataUrl, "You are not connected to Project Web App.");
                    }
                    getDocumentUrl();
                    $("#projectDataEndPoint").text(_odataUrl);
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }

    // Get the GUID of the active project.
    function getProjectGuid() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _projectUid = asyncResult.value.fieldValue;
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }

    // Get the path of the project in Project web app, which is in the form <>\ProjectName .
    function getDocumentUrl() {
        _docUrl = "Document path:\r\n" + Office.context.document.url;
    }
    ```

3. `retrieveOData`関数を追加します。この関数は、REST クエリの値を`ajax`連結し、jQuery の関数を呼び出して、 **projectdata**サービスから要求されたデータを取得します。サポートされている**cors**変数は、 `ajax`関数との間でのクロスオリジンリソース共有 (cors) を有効にします。サポートされている**cors**ステートメントが存在しない場合、または`ajax` **false**に設定されている場合、この関数は**No transport**エラーを返します。

   > [!NOTE]
   > 次に示すコードは、Project Server 2013 のオンプレミスのインストールで動作します。Project on the web の場合は、トークン ベースの認証に OAuth を使用できます。詳細については、「[Office アドインにおける同一生成元ポリシーの制限への対処](../develop/addressing-same-origin-policy-limitations.md)」を参照してください。

   `ajax`呼び出しでは、 _headers_パラメーターまたは_beforesend_パラメーターのいずれかを使用できます。_Complete_パラメーターは、の`retrieveOData`変数と同じスコープにある匿名関数です。_Complete_パラメーターの関数は、 `odataText`コントロールに結果を表示し、JSON 応答`parseODataResult`を解析して表示するメソッドも呼び出します。_Error_パラメーターは、指定さ`getProjectDataErrorHandler`れた関数を指定します。この`odataText`関数は、エラーメッセージ`throwError`をコントロールに書き込み、また、このメソッドを使用してポップアップメッセージを表示します。

    ```js
    // Functions to get and parse the Project Server reporting data./

    // Get data about all projects on Project Server,
    // by using a REST query with the ajax method in jQuery.
    function retrieveOData() {
        var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
        var accept = "application/json; odata=verbose";
        accept.toLocaleLowerCase();

        // Enable cross-origin scripting (required by jQuery 1.5 and later).
        // This does not work with Project on the web.
        $.support.cors = true;

        $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json",
            data: "",      // Empty string for the optional data.
            //headers: { "Accept": accept },
            beforeSend: function (xhr) {
                xhr.setRequestHeader("ACCEPT", accept);
            },
            complete: function (xhr, textStatus) {
                // Create a message to display in the text box.
                var message = "\r\ntextStatus: " + textStatus +
                    "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                    "\r\nStatus: " + xhr.status +
                    "\r\nResponseText:\r\n" + xhr.responseText;

                // xhr.responseText is the result from an XmlHttpRequest, which
                // contains the JSON response from the OData service.
                parseODataResult(xhr.responseText, _projectUid);

                // Write the document name, response header, status, and JSON to the odataText control.
                $("#odataText").text(_docUrl);
                $("#odataText").append("\r\nREST query:\r\n" + restUrl);
                $("#odataText").append(message);

                if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                    $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
                }
            },
            error: getProjectDataErrorHandler
        });
    }

    function getProjectDataErrorHandler(data, errorCode, errorMessage) {
        $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
        throwError(errorCode, errorMessage);
    }
    ```

4. このメソッド`parseODataResult`を追加します。このメソッドは、OData サービスから JSON 応答を逆シリアル化し、処理します。この`parseODataResult`メソッドは、コストと作業時間データの平均値を1桁または2桁の小数点以下の桁数で計算し、値を正しい**$** 色で書式設定**%** し、単位 (、**時間**、または) を追加して、指定された表のセルの値を表示します。

   アクティブなプロジェクトの GUID が`ProjectId`値と一致する場合、 `myProjectIndex`変数はプロジェクトインデックスに設定されます。作業`myProjectIndex`中のプロジェクトが project Server で発行された`parseODataResult`場合、このメソッドはそのプロジェクトのコストと作業時間データを書式設定して表示します。作業中のプロジェクトが発行されていない場合、作業中のプロジェクトの値は青色の**NA**として表示されます。

    ```js
    // Calculate the average values of actual cost, cost, work, and percent complete
    // for all projects, and compare with the values for the current project.
    function parseODataResult(oDataResult, currentProjectGuid) {
        // Deserialize the JSON string into a JavaScript object.
        var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
        var len = res.d.results.length;
        var projActualCost = 0;
        var projCost = 0;
        var projWork = 0;
        var projPercentCompleted = 0;
        var myProjectIndex = -1;
        for (i = 0; i < len; i++) {
            // If the current project GUID matches the GUID from the OData query,  
            // store the project index.
            if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
                myProjectIndex = i;
            }
            projCost += Number(res.d.results[i].ProjectCost);
            projWork += Number(res.d.results[i].ProjectWork);
            projActualCost += Number(res.d.results[i].ProjectActualCost);
            projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);
        }
        var avgProjCost = projCost / len;
        var avgProjWork = projWork / len;
        var avgProjActualCost = projActualCost / len;
        var avgProjPercentCompleted = projPercentCompleted / len;

        // Round off cost to two decimal places, and round off other values to one decimal place.
        avgProjCost = avgProjCost.toFixed(2);
        avgProjWork = avgProjWork.toFixed(1);
        avgProjActualCost = avgProjActualCost.toFixed(2);
        avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);

        // Display averages in the table, with the correct units.
        document.getElementById("AverageProjectCost").innerHTML = "$"
            + avgProjCost;
        document.getElementById("AverageProjectActualCost").innerHTML
            = "$" + avgProjActualCost;
        document.getElementById("AverageProjectWork").innerHTML
            = avgProjWork + " hrs";
        document.getElementById("AverageProjectPercentComplete").innerHTML
            = avgProjPercentCompleted + "%";

        // Calculate and display values for the current project.
        if (myProjectIndex != -1) {
            var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
            var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
            var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
            var myProjPercentCompleted =
            Number(res.d.results[myProjectIndex].ProjectPercentCompleted);

            myProjCost = myProjCost.toFixed(2);
            myProjWork = myProjWork.toFixed(1);
            myProjActualCost = myProjActualCost.toFixed(2);
            myProjPercentCompleted = myProjPercentCompleted.toFixed(1);

            document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;

            if (Number(myProjCost) <= Number(avgProjCost)) {
                document.getElementById("CurrentProjectCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectCost").style.color = "red"
            }

            document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;

            if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
                document.getElementById("CurrentProjectActualCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectActualCost").style.color = "red"
            }

            document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";

            if (Number(myProjWork) <= Number(avgProjWork)) {
                document.getElementById("CurrentProjectWork").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectWork").style.color = "green"
            }

            document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";

            if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
                document.getElementById("CurrentProjectPercentComplete").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectPercentComplete").style.color = "green"
            }
        }
        else {
            document.getElementById("CurrentProjectCost").innerHTML = "NA";
            document.getElementById("CurrentProjectCost").style.color = "blue"

            document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
            document.getElementById("CurrentProjectActualCost").style.color = "blue"

            document.getElementById("CurrentProjectWork").innerHTML = "NA";
            document.getElementById("CurrentProjectWork").style.color = "blue"

            document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
            document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
        }
    }
    ```

## <a name="testing-the-helloprojectodata-add-in"></a>HelloProjectOData アドインのテスト

Visual Studio 2015 を使用して**HelloProjectOData**アドインをテストおよびデバッグするには、開発用コンピューターに Project Professional 2013 がインストールされている必要があります。異なるテストシナリオを有効にするには、ローカルコンピューター上のファイルに対してプロジェクトを開くか、Project Web App を使用して接続するかを選択できるようにします。たとえば、次の手順を実行します。

1. リボンの [**ファイル**] タブで、Backstage ビューの [**情報**] タブを選択し、[**アカウントの管理**] を選択します。

2. [ **Project web App アカウント**] ダイアログボックスの [**利用可能なアカウント**] リストには、ローカル**コンピューター**アカウントに加えて、複数の Project web app アカウントを含めることができます。[**開始時**] セクションで、[**アカウントの選択**] を選択します。

3. Project を閉じます。それにより、Visual Studio でアドインのデバッグ用に Project 起動できるようになります。

基本的なテストでは、次のことを行う必要があります。

- Visual Studio からアドインを実行し、コストと作業データを含む Project Web App から発行済みのプロジェクトを開きます。アドインに**projectdata**エンドポイントが表示され、テーブル内のコストと作業データが正しく表示されることを確認します。**Odatatext**コントロールの出力を使用して、REST クエリやその他の情報を確認できます。

- プロジェクトが開始されたら、[**ログイン**] ダイアログボックスでローカルコンピュータープロファイルを選択して、アドインを再度実行します。ローカルの .mpp ファイルを開き、アドインをテストします。 **Projectdata**エンドポイントを取得しようとしたときに、アドインにエラーメッセージが表示されることを確認します。

- アドインを再度実行して、コストと作業データを含むタスクを含むプロジェクトを作成します。プロジェクトを Project Web App に保存することはできますが、発行することはできません。アドインに Project Server のデータが表示されることを確認します。ただし、現在のプロジェクトの場合は**NA**となります。

### <a name="procedure-6-to-test-the-add-in"></a>手順 6. アドインをテストするには

1. Project Professional 2013 を実行し、Project Web App と接続してから、テスト プロジェクトを作成します。ローカル リソースまたはエンタープライズ リソースにタスクを割り当て、いくつかのタスクに対して達成率のさまざまな値を設定してから、そのプロジェクトを発行します。Project を終了します。それにより、Visual Studio がアドインのデバッグ用に Project を起動できるようになります。

2. Visual Studio で、 **F5**キーを押します。Project Web App にログオンし、前の手順で作成したプロジェクトを開きます。プロジェクトは、読み取り専用モードまたは編集モードで開くことができます。

3. リボンの [**プロジェクト**] タブの [ **Office アドイン**] ドロップダウンリストで、[ **Hello projectdata** ] を選択します (図5を参照)。[**すべてのプロジェクトを比較**] ボタンを無効にする必要があります。

    *図 5. HelloProjectOData アドインの開始*

    ![HelloProjectOData アプリのテスト](../images/pj15-hello-project-data-test-the-app.png)

4. [**プロジェクトデータの Hello** ] 作業ウィンドウで、[ **Projectdata エンドポイントの取得**] を選択します。**Projectdataendpoint**行に**PROJECTDATA**サービスの URL が表示され、[**すべてのプロジェクトを比較**] ボタンが有効になっている必要があります (図6を参照)。

5. [**すべてのプロジェクトを比較**] を選択します。アドインは、 **projectdata**サービスからデータを取得するときに一時停止する可能性があります。その後、書式設定された平均値と現在の値を表に表示する必要があります。

    *図 6. REST クエリの結果の表示*

    ![REST クエリの結果の表示](../images/pj15-hello-project-data-rest-results.png)

6. テキストボックス内の出力を調べます。これは、 **ajax**および**parseodataresult**への呼び出しからのドキュメントパス、REST クエリ、ステータス情報、および JSON 結果を表示する必要があります。出力は、などの`parseODataResult`メソッドでコードを理解、作成、デバッグするの`projCost += Number(res.d.results[i].ProjectCost);`に便利です。

    次に示すのは、Project Web App インスタンスの 3 つのプロジェクトの出力例です。わかりやすくするためにテキストに改行と空白を追加してあります。

    ```json
    Document path: <>\WinProj test1

    REST query:
    http://sphvm-37189/pwa/_api/ProjectData/Projects?$filter=ProjectName ne 'Timesheet Administrative Work Items'
        &amp;$select=ProjectId, ProjectName, ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost

    textStatus: success
    ContentType: application/json;odata=verbose;charset=utf-8
    Status: 200

    ResponseText:
    {"d":{"results":[
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "type":"ReportingData.Project"},
        "ProjectId":"ce3d0d65-3904-e211-96cd-00155d157123",
        "ProjectActualCost":"0.000000",
        "ProjectCost":"0.000000",
        "ProjectName":"Task list created in PWA",
        "ProjectPercentCompleted":0,
        "ProjectWork":"16.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"c31023fc-1404-e211-86b2-3c075433b7bd",
        "ProjectActualCost":"700.000000",
        "ProjectCost":"2400.000000",
        "ProjectName":"WinProj test 2",
        "ProjectPercentCompleted":29,
        "ProjectWork":"48.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"dc81fbb2-b801-e211-9d2a-3c075433b7bd",
        "ProjectActualCost":"1900.000000",
        "ProjectCost":"5200.000000",
        "ProjectName":"WinProj test1",
        "ProjectPercentCompleted":37,
        "ProjectWork":"104.000000"}
    ]}}
    ```

7. デバッグを停止し (Shift キーを押し**ながら f5**キーを押し)、もう一度**F5**キーを押して Project の新しいインスタンスを実行します。[**ログイン**] ダイアログボックスで、[Project Web App] ではなく [ローカル**コンピューター** ] プロファイルを選択します。ローカルのプロジェクト .mpp ファイルを作成または開き、[ **Hello projectdata** ] 作業ウィンドウを開き、[ **projectdata エンドポイントの取得**] を選択します。アドインは [**接続しない**] を表示する必要があります。エラー (図7を参照)。 [すべての**プロジェクトを比較**] ボタンは無効のままにしておきます。

   *図 7. Project Web App 接続がない状態でのアドインの使用*

   ![Project Web App 接続がない状態でのアプリの使用](../images/pj15-hello-project-data-no-connection.png)

8. デバッグを停止してから、もう一度**F5**キーを押します。Project Web App にログオンし、コストと作業データを含むプロジェクトを作成します。プロジェクトを保存することはできますが、発行することはできません。

   [ **ProjectData の Hello** ] 作業ウィンドウで、[**すべてのプロジェクトを比較**] を選択すると、**現在**の列のフィールドに青い**NA**が表示されます (図8を参照)。

   *図 8. 未発行のプロジェクトと他のプロジェクトの比較*

   ![未発行のプロジェクトと他のプロジェクトの比較](../images/pj15-hello-project-data-not-published.png)

アドインがこれまでのテストで正常に動作しているとしても、実行する必要のあるテストはまだあります。たとえば、次のことを行います。

- タスクのコストまたは作業時間データを持たないプロジェクトを Project Web App から開きます。**現在**の列のフィールドに0の値が表示されます。

- タスクがないプロジェクトをテストします。

- アドインに変更を加えて発行した場合は、発行したアドインで再び同様のテストを実行する必要があります。その他の考慮事項については、「[次のステップ](#next-steps)」を参照してください。

> [!NOTE]
> **Projectdata**サービスの1つのクエリで返されるデータの量には制限があります。データの量は、エンティティによって異なります。たとえば、 `Projects`エンティティセットの既定の制限は、クエリごとに100プロジェクトですが、 `Risks`エンティティセットの既定の制限は200です。運用環境のインストールでは、 **HelloProjectOData**例のコードを変更して、100を超えるプロジェクトのクエリを有効にする必要があります。詳細については、「[プロジェクトレポートデータの OData フィードを照会する](/previous-versions/office/project-odata/jj163048(v=office.15))」と「クエリを実行する[方法](#next-steps)」を参照してください。

## <a name="example-code-for-the-helloprojectodata-add-in"></a>HelloProjectOData アドインのコード例

### <a name="helloprojectodatahtml-file"></a>HelloProjectOData.html ファイル

次のコードは、**HelloProjectODataWeb** プロジェクトの `Pages\HelloProjectOData.html` ファイルに収められています。

```HTML
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Test ProjectData Service</title>

        <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

        <!-- Add your CSS styles to the following file -->
        <link rel="stylesheet" type="text/css" href="../Content/App.css" />

        <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
        <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
        <script src="../Scripts/jquery-1.7.1.js"></script>

        <!-- Use the CDN reference to Office.js when deploying your add-in -->
        <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

        <!-- Use the local script references for Office.js to enable offline debugging -->
        <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
        <script src="../Scripts/Office/1.0/Office.js"></script>

        <!-- Add your JavaScript to the following files -->
        <script src="../Scripts/HelloProjectOData.js"></script>
        <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br />
            <br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
            <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">
            Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
        <table class="infoTable" aria-readonly="True" style="width: 100%;">
            <tr>
            <td class="heading_leftCol"></td>
            <td class="heading_midCol"><strong>Average</strong></td>
            <td class="heading_rightCol"><strong>Current</strong></td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Cost</strong></td>
            <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
            <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Work</strong></td>
            <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project % Complete</strong></td>
            <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
            </tr>
        </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
</html>
```

### <a name="helloprojectodatajs-file"></a>HelloProjectOData.js ファイル

次のコードは、**HelloProjectODataWeb** プロジェクトの `Scripts\Office\HelloProjectOData.js` ファイルに収められています。

```js
/* File: HelloProjectOData.js
* JavaScript functions for the HelloProjectOData example task pane app.
* October 2, 2012
*/

var PROJDATA = "/_api/ProjectData";
var PROJQUERY = "/Projects?";
var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
var _pwa;           // URL of Project Web App.
var _projectUid;    // GUID of the active project.
var _docUrl;        // Path of the project document.
var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
    });
}

// Set the global variables, enable the Compare All Projects button,
// and display the URL of the ProjectData service.
// Display an error if Project is not connected with Project Web App.
function setOdataUrl() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.ProjectServerUrl,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _pwa = String(asyncResult.value.fieldValue);

                // If you debug with Visual Studio on a local Project Server computer,
                // uncomment the following lines to use the localhost URL.
                //var localhost = location.host.split(":", 1);
                //var pwaStartPosition = _pwa.lastIndexOf("/");
                //var pwaLength = _pwa.length - pwaStartPosition;
                //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                //_pwa = location.protocol + "//" + localhost + pwaName;

                if (_pwa.substring(0, 4) == "http") {
                    _odataUrl = _pwa + PROJDATA;
                    $("#compareProjects").removeAttr("disabled");
                    getProjectGuid();
                }
                else {
                    _odataUrl = "No connection!";
                    throwError(_odataUrl, "You are not connected to Project Web App.");
                }
                getDocumentUrl();
                $("#projectDataEndPoint").text(_odataUrl);
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the GUID of the active project.
function getProjectGuid() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.GUID,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _projectUid = asyncResult.value.fieldValue;
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the path of the project in Project web app, which is in the form <>\ProjectName .
function getDocumentUrl() {
    _docUrl = "Document path:\r\n" + Office.context.document.url;
}

//  Functions to get and parse the Project Server reporting data./

// Get data about all projects on Project Server,
// by using a REST query with the ajax method in jQuery.
function retrieveOData() {
    var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
    var accept = "application/json; odata=verbose";
    accept.toLocaleLowerCase();

    // Enable cross-origin scripting (required by jQuery 1.5 and later).
    // This does not work with Project on the web.
    $.support.cors = true;

    $.ajax({
        url: restUrl,
        type: "GET",
        contentType: "application/json",
        data: "",      // Empty string for the optional data.
        //headers: { "Accept": accept },
        beforeSend: function (xhr) {
            xhr.setRequestHeader("ACCEPT", accept);
        },
        complete: function (xhr, textStatus) {
            // Create a message to display in the text box.
            var message = "\r\ntextStatus: " + textStatus +
                "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                "\r\nStatus: " + xhr.status +
                "\r\nResponseText:\r\n" + xhr.responseText;

            // xhr.responseText is the result from an XmlHttpRequest, which 
            // contains the JSON response from the OData service.
            parseODataResult(xhr.responseText, _projectUid);

            // Write the document name, response header, status, and JSON to the odataText control.
            $("#odataText").text(_docUrl);
            $("#odataText").append("\r\nREST query:\r\n" + restUrl);
            $("#odataText").append(message);

            if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
            }
        },
        error: getProjectDataErrorHandler
    });
}

function getProjectDataErrorHandler(data, errorCode, errorMessage) {
    $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
    throwError(errorCode, errorMessage);
}

// Calculate the average values of actual cost, cost, work, and percent complete
// for all projects, and compare with the values for the current project.
function parseODataResult(oDataResult, currentProjectGuid) {
    // Deserialize the JSON string into a JavaScript object.
    var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
    var len = res.d.results.length;
    var projActualCost = 0;
    var projCost = 0;
    var projWork = 0;
    var projPercentCompleted = 0;
    var myProjectIndex = -1;

    for (i = 0; i < len; i++) {
        // If the current project GUID matches the GUID from the OData query,  
        // then store the project index.
        if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
            myProjectIndex = i;
        }
        projCost += Number(res.d.results[i].ProjectCost);
        projWork += Number(res.d.results[i].ProjectWork);
        projActualCost += Number(res.d.results[i].ProjectActualCost);
        projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);

    }
    var avgProjCost = projCost / len;
    var avgProjWork = projWork / len;
    var avgProjActualCost = projActualCost / len;
    var avgProjPercentCompleted = projPercentCompleted / len;

    // Round off cost to two decimal places, and round off other values to one decimal place.
    avgProjCost = avgProjCost.toFixed(2);
    avgProjWork = avgProjWork.toFixed(1);
    avgProjActualCost = avgProjActualCost.toFixed(2);
    avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);

    // Display averages in the table, with the correct units. 
    document.getElementById("AverageProjectCost").innerHTML = "$"
        + avgProjCost;
    document.getElementById("AverageProjectActualCost").innerHTML
        = "$" + avgProjActualCost;
    document.getElementById("AverageProjectWork").innerHTML
        = avgProjWork + " hrs";
    document.getElementById("AverageProjectPercentComplete").innerHTML
        = avgProjPercentCompleted + "%";

    // Calculate and display values for the current project.
    if (myProjectIndex != -1) {

        var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
        var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
        var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
        var myProjPercentCompleted = Number(res.d.results[myProjectIndex].ProjectPercentCompleted);

        myProjCost = myProjCost.toFixed(2);
        myProjWork = myProjWork.toFixed(1);
        myProjActualCost = myProjActualCost.toFixed(2);
        myProjPercentCompleted = myProjPercentCompleted.toFixed(1);

        document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;

        if (Number(myProjCost) <= Number(avgProjCost)) {
            document.getElementById("CurrentProjectCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectCost").style.color = "red"
        }

        document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;

        if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
            document.getElementById("CurrentProjectActualCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectActualCost").style.color = "red"
        }

        document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";

        if (Number(myProjWork) <= Number(avgProjWork)) {
            document.getElementById("CurrentProjectWork").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectWork").style.color = "green"
        }

        document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";

        if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
            document.getElementById("CurrentProjectPercentComplete").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectPercentComplete").style.color = "green"
        }
    }
    else {    // The current project is not published.
        document.getElementById("CurrentProjectCost").innerHTML = "NA";
        document.getElementById("CurrentProjectCost").style.color = "blue"

        document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
        document.getElementById("CurrentProjectActualCost").style.color = "blue"

        document.getElementById("CurrentProjectWork").innerHTML = "NA";
        document.getElementById("CurrentProjectWork").style.color = "blue"

        document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
        document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
    }
}
```

### <a name="appcss-file"></a>App.css ファイル

次のコードは、**HelloProjectODataWeb** プロジェクトの `Content\App.css` ファイルに収められています。

```css
/*
*  File: App.css for the HelloProjectOData app.
*  Updated: 10/2/2012
*/

body
{
    font-size: 11pt;
}
h1
{
    font-size: 22pt;
}
h2
{
    font-size: 16pt;
}

/******************************************************************
Code label class
******************************************************************/

.rest 
{
    font-family: 'Courier New';
    font-size: 0.9em;
}

/******************************************************************
Button classes
******************************************************************/

.button-wide {
    width: 210px;
    margin-top: 2px;
}
.button-narrow 
{
    width: 80px;
    margin-top: 2px;
}

/******************************************************************
Table styles
******************************************************************/

.infoTable
{
    text-align: center; 
    vertical-align: middle
}
.heading_leftCol
{
    width: 20px;
    height: 20px;
}
.heading_midCol
{
    width: 100px;
    height: 20px;
    font-size: medium; 
    font-weight: bold; 
}
.heading_rightCol
{
    width: 101px;
    height: 20px;
    font-size: medium;
    font-weight: bold;
}
.row_leftCol
{
    width: 20px;
    font-size: small;
    font-weight: bold;
}
.row_midCol
{
    width: 100px;
}
.row_rightCol
{
    width: 101px;
}
.logo
{
    width: 135px;
    height: 53px;
}
```

### <a name="surfaceerrorsjs-file"></a>SurfaceErrors.js ファイル

SurfaceErrors.js ファイルのコードは、「[テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)」の「_堅牢なプログラミング_」セクションからコピーできます。

## <a name="next-steps"></a>次の手順

**HelloProjectOData**が appsource で販売される、または SharePoint アプリカタログで配布される運用アドインの場合、設計方法は異なります。たとえば、テキストボックスにデバッグ出力はありません。また、 **Projectdata**エンドポイントを取得するためのボタンがないことがあります。また、 `retireveOData` 100 個を超えるプロジェクトを含む Project Web App インスタンスを処理するように関数を書き換える必要があります。

このアドインには、追加のエラー チェックと、エッジ ケースをキャッチして説明または表示するためのロジックを組み込む必要があります。たとえば、Project Web App インスタンスに、平均期間が 5 日で平均コストが $2400 になる 1000 個のプロジェクトがあって、期間が 20 日より長いのはアクティブ プロジェクトだけだとすると、コストと作業の比較は歪んだものになるでしょう。それは頻度グラフで示すことができます。期間を表示したり、同じような長さのプロジェクトを比較したり、同じ部門または異なる部門のプロジェクトを比較したりするオプションを追加するとよいでしょう。あるいは、表示するフィールドのリストからユーザーが選択できるような方法を追加することもできます。

**Projectdata**サービスのその他のクエリの場合、クエリ文字列の長さに制限があり、クエリが親コレクションから子コレクション内のオブジェクトに対して実行できる手順の数に影響します。たとえば、タスクに対する**プロジェクト**の2段階のクエリ**をタスク**アイテムに対して実行することはできますが **、割り当てアイテムへの****タスク**への**プロジェクト**などの3段階のクエリは、既定の最大 URL の長さを超える場合があります。詳細については、「[プロジェクトレポートデータの OData フィードを照会する](/previous-versions/office/project-odata/jj163048(v=office.15))」を参照してください。

**HelloProjectOData**アドインを運用環境で使用するように変更する場合は、次の手順を実行します。

- HelloProjectOData.html ファイルで、パフォーマンスを向上させるために、office.js の参照をローカル プロジェクトから CDN の参照に変更します。

    ```HTML
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

- 100を`retrieveOData`超えるプロジェクトのクエリを有効にするように関数を書き直します。たとえば、 `~/ProjectData/Projects()/$count`クエリを使用してプロジェクトの数を取得し、プロジェクトデータの REST クエリで _$skip_演算子と _$top_演算子を使用することができます。ループで複数のクエリを実行してから、各クエリのデータの平均値を計算します。プロジェクトデータの各クエリは次の形式になります。 

  `~/ProjectData/Projects()?skip= [numSkipped]&amp;$top=100&amp;$filter=[filter]&amp;$select=[field1,field2, ???????]`

  For more information, see [OData System Query Options Using the REST Endpoint](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/sharepoint-server/Set-SPProjectOdataConfiguration?view=sharepoint-ps) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).

- アドインを展開するには、「[Office アドインを発行する](../publish/publish.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Project 用の作業ウィンドウ アドイン](project-add-ins.md)
- [テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- [ProjectData - Project OData サービス リファレンス](/previous-versions/office/project-odata/jj163015(v=office.15))
- [Office アドインの XML マニフェスト](../develop/add-in-manifests.md)
- [Office アドインを発行する](../publish/publish.md)
