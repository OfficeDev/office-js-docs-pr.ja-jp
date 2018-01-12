
# <a name="task-pane-add-ins-for-project"></a>Project 用の作業ウィンドウ アドイン


Project Standard 2013 と Project Professional 2013 はどちらも作業ウィンドウ アドインに対応しています。Word 2013 または Excel 2013 用に開発された一般的な作業ウィンドウ アドインを実行できます。また、カスタム アドインを開発して、Project の一部のイベントを処理し、プロジェクトのタスク、リソース、ビュー、およびその他のセルレベルのデータを、SharePoint のリスト、SharePoint アドイン、Web パーツ、Web サービス、およびエンタープライズ アプリケーションに統合することもできます。

 >**メモ**[Project 2013 SDK のダウンロード](https://www.microsoft.com/en-us/download/details.aspx?id=30435%20)には、Project のアドイン オブジェクト モデルの使用方法と、Project Server 2013 のレポート データ用 OData サービスの使用方法を示すサンプル アドインが含まれています。SDK を展開してインストールしたら、 `\Samples\Apps\` サブディレクトリを確認します。

Office アドインの概要については、「[Office アドイン プラットフォームの概要](../../docs/overview/office-add-ins.md)」を参照してください。

## <a name="add-in-scenarios-for-project"></a>Project 用のアドインのシナリオ


プロジェクト管理者は、Project の作業ウィンドウ アドインを使用して、プロジェクトの管理作業を円滑に進めることができます。よく使用する情報を調べるときに、Project から離れて別のアプリケーションを起動する必要がなく、Project 内で情報に直接アクセスできます。作業ウィンドウ アドインはコンテキストに応じたコンテンツを表示でき、選択中のタスク、リソース、またはビューに基づくコンテンツや、ガント チャート、タスク使用状況ビュー、またはリソース使用状況ビューのセルに含まれているその他のデータに基づくコンテンツを使用できます。


 >**メモ**  Project Professional 2013 では、社内インストールの Project Server 2013、Project Online、および社内またはオンラインの SharePoint 2013 にアクセスする作業ウィンドウ アドインを開発できます。Project Standard 2013では、Project Server データ、または Project Server と同期している SharePoint タスク リストとの直接の統合をサポートしていません。

Project 用のアドイン シナリオとして、次のようなものがあります。


-  **プロジェクトのスケジュール**???関連性がありスケジュールに影響するプロジェクトのデータを表示できます。作業ウィンドウ アドインでは、Project Server 2013 の他のプロジェクトから、関連するデータを統合できます。たとえば、部署ごとのプロジェクトとマイルストーンの日付の一覧を確認したり、選択したカスタム フィールドに基づいて他のプロジェクトの特定のデータを参照したりできます。
    
-  **リソース管理**???Project Server 2013のリソース共有元の全データや、特定のスキルに基づく一部のデータを、コスト データやリソースの使用可能時間を含めて確認でき、適切なリソースの選択に役立ちます。
    
-  **状況と承認**???作業ウィンドウ アドインの Web アプリケーションを使用して、外部のエンタープライズ リソース プラニング (ERP) アプリケーション、タイムシート システム、または会計アプリケーションのデータを更新または参照できます。または、Project Web App と Project Professional 2013の両方で使用できるカスタムの状況承認 Web パーツを作成します。
    
-  **チームのコミュニケーション**???プロジェクトのコンテキスト内で、作業ウィンドウ アドインからチームのメンバーやリソースと直接コミュニケーションを図ることができます。または、プロジェクトに従事する中で、コンテキストに応じた自分用のメモを簡単に保持できます。
    
-  **作業用のパッケージ**???SharePoint ライブラリやオンラインのテンプレート コレクションから、特別な種類のプロジェクト テンプレートを検索できます。たとえば、建設プロジェクト用のテンプレートを見つけ、自分の Project テンプレート コレクションに追加できます。
    
-  **関連アイテム**???プロジェクト計画の特定のタスクに関連するメタデータ、ドキュメント、およびメッセージを参照できます。たとえば、Project Professional 2013 で SharePoint タスク リストからインポートしたプロジェクトを管理し、プロジェクトに加えた変更にタスク リストを同期できます。作業ウィンドウ アドインを使用して、SharePoint リストのタスクに関して Project がインポートしなかった追加のフィールドやメタデータを表示できます。
    
-  **Project Server のオブジェクト モデルの使用**???選択したタスクの GUID を、Project Server Interface (PSI) または Project Server のクライアント側オブジェクト モデル (CSOM) のメソッドで使用できます。たとえば、アドインの Web アプリケーションで、選択したタスクとリソースの状況データの読み取りや更新を行ったり、外部のタイムシート アプリケーションと統合したりできます。
    
-  **レポート データの取得**???Representational State Transfer (REST)、JavaScript、または LINQ クエリを使用して、Project Web App のレポート テーブル用 OData サービスで、選択したタスクまたはリソースに関連する情報を検索します。OData サービスを使用するクエリは、Project Server 2013のオンライン インストールまたは社内インストールで実行できます。
    
    「[社内の Project Server OData サービスで REST を使用する Project アドインを作成する](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)」などを参照してください。
    

## <a name="developing-project-add-ins"></a>Project アドインの開発


Project アドイン の JavaScript ライブラリには、 **Office** 名前空間エイリアスの拡張機能が含まれています。開発者は、これらの拡張機能を使用して、プロジェクト内で Project アプリケーションのプロパティとタスク、リソース、およびビューにアクセスできます。Project-15.js ファイルに含まれている JavaScript ライブラリの拡張機能は、Visual Studio 2015 で作成された Project アドインで使用されます。Office.js、Office.debug.js、Project-15.js、Project-15.debug.js、および関連ファイルも、Project 2013 SDK ダウンロードで提供されます。

アドインを作成するには、基本的なテキスト エディターを使用して、HTML の Web ページと、関連する JavaScript ファイル、CSS ファイル、および REST クエリを作成します。アドインには、HTML ページや Web アプリケーションに加えて、構成用の XML マニフェスト ファイルも必要です。Project では、 **type** 属性に **TaskPaneExtension** を指定したマニフェスト ファイルを使用できます。同じマニフェスト ファイルを複数の Office 2013 クライアント アプリケーションで使用することも、Project 2013 専用のマニフェスト ファイルを作成することもできます。詳細については、「 _Office アドイン プラットフォームの概要_」の「 [開発の基本](../../docs/overview/office-add-ins.md) 」セクションを参照してください。

複雑なカスタム アプリケーションを作成する場合、デバッグを容易にするためには、Visual Studio 2015 を使用してアドイン用の Web サイトを開発することをお勧めします。Visual Studio 2015 にはアドイン プロジェクト用のテンプレートが含まれており、アドインの種類 (作業ウィンドウ、コンテンツ、またはメール) とホスト アプリケーション (Project、Word、Excel、または Outlook) を選択できます。Project Online からのデータと統合する例については、MSDN の Project プログラミングのブログの「[Project 作業ウィンドウ アドインを PWA に接続する](http://blogs.msdn.com/b/project_programmability/archive/2012/11/02/connecting-a-project-task-pane-app-to-pwa.aspx)」を参照してください。

Project 2013 SDK のダウンロード ファイルをインストールすると、`\Samples\Apps\` サブディレクトリに以下のサンプル アドインが置かれます。


-  **Bing Search:**??BingSearch.xml マニフェスト ファイルでは、モバイル デバイス用の Bing の検索ページが指定されています。Bing の Web アドインは既にインターネット上に存在するため、Bing Search アドインでは他のソース コード ファイルや Project 用のアドイン オブジェクト モデルは使用しません。
    
-  **Project OM Test:**??JSOM_SimpleOMCalls.xml マニフェスト ファイルと JSOM_Call.html ファイルのペアは、オブジェクト モデルとアドインの機能を Project 2013でテストするサンプルです。HTML ファイルでは JSOM_Sample.js ファイルを参照しています。この中には JavaScript 関数が記述されており、Office.js ファイルと Project-15.js ファイルを使用して主な機能を実現しています。Project OM Test アドインに必要なソース コード ファイルとマニフェスト XML ファイルは SDK のダウンロード ファイルにすべて含まれています。Project OM Test サンプルの開発とインストールについては、「 [テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)」で説明しています。
    
-  **HelloProject_OData:**??これは Project Professional 2013 用の Visual Studio ソリューションです。アクティブ プロジェクトのコスト、作業時間、達成率などのデータを示し、アクティブ プロジェクトが格納されている Project Web App インスタンス内のすべての発行済みプロジェクトの平均と比較します。サンプル (Project Web App の  **ProjectData** サービスと REST プロトコルを使用) の開発、インストール、およびテストについては、「 [社内の Project Server OData サービスで REST を使用する Project アドインを作成する](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)」を参照してください。
    

### <a name="creating-an-add-in-manifest-file"></a>アドインのマニフェスト ファイルの作成


マニフェスト ファイルでは、アドインの Web ページまたは Web アプリケーションの URL、アドインの種類 (Project 用の作業ウィンドウ アドイン)、他の言語とロケール用のコンテンツを表すオプションの URL、およびその他のプロパティを指定します。


### <a name="procedure-1-to-create-the-add-in-manifest-file-for-bing-search"></a>手順 1. Bing Search 用のアドインのマニフェスト ファイルを作成するには


- ローカル ディレクトリに XML ファイルを作成します。この XML ファイルには  **OfficeApp** 要素と子要素を記述します。詳細については「 [Office アドインの XML マニフェスト](../../docs/overview/add-in-manifests.md)」を参照してください。たとえば、以下の XML を記述したファイルを BingSearch.xml という名前で作成します。
    
```XML
   <?xml version="1.0" encoding="utf-8"?>
 <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" 
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xsi:type="TaskPaneApp">
   <Id>1234-5678</Id>
   <Version>15.0</Version>
   <ProviderName>Microsoft</ProviderName>
   <DefaultLocale>en-us</DefaultLocale>
   <DisplayName DefaultValue="Bing Search">
   </DisplayName>
   <Description DefaultValue="Search selected data on Bing">
   </Description>
   <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
   </IconUrl>
   <Capabilities>
     <Capability Name="Project"/>
   </Capabilities>
   <DefaultSettings>
     <SourceLocation DefaultValue="http://m.bing.com">
     </SourceLocation>
   </DefaultSettings>
   <Permissions>ReadWriteDocument</Permissions>
 </OfficeApp>
```

- アドインのマニフェストで必要な要素を次に示します。
  - **OfficeApp** 要素では、アドインの種類が作業ウィンドウであることを `xsi:type="TaskPaneApp"` 属性で指定します。
  - **Id** 要素は UUID で、一意である必要があります。
  - **Version** 要素は、アドインのバージョンです。 **ProviderName** 要素は、アドインを提供する企業または開発者の名前です。 **DefaultLocale** 要素は、マニフェストで指定する文字列の既定のロケールです。
  - **DisplayName** 要素は、Project 2013 のリボンの [ **ビュー**] タブで [ **作業ウィンドウ アドイン**] ドロップダウン リストに表示される名前です。値は最大 32 文字です。
  - **Description** 要素は、既定のロケールでのアドインの説明です。値は最大 2000 文字です。
  - **Capabilities** 要素は、1 つまたは複数の **Capability** 子要素を持ち、その中でホスト アプリケーションを指定します。
  - **DefaultSettings** 要素には、アドインが使用するファイル共有上の HTML ファイルのパスまたは Web ページの URL を指定する **SourceLocation** 要素が含まれています。作業ウィンドウ アドインでは、 **RequestedHeight** 要素と **RequestedWidth** 要素は無視されます。
  - **IconUrl** 要素は省略可能です。ファイル共有のアイコンまたは Web アプリケーションのアイコンの URL を指定できます。
    
- (省略可能) 他のロケール用の値を表す  **Override** 要素を追加します。たとえば、次のマニフェストでは、 **DisplayName**、 **Description**、 **IconUrl**、および  **SourceLocation** に対し、フランス語の値を表す **Override** 要素を指定しています。
    
```XML
   <?xml version="1.0" encoding="utf-8"?>
 <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" 
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xsi:type="TaskPaneApp">
   <Id>1234-5678</Id>
   <Version>15.0</Version>
   <ProviderName>Microsoft</ProviderName>
   <DefaultLocale>en-us</DefaultLocale>
   <DisplayName DefaultValue="Bing Search">
     <Override Locale="fr-fr" Value="Bing Search"/>
   </DisplayName>
   <Description DefaultValue="Search selected data on Bing">
     <Override Locale="fr-fr" Value="Search selected data on Bing"></Override>
   </Description>
   <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
     <Override Locale="fr-fr" Value="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>
   </IconUrl>
   <Capabilities>
     <Capability Name="Project"/>
   </Capabilities>
   <DefaultSettings>
     <SourceLocation DefaultValue="http://m.bing.com">
       <Override Locale="fr-fr" Value="http://m.bing.com"/>
     </SourceLocation>
   </DefaultSettings>
   <Permissions>ReadWriteDocument</Permissions>
 </OfficeApp>
```


## <a name="installing-project-add-ins"></a>Project アドインのインストール


Project 2013 では、ファイル共有上のスタンドアロン ソリューションとして、またはプライベート アドイン カタログに、アドインをインストールできます。パブリック Office ストアでアドインをレビューおよび購入することもできます。

ファイル共有の中には、アドイン マニフェストの XML ファイルとサブディレクトリを複数配置することもできます。マニフェストのディレクトリの場所とカタログを追加または削除するには、Project 2013 の [ **セキュリティ センター**] ダイアログ ボックスの [ **信頼されているアドイン カタログ**] タブを使用します。Project にアドインを表示するには、マニフェスト内の  **SourceLocation** 要素で既存の Web サイトまたは HTML ソース ファイルを指定する必要があります。


 >**メモ**  Internet Explorer 9 以降がインストールされている必要がありますが、既定のブラウザーになっている必要はありません。Office アドインには Internet Explorer 9 のコンポーネントが必要です。既定のブラウザーとして使用できるのは、Internet Explorer 9 以降、Safari 5.0.6 以降、Firefox 5 以降、または Chrome 13 以降です。

手順 2. では、Project 2013 がインストールされているローカル コンピューター上に Bing Search アドインをインストールします。しかし、アドインのインフラストラクチャでは  `C:\Project\AppManifests` などのローカル ファイル パスを直接使用しないので、ローカル コンピューター上にネットワーク共有を作成できます。希望に応じて、リモート コンピューター上にファイル共有を作成することもできます。


### <a name="procedure-2-to-install-the-bing-search-add-in"></a>手順 2. Bing Search アドインをインストールするには


1. アドイン マニフェスト用のローカル ディレクトリを作成します。たとえば、 `C:\Project\AppManifests` ディレクトリを作成します。
    
2. `C:\Project\AppManifests` ディレクトリをAppManifests として共有し、ファイル共有へのネットワーク パスが `\\ServerName\AppManifests` になるようにします。
    
3. BingSearch.xml マニフェスト ファイルを  `C:\Project\AppManifests` ディレクトリにコピーします。
    
4. Project 2013で、[ **Project のオプション**] ダイアログ ボックスを開き、[ **セキュリティ センター**]、[ **セキュリティ センターの設定**] の順に選択します。
    
5. [ **セキュリティ センター**] ダイアログ ボックスの左側のウィンドウで、[ **信頼されているアドイン カタログ**] を選択します。
    
6. [ **信頼されているアドイン カタログ**] ウィンドウ (図 1 を参照) で、 [ **カタログの URL**] テキスト ボックスにパス「 `\\ServerName\AppManifests`」を追加し、[ **カタログの追加**]、[ **OK**] の順に選択します。
    
     >**注** 図 1 は [**信頼できるカタログのアドレス**] リストの非公開カタログの、2 つのファイル共有と 1 つの架空の URL を示しています。既定のファイル共有に設定できるのは 1 つのファイル共有のみで、既定のカタログに設定できるのは 1 つのカタログ URL のみです。たとえば、`\\Server2\AppManifests` を既定に設定した場合、Project は `\\ServerName\AppManifests` の **[既定]** チェック ボックスをオフにします。既定の選択を変更した場合、**[クリア]** を選択すると、インストールしたアドインを削除して、Project を再起動することができます。Project が開いているときに既定のファイル共有または SharePoint カタログにアドインを追加した場合、Project を再起動する必要があります。

    **図 1.セキュリティ センターを使用したアドインのマニフェストのカタログの追加**

    ![セキュリティ センターを使用してアプリ マニフェストを追加](../../images/pj15_AgaveOverview_TrustCenter.PNG)

7. [ **プロジェクト**] リボンで、[ **Office アドイン**] ドロップダウン メニューの [ **すべて表示**] を選択します。[ **アドインの挿入**] ダイアログ ボックスで、[ **共有フォルダー**] を選択します (図 2 を参照)。
    
    **図 2.ファイル共有にあるアドインの起動**

    ![ファイル共有にある Office アプリの起動](../../images/pj15_AgaveOverview_StartAgaveApp.PNG)

8. Bing Search アドインを選択し、 [ **挿入**] を選択します。
    
図 3 のように、作業ウィンドウに Bing Search アドインが表示されます。作業ウィンドウのサイズを手動で変更して、Bing Search アドインを使用できます。

**図 3.Bing Search アドインの使用**

![Bing Search アプリの使用](../../images/pj15_AgaveOverview_BingSearch.gif)


## <a name="distributing-project-add-ins"></a>Project アドインの配布


アドインの配布は、ファイル共有、SharePoint ライブラリのアドイン カタログ、または Office ストア の プロジェクトのアドイン で行えます。詳細については、「 [Office アドインを発行する](../publish/publish.md)」を参照してください。


## <a name="additional-resources"></a>その他のリソース



- [Office アドイン プラットフォームの概要](../../docs/overview/office-add-ins.md)
    
- [Office アドインの XML マニフェスト](../../docs/overview/add-in-manifests.md)
    
- [JavaScript API for Office](../../reference/javascript-api-for-office.md)
    
- [テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
    
- [社内の Project Server OData サービスで REST を使用する Project アドインを作成する](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)
    
- [Project 用の作業ウィンドウ アドインを PWA に接続する](http://blogs.msdn.com/b/project_programmability/archive/2012/11/02/connecting-a-project-task-pane-app-to-pwa.aspx)
    
- [Project 2013 SDK のダウンロード](https://www.microsoft.com/en-us/download/details.aspx?id=30435%20)
    
