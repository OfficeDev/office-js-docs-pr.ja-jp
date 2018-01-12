# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>発行のための準備として Visual Studio を使用してアドインをパッケージ化する

Office アドイン パッケージには、アドインの発行に使用する XML [マニフェスト ファイル](../overview/add-in-manifests.md)が含まれています。 プロジェクトの Web アプリケーション ファイルは個別に発行する必要があります。 この記事では、Visual Studio 2015 を使用して、Web プロジェクトを展開し、アドインをパッケージ化する方法について説明します。

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a>Visual Studio 2015 を使用して Web プロジェクトを展開するには

次に示す、Visual Studio 2015 を使用して Web プロジェクトを展開する手順を実行します。

1. **[ソリューション エクスプローラー]** で、アドイン プロジェクトのショートカット メニューを開き、**[発行]** を選択します。
    
    [**アドインの発行**] ページが表示されます。
    
2. **[現在のプロファイル]** ドロップダウン リストで、プロファイルを選択するか、**[新規…]** を選択して新しいプロファイルを作成します。
    
     >**注:** 発行プロファイルでは、展開先のサーバー、サーバーへのログオンに必要な資格情報、展開するデータベース、およびその他の展開オプションを指定します。

    **[新規...]** を選択すると、**[発行プロファイルの作成]** ウィザードが表示されます。 このウィザードを使用して、Microsoft Azure などの Web サイトをホストするプロバイダーから発行プロファイルをインポートするか、新しいプロファイルを作成するかして、次の手順でサーバー、資格情報、その他の設定を追加することができます。
    
    発行プロファイルのインポートまたは新しい発行プロファイルの作成の詳細については、「[発行プロファイルの作成](http://msdn.microsoft.com/ja-JP/library/dd465337.aspx#creating_a_profile)」を参照してください。
    
3. 「**アドインを発行する**」ページで、 [**Web プロジェクトの配置**] リンクを選択します。
    
    [**Web の発行**] ダイアログ ボックスが表示されます。 このウィザードの使用方法の詳細については、「[方法: Visual Studio でオンクリック発行を使用して Web プロジェクトを配置する](http://msdn.microsoft.com/ja-JP/library/dd465337.aspx)」を参照してください。
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a>Visual Studio 2015 を使用してアドインをパッケージ化するには

次に示す、Visual Studio 2015 を使用してアドインをパッケージ化する手順を実行します。

1. **[アドインを発行する]** ページで、**[アドインのパッケージ化]** リンクをクリックします。
    
    **[Office/SharePoint アドインの発行]** ウィザードが表示されます。
    
2. **[Web サイトがホストされている場所]** ドロップダウン リストで、アドインのコンテンツ ファイルをホストする Web サイトの URL を選択するか入力して、**[完了]** を選択します。
    
    このウィザードを完了するには、HTTPS プレフィックスで始まるアドレスを指定する必要があります。 一般に、Web サイトの HTTPS を使用することが推奨されていますが、アドインを Office ストアに発行する予定がない場合、その必要はありません。 Web サイトの HTTP エンドポイントを使用する場合は、パッケージの作成の完了後に、テキスト エディターで XML マニフェスト ファイルを開いて、Web サイトの HTTPS プレフィックスを HTTP プレフィックスに置換します。 詳細については、「[アプリおよびアドインを SSL でセキュリティ保護する理由](http://msdn.microsoft.com/ja-JP/library/jj591603#bk_q7)」を参照してください。
    
     >**注:** Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。

    Visual Studio は、アドインの発行に必要なファイルを生成して、発行の出力フォルダーを開きます。 
    
Office ストアにアドインを提出する予定がある場合は、**[検証チェックを実行する**] リンクをクリックして、アドインが受け入れられなくなる問題点を特定します。 アドインをストアに提出する前に、すべての問題に対処してください。

XML マニフェストを適切な場所にアップロードして、[アドインを発行](../publish/publish.md)できるようになりました。 XML マニフェストは、`app.publish` フォルダーの `OfficeAppManifests` にあります。 次に例を示します。

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="additional-resources"></a>その他の技術資料



- [Office アドインを発行する](../publish/publish.md)
    
- [Office ストアに Office アドインと SharePoint アドインおよび Office 365 Web アプリを提出する](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
