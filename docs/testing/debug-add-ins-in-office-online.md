
# <a name="debug-add-ins-in-office-online"></a>Office Online でアドインをデバッグする


Windows、Office 2013、または Office 2016 デスクトップ クライアントを実行していないコンピューター (たとえば、Mac で開発を行っている場合) でアドインの作成とデバッグを行えます。この記事では、Office Online を使用してアドインのテストとデバッグを行う方法について説明します。 

開始するには、


- Office 365 の開発者アカウントをまだお持ちでない場合はこれを取得します。または SharePoint サイトにアクセスできるようにします。
    
     >**注** 無料の Office 365 開発者アカウントにサインアップするには、[Office 365 開発者プログラム](https://dev.office.com/devprogram)にご参加ください。
     
- Office 365 (SharePoint Online) 上でアドイン カタログをセットアップするアドイン カタログとは、Office アドイン用のドキュメント ライブラリをホストする SharePoint Online の 専用サイト コレクションです。独自の SharePoint サイトを所有している場合は、アドイン カタログのドキュメント ライブラリをセットアップすることができます。詳細については、「[作業ウィンドウ アドインとコンテンツ アドインを SharePoint のアドイン カタログに発行する](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)」をご覧ください。
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a>Excel Online または Word Online からアドインをデバッグする

Office Online を使用してアドインをデバッグするには、


1. SSL をサポートするサーバーにアドインを展開します。
    
     >**注:**[Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して、アドインを作成し、ホストすることをお勧めします。
     
2. [アドイン マニフェスト ファイル](../../docs/overview/add-in-manifests.md)で、相対 URI ではなく絶対 URI を含めるように **SourceLocation** 要素の値を更新します。たとえば次のようにします。
    
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. SharePoint のアドイン カタログにある Office アドイン ライブラリにマニフェストをアップロードします。
    
4. Office 365 のアプリ起動ツールから Excel Online または Word Online を起動し、新しいドキュメントを開きます。
    
5. [挿入] タブで、 **[個人用アドイン]** または **[Office アドイン]** をクリックし、アプリにアドインを挿入してテストします。
    
6. お気に入りのブラウザーのツール デバッガーを使用してアドインをデバッグします。
    
    以下は、デバッグ時に発生する可能性がある問題です。
    
  - 表示される JavaScript エラーのいくつかは Office Online に起因している可能性があります。
    
  - ブラウザーが、バイパスが必要になる、無効な証明書エラーを表示することがあります。
    
  - コードにブレークポイントを設定する場合、Office Online から、保存できないというエラーがスローされることがあります。
    

## <a name="additional-resources"></a>その他のリソース


- [Office アドイン開発のベスト プラクティス](../overview/add-in-development-best-practices.md)
    
- 
  [Office ストアに提出されたアプリとアドインの検証ポリシー (バージョン 1.9)](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)
    
- 
  [効果的な Office ストア アプリおよびアドインを作成する](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)
    
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)
    
