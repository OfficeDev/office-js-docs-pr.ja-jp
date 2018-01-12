
# <a name="create-an-office-add-in-using-any-editor"></a>任意のエディターを使用して Office アドインを作成する

Office アドインの作成に Yeoman ジェネレーターを使用することができます。Yeoman ジェネレーターは、プロジェクトのスキャフォールディングとビルドの管理を提供します。`manifest.xml` ファイルで、アドインが格納されている場所と表示方法を Office アプリケーションに指示します。Office アプリケーションは Office 内でホスティングを行います。

 >**メモ:**この手順では、Mac のターミナルを使いますが、その他のシェル環境を使うこともできます。 


## <a name="prerequisites-for-the-yeoman-generator"></a>Yeoman ジェネレーターの前提条件

Yeoman Office ジェネレーターをインストールするには、事前に [git](https://git-scm.com/downloads) と node.js をコンピューターにインストールしておく必要があります。Mac の場合、[Node Version Manager](https://github.com/creationix/nvm) を使用して、正しいアクセス許可で node.js をインストールすることをお勧めします。Windows の場合、node.js を [nodejs.org](https://nodejs.org/en/) からインストールできます。

>**メモ:**Windows の場合、git をインストールするときは既定値を使用します。ただし、次の場合は例外です。

>- Windows コマンド プロンプトから Git を使用する
>- Windows の既定のコンソール ウィンドウを使用する

node.js をインストールした後、ターミナルを開き、ジェネレーターをグローバルにインストールします。

```
npm install -g yo generator-office
```


## <a name="create-the-default-files-for-your-add-in"></a>アドインの既定のファイルを作成する

プロジェクトをスキャホールディングするディレクトリで Yeoman ジェネレーターを実行します。Office アドインを開発する前に、プロジェクト用のフォルダーを作成します。

ターミナルで、プロジェクトを作成する親フォルダーに移動します。その後、次のコマンドを使用して、_myHelloWorldaddin_ という名前の新しいフォルダーを作成し、現在のディレクトリをそのフォルダーに切り替えます。




```
mkdir myHelloWorldaddin
cd myHelloWorldaddin
```

Yeoman ジェネレーターを使用して、任意のアドインを作成します。この記事の手順では、単純な作業ウィンドウ アドインを作成します。ジェネレーターを実行するには、次のコマンドを入力します。




```
yo office
```

**アドイン用の Yeoman ジェネレーターの入力**

ジェネレーターは、確認を求める以下のメッセージを表示します。 


- 新しいサブフォルダ― -- 「_N_」と入力します
- アドイン名 -- 「_myHelloWorldaddin_」と入力します 
- サポートされている Office アプリケーション - 任意のアプリケーションを選択できます
- 新しいアドインの作成 -- 「_はい、新しいアドインを作成します_」と入力します
- [TypeScript](https://www.typescriptlang.org/) の追加 -- 「_N_」と入力します
- フレームワークの選択 -- _Jquery_ を使用します

>**メモ:**Office UI Fabric React を使用する Office アドインを作成するには、次を入力します:
>- [TypeScript](https://www.typescriptlang.org/) の追加 -- _Y_ を使用します
>- フレームワークの選択 -- _React_ を使用します

![プロジェクトの入力を要求する Yeoman ジェネレーターの Git](../../images/gettingstarted-fast.gif)

これは、アドインの構造および基本的なファイルを作成します。


## <a name="hosting-your-office-add-in"></a>Office アドインをホストする

Office アドインは、開発中の場合でも HTTPS 経由でホストする必要があります。Yo Office は bsconfig.json を作成します。これは Browsersync を使用して、複数のデバイスでファイルの変更を同期することによりアドインの微調整とテストをより迅速に行えるようにします。 

コンソールで次のコマンドを入力することにより、https://localhost:3000 でローカルの HTTPS サイトを起動します。


```
npm start
```

Browsersync は HTTPS サーバーを起動し、プロジェクトの index.html ファイルを起動します。"この Web サイトのセキュリティ証明書には問題があります。" というエラー メッセージが表示されます。


![エラーを回避して既定の index.html ファイルを表示するためのプロセスを示す gif](../../images/ssl-chrome-bypass.gif)

このエラーは、開発環境が信頼する必要がある自己署名 SSL 証明書が Browsersync に含まれているために発生します。このエラーの解決方法の詳細については、[自己署名証明書の追加](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)に関する記事をご覧ください。

## <a name="sideload-the-add-in-into-office"></a>アドインを Office にサイドロードする

サイドロードを使用して、Office クライアント内でのテストのためにアドインをインストールできます。

- [テスト用に Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [テスト用に iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)   
- [テスト用に Outlook アドインをサイドロードする](../outlook/testing-and-tips.md)

## <a name="develop-your-office-add-in"></a>Office アドインを開発する

任意のテキスト エディターを使用して、カスタム Office アドインのファイルを開発できます。

> **重要:**manifest-myHelloWorldaddin.xml ファイルは、アドインを操作する方法を Office クライアント アプリケーションに指示します。`<id>` タグの値は、Yo Office がプロジェクトを生成するときに作成する GUID です。アドインの GUID は変更しないでください。ホストが Azure の場合、`SourceLocation` の値は _https:// [Web アプリの名前].azurewebsites.net/[アドインのパス]_ のような URL です。この例のように、自己ホスト型のオプションを使用する場合は、_https://localhost:3000/[アドインのパス]_ になります。


## <a name="debug-your-office-add-in"></a>Office アドインをデバッグする


アドインをデバッグするには、いくつかの方法があります。

- 作業ウィンドウからデバッガーをアタッチする (Office 2016 for Windows)。
- ブラウザーの開発者ツールを使用する。
- Windows 10 で F12 開発者ツールを使用する。

### <a name="attach-debugger-from-the-task-pane"></a>作業ウィンドウからデバッガーをアタッチする

Office 2016 for Windows のビルド 77xx.xxxx 以降では、作業ウィンドウからデバッガーをアタッチすることができます。デバッガーのアタッチ機能によって、デバッガーが適切な Internet Explorer プロセスに直接アタッチされます。デバッガーは、Yeoman Generator、Visual Studio Code、node.js、Angular、その他のツールのどれを使用しているかに関係なくアタッチすることができます。 

詳細については、「[作業ウィンドウからデバッガーをアタッチする](../testing/attach-debugger-from-task-pane.md)」を参照してください。


### <a name="browser-developer-tools"></a>ブラウザーの開発者ツール 

Office Web クライアントを使用して、ブラウザーの開発者ツールを開き、ほかのクライアント側 JavaScript アプリケーションをデバッグした方法で、アドインをデバッグします。 

### <a name="f12-developer-tools-on-windows-10"></a>Windows 10 の F12 開発者ツール

Windows 10 で Office のデスクトップ クライアントを使用している場合、[Windows 10 で F12 開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)ことが可能です。
    
## <a name="next-steps"></a>次の手順

- [Office アドインを展開し、発行する](../publish/publish.md)
    
