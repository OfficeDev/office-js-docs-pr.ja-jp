# <a name="debug-office-add-ins-on-ipad-and-mac"></a>iPad と Mac で Office アドインをデバッグする

Windows でのアドインの開発とデバッグには Visual Studio を使用できますが、iPad と Mac で使用して アドインをデバッグすることはできません。アドインは HTML と Javascript を使用して開発されているため、さまざまなプラットフォームで機能するように設計されていますが、さまざまなブラウザーで HTML の表示方法に微妙な違いがあります。この記事では、iPad または Mac で動作するアドインをデバッグする方法を説明します。 

## <a name="debugging-with-vorlonjs"></a>Vorlon.JS を使用したデバッグ 

Vorlon.JS は、F12 ツールに似た、Web ページのデバッガーです。 リモートで動作するように設計されているため、異なるデバイス間で Web ページをデバッグすることができます。 詳細については、[Vorlon の Web サイト](http://www.vorlonjs.com)を参照してください。  

### <a name="install-and-set-up-up-vorlonjs-on-a-mac-or-ipad"></a>Mac または iPad に Vorlon.JS をインストールしてセットアップする 

1.  管理者としてデバイスにログオンします。

2.  まだ [Node.js](https://nodejs.org) をインストールしていない場合は、インストールします。 

2.  **[ターミナル]** ウィンドウを開き、コマンド `npm i -g vorlon` を入力します。 ツールが `/usr/local/lib/node_modules/vorlon` にインストールされます。

### <a name="configure-vorlonjs-to-use-https"></a>HTTPS を使用するように Vorlon.JS を構成する

Vorlon.JS を使用してアプリケーションをデバッグするには、既知の場所から Vorlon.JS スクリプトを読み込むアプリケーションの開始ページに `<script>` タグを追加します (詳細については、次の手順を参照してください)。 アドインには、HTTPS プロトコル、つまり SSL が必要です。 アドインで使用するすべてのスクリプトは HTTPS サーバーからホストされるように拡張する必要があります。これには、Vorlon.JS スクリプトも含まれます。 そのため、アドインで Vorlon.JS を使用するには、SSL を使用するように Vorlon.JS を構成することが必要になります。 

4.  **[Finder]** で、`/usr/local/lib/node_modules/vorlon` に移動し、`/Server` フォルダーのコンテキスト メニュー (右クリック) を開き、**[情報を見る]** を選択します。

5.  **[サーバー情報]** ウィンドウの右下隅にある南京錠アイコンを選択して、フォルダーのロックを解除します。

6. ウィンドウの **[共有とアクセス権]** セクションで、**スタッフ** グループの **[特権]** を **[読み取り/書き込み]** に設定します。

7. 南京錠アイコンをもう一度選択して、フォルダーを***再度ロック***します。

8. **[Finder]** に戻り、`/Server` サブフォルダーを展開し、ファイル `config.json` を右クリックして、**[情報を見る]** を選択します。

9. **[config.json 情報]** ウィンドウで、親 `/Server` フォルダーに対して行ったものと同じ方法でファイルの特権を変更します。 必ず再度ロックしてからウィンドウを閉じてください。

10. **[Finder]** に戻り、ファイル `config.json` を右クリックして、**[このアプリケーションで開く]**、**[テキストエディット]** の順に選択します。 ファイルがテキスト エディターで開きます。

11. **useSSL** プロパティの値を `true` に変更します。

12. **[プラグイン]** セクションで、**ID** が `OFFICE` で**名前**が `Office Addin` のプライグインを検索します。 プラグインの **enabled** プロパティがまだ `true` になっていない場合は、`true` に設定します。

13. ファイルを保存し、エディターを閉じます。

5.  **[検索]** で `/usr/local/lib/node_modules/vorlon` に移動して、`Server` サブフォルダーを右クリックし、**[フォルダーの新しいターミナル]** を選択します。 
    
7.  **[ターミナル]** ウィンドウで、`sudo vorlon` と入力します。 管理者パスワードの入力を求めるダイアログ ボックスが表示されます。 Vorlon サーバーが起動します。 **[ターミナル]** ウィンドウを開いたままにしておきます。

6.  ブラウザー ウィンドウを開き、Vorlon.JS インターフェイスの `https://localhost:1337` に進みます。 ダイアログ ボックスが表示されたら、**[常に]** を選択して、セキュリティ証明書を信頼します。 

    >**注:**ダイアログ ボックスが表示されない場合は、手動で証明書を信頼する必要があります。 証明書ファイルは `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt` です。 次の手順を実行します。 問題が発生した場合は、Macintosh または iPad のヘルプを参照してください。 
    >
    >1. ブラウザー ウィンドウを閉じ、Vorlon サーバーを実行している **[ターミナル]** ウィンドウで、Control-C を使用してサーバーを停止します。
    >2. **[Finder]** で、`server.crt` ファイルを右クリックして、**[キーチェーン アクセス]** を選択します。 **[キーチェーン アクセス]** ウィンドウが開きます。
    >2. 左側の **[キーチェーン]** リストで、**[ログイン]** がまだ選択されていない場合は選択し、**[カテゴリ]** セクションで **[証明書]** を選択します。 証明書 **localhost** が一覧表示されます。
    >3. 証明書 **localhost** を右クリックし、**[情報を見る]** を選択します。 **[localhost]** ウィンドウが開きます。
    >4. **[信頼]** セクションで、**[この証明書を使用する場合]** というラベルの付いたセレクターを開いて、**[常に信頼する]** を選択します。 
    >5. **[localhost]** ウィンドウを閉じます。 アクションが成功すると、**[キーチェーン アクセス]** ウィンドウの **localhost** 証明書のアイコンに青い円で囲まれた白い十字が表示されます。

### <a name="configure-the-add-in-for-vorlonjs-debugging"></a>Vorlon.JS デバッグ用のアドインを構成する

1. 次のスクリプト タグを、アドインの home.html ファイル (またはメイン HTML ファイル) の `<head>` セクションに追加します。

    ```    
    <script src="https://localhost:1337/vorlon.js"></script>    
    ```  

2. Azure Web サイトなど、Mac または iPad からアクセス可能な Web サーバーにアドイン Web アプリケーションを展開します。 

3. アドイン マニフェストに URL が表示されるすべての場所で、アドインの URL を更新します。

4. アドイン マニフェストを Mac または iPad 上の次のフォルダーにコピーします: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`。ここで、*{host_name}* は、Word、Excel、PowerPoint、または Outlook です。

### <a name="inspect-an-add-in-in-vorlonjs"></a>Vorlon.JS でアドインを検査する

1. Vorlon サーバーが実行されていない場合、**[Finder]** で `/usr/local/lib/node_modules/vorlon` に移動して、`Server` サブフォルダーを右クリックし、**[フォルダーの新しいターミナル]** を選択します。 
    
7.  **[ターミナル]** ウィンドウで、`sudo vorlon` と入力します。 管理者パスワードの入力を求めるダイアログ ボックスが表示されます。 Vorlon サーバーが起動します。 **[ターミナル]** ウィンドウを開いたままにしておきます。

6.  ブラウザー ウィンドウを開き、Vorlon.JS インターフェイスの `https://localhost:1337` に進みます。

7. アドインをサイドロードします。 アドインが Excel、PowerPoint、Word 用の場合は、「[iPad または Mac で Office アドインをサイドロードする](https://dev.office.com/docs/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)」の説明に従ってサイドロードします。 アドインが Outlook アドインである場合は、「[テストのために Outlook アドインをサイドロードする](https://dev.office.com/docs/add-ins/testing/sideload-outlook-add-ins-for-testing)」の説明に従ってサイドロードします。 アドインでアドイン コマンドを使用しない場合は、アドインが直ちに開きます。 それ以外の場合は、ボタンを選択してアドインを開きます。 Office ホスト アプリケーションのビルドに応じて、ボタンは **[ホーム]** タブまたは **[アドイン]** タブのいずれかに表示されます。

アドインは、Vorlon.JS のクライアントのリスト (Vorlon.JS インターフェイスの左側) に **{OS} - n** として表示されます。*n* は数値、*{OS}* は "Macintosh" などのデバイスの種類です。 

![Vorlon.js インターフェイスを示すスクリーンショット](../../images/vorlon_interface.png)

Vorlon ツールには、さまざまなプラグインがあります。現在有効になっているプラグインはツールの上部にタブとして表示されます。 (左側にある歯車アイコンを選択すると、さらに別のプラグインを有効にすることができます。)これらのプラグインは、F12 ツールの機能に似ています。 たとえば、DOM 要素の強調表示、コマンドの実行などを行えます。 詳細については、[Vorlon ドキュメントの「コア プラグイン」](http://vorlonjs.com/documentation/#console)を参照してください。 

**Office アドイン** プラグインにより Office.js に特別な機能 (オブジェクト モデルを調査する機能、Office.js の呼び出しを実行する機能、オブジェクト プロパティの値を読み取る機能など) が追加されます。 手順については、「[Office アドインをデバッグするための VorlonJS プラグイン](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/)」を参照してください。

>**注:**Vorlon.JS にブレーク ポイントを設定する方法はありません。

## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a>Mac または iPad 上の Office アプリケーションのキャッシュのクリア

アドインはパフォーマンス上の理由から、Office for Mac でキャッシュされることが多いです。通常、キャッシュはアドインを再読み込みすることでクリアされます。同じドキュメント内に複数のアドインが存在する場合、再読み込み時にキャッシュを自動的にクリアするプロセスは信頼できない場合があります。 

Mac では、`/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` フォルダー内にあるすべてを削除することによってキャッシュを手動でクリアできます。 

iPad では、アドインの JavaScript から `window.location.reload(true)` を呼び出して、強制的に再読み込みすることができます。 または、Office を再インストールすることができます。
