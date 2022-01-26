ローカル Web サーバーが既に実行されていて、アドインが Excel に既に読み込まれている場合は、手順 2 に進みます。 そうでない場合は、ローカル Web サーバーを起動し、アドインのサイドロードを行います。 

- Excel でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインが読み込まれた Excel が開きます。

    ```command&nbsp;line
    npm start
    ```

- Excel on the web でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します。 "{url}" を、アクセス許可を持っている OneDrive または SharePoint ライブラリ上の Excel ドキュメントの URL に置き換えます。

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

