
以下の手順を実行し、ローカル Web サーバーを起動してアドインのサイドロードを行います。

[!INCLUDE [alert use https](alert-use-https.md)]

> [!TIP]
> Mac でアドインをテストしている場合は、先に進む前に次のコマンドを実行してください。 このコマンドを実行すると、ローカル Web サーバーが起動します。
>
> ```command&nbsp;line
> npm run dev-server
> ```

- Excel でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 これにより、ローカル Web サーバーが起動し、アドインが読み込まれた状態で Excel が開きます。

    ```command&nbsp;line
    npm start
    ```

- ブラウザー上の Excel でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します。 "{url}" を、アクセス許可を持っている OneDrive または SharePoint ライブラリ上の Excel ドキュメントの URL に置き換えます。

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]
