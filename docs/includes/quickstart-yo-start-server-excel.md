
以下の手順を実行し、ローカル Web サーバーを起動してアドインのサイドロードを行います。

> [!NOTE]
> 開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。 次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。

> [!TIP]
> Mac でアドインをテストしている場合は、先に進む前に次のコマンドを実行してください。 このコマンドを実行すると、ローカル Web サーバーが起動します。
>
> ```command&nbsp;line
> npm run dev-server
> ```

- Excel でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインが読み込まれた Excel が開きます。

    ```command&nbsp;line
    npm start
    ```

- ブラウザー上の Excel でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。

    ```command&nbsp;line
    npm run start:web
    ```

    アドインを使用するには、Excel on the web で新しいブックを開き、「[Office on the web で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)」の手順に従ってアドインをサイドロードします。

