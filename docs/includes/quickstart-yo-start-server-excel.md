
以下の手順を実行し、ローカル Web サーバーを起動してアドインのサイドロードを行います。

> [!NOTE]
> 開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。 次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。

> [!TIP]
> Mac でアドインをテストしている場合は、先に進む前に次のコマンドを実行してください。 このコマンドを実行すると、ローカル Web サーバーが起動します。
>
> ```command&nbsp;line
> npm run dev-server
> ```

- Excel でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動し (既に実行されていない場合)、アドインが読み込まれたときに Excel が開きます。

    ```command&nbsp;line
    npm start
    ```

- Excel Online でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。

    ```command&nbsp;line
    npm run start:web
    ```

    アドインを使用するには、Excel Online で新しいブックを開き、「[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)」の手順に従ってアドインをサイドロードします。

