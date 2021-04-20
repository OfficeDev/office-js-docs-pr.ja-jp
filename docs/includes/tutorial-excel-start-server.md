ローカル web サーバーが既に実行されていて、アドインが既に Excel に読み込まれている場合は、手順2に進みます。 それ以外の場合は、ローカル web サーバーを起動して、アドインをサイドロードします。 

- Excel でアドインをテストするには、プロジェクトのルートディレクトリで次のコマンドを実行します。 これにより、ローカル web サーバーが起動 (まだ実行されていない場合) し、アドインが読み込まれた状態で Excel が開きます。

    ```command&nbsp;line
    npm start
    ```

- Web 上の Excel でアドインをテストするには、プロジェクトのルートディレクトリで次のコマンドを実行します。 このコマンドを実行すると、ローカル web サーバーが起動します (まだ実行していない場合)。

    ```command&nbsp;line
    npm run start:web
    ```

    アドインを使用するには、web 上の Excel で新しいドキュメントを開き、「 [office のサイドロード Office アドイン](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)」の手順に従ってアドインをサイドロードします。
