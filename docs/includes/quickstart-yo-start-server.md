1. プロジェクトのルートで bash ターミナルを開いて、次に示すコマンドを実行して開発用サーバーを起動します。

    ```bash
    npm start
    ```

    これで、`https://localhost:3000` で Web サーバーが起動し、そのアドレスで既定のブラウザーが開きます。

2. Office Web アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。 ブラウザーにサイトの証明書が信頼されていないことが示された場合は、その証明書を信頼された証明書として追加する必要があります。 詳細については、「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」を参照してください。

    > [!NOTE]
    > 「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」で説明されているプロセスを完了した後でも、Chrome (Web ブラウザー) は、サイトの証明書が信頼されていないことを引き続き示すことがあります。 Chrome でのこの警告は無視して構いません。また、Internet Explorer あるいは Microsoft Edge の `https://localhost:3000` に移動して、証明書が信頼されていることを確認することもできます。 

3. ブラウザーに証明書エラーなしでアドイン ページが読み込まれたら、アドインをテストする準備ができています。 
