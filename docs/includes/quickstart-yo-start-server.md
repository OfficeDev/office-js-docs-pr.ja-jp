1. プロジェクトのルートで bash ターミナルを開いて  (**[...]/My Office Add-in**)、次に示すコマンドを実行して開発用サーバーを起動します。

    ```bash
    npm start
    ```

2. Internet Explorer または Microsoft Edge のいずれかを開き `https://localhost:3000`に移動します。 証明書エラーを発生せずにページを読み込んだ場合は、この記事 (**試してみる**) の次のセクションに進んでください。 お使いのブラウザが、サイトの証明書が信頼されていないことを示している場合は、次の手順に進みます。

3. Office Web アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。 ブラウザにサイトの証明書が信頼されていないことが示された場合は、その証明書を信頼された証明書として追加する必要があります。 詳細については、「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」を参照してください。

    > [!NOTE]
    > 「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」で説明されているプロセスを完了した後でも、Chrome (Web ブラウザー) は、サイトの証明書が信頼されていないことを引き続き示すことがあります。 したがって、証明書が信頼できることを確認するには、Internet Explorer または Microsoft Edge のいずれかを使用する必要があります。 

4. ブラウザに証明書エラーなしでアドイン ページが読み込まれたら、アドインをテストする準備ができています。
