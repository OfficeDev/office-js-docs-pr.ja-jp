1. プロジェクトのルート (**[...]/My Office Add-in**) で bash ターミナルを開いて、次に示すコマンドを実行して開発用サーバーを起動します。

    ```command&nbsp;line
    npm start
    ```

2. Internet Explorer または Microsoft Edge のいずれかを開き、`https://localhost:3000` にアクセスします。 証明書エラーなしでページが読み込まれた場合は、この記事の次のセクションに進みます (**お試しください**)。 サイトの証明書が信頼できないとブラウザーに表示された場合は、次の手順に進みます。

3. Office Web アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。 ブラウザーにサイトの証明書が信頼されていないことが示された場合は、その証明書を信頼された証明書として追加する必要があります。 詳細については、「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」を参照してください。

    > [!NOTE]
    > 「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」で説明されているプロセスを完了した後でも、Chrome (Web ブラウザー) は、サイトの証明書が信頼されていないことを引き続き示すことがあります。 そのため、Internet Explorer または Microsoft Edge のいずれかを使用して、証明書が信頼できることを確認する必要があります。 

4. ブラウザーに証明書エラーなしでアドイン ページが読み込まれたら、アドインをテストする準備ができています。
