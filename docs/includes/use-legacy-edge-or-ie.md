プロジェクトがnode.js ベースの場合 (Visual Studio とインターネット インフォメーション サーバー (IIS) では開発されていません)、通常は最新のブラウザーを使用する Windows バージョンと Office バージョンの組み合わせがある場合でも、Office on Windows で Edge レガシまたは Internet Explorer を使用してアドインを実行するように強制できます。 Windows バージョンと Office バージョンのさまざまな組み合わせによって使用されるブラウザーの詳細については、「 [Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

> [!NOTE]
> ブラウザーの変更を強制するために使用されるツールは、Microsoft 365 のベータ 版サブスクリプション チャネルでのみサポートされています。 [Office Insider プログラム](https://insider.office.com/join/windows)に参加し、**ベータ チャネル** オプションを選択して Office Beta ビルドにアクセスします。 「 [Office について: 使用している Office のバージョン](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)」も参照してください。
>
> 厳密には、ベータ チャネルを `webview` 必要とするのは、このツールの切り替えです ( **手順 2** を参照)。 このツールには、この要件を持たない他のスイッチがあります。

1. [プロジェクトが Office アドイン用 Yeoman ジェネレーターツールで](../develop/yeoman-generator-overview.md)作成 *されていない* 場合は、office-addin-dev-settings ツールをインストールする必要があります。 コマンド プロンプトで次のコマンドを実行します。

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. プロジェクトのルートにあるコマンド プロンプトで、Office で次のコマンドで使用するブラウザーを指定します。 相対パスに置き換えます `<path-to-manifest>` 。これは、プロジェクトのルートにある場合はマニフェスト ファイル名に過ぎません。 いずれか`ie`に置き換えるか`<webview>`、または `edge-legacy`.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    次に例を示します。

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    コマンド ラインに、Webview の種類が IE (または Edge レガシ) に設定されているというメッセージが表示されます。

1. 完了したら、次のコマンドを使用して、Windows バージョンと Office バージョンの組み合わせに対して既定のブラウザーの使用を再開するように Office を設定します。

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
