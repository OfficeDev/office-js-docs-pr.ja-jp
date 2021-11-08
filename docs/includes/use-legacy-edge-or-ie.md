プロジェクトが node.js ベース (つまり、Visual Studio およびインターネット インフォメーション サーバー (IIS) で開発されていない) 場合は、Windows の Office を強制的に使用して、エッジ レガシまたは Internet Explorer を使用してアドインを実行できます。通常は新しいブラウザーを使用する Windows バージョンと Office バージョンを組み合わせて使用している場合でも、アドインを実行できます。 Windows バージョンと Office バージョンのさまざまな組み合わせで使用されるブラウザーの詳細については、「Office アドインで使用されるブラウザー」を[参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

1. プロジェクトが *Yo* Officeツールで作成されていない場合は、office-addin-dev-settings ツールをインストールする必要があります。 コマンド プロンプトで次のコマンドを実行します。

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. プロジェクトのルートにあるコマンド プロンプトOfficeコマンドで使用するブラウザーを指定します。 プロジェクト `<path-to-manifest>` のルートにある場合は、マニフェストファイル名の相対パスに置き換える。 どちらか `<webview>` またはで置き `ie` 換える `edge-legacy` 。

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    次に例を示します。

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    Webview の種類が IE (またはエッジ レガシ) に設定されているというメッセージがコマンド ラインに表示されます。

1. 完了したら、Office を設定して、Windows バージョンと Office バージョンを次のコマンドと組み合わせて使用して、既定のブラウザーの使用を再開します。

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
