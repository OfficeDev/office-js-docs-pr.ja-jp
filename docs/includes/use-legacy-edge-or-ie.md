プロジェクトが node.js ベース (つまり、Visual Studio およびインターネット インフォメーション サーバー (IIS Office) で開発されていない) 場合は、Windows バージョンと Office バージョンを組み合わせて使用している場合でも、Windows でエッジ レガシまたは Internet Explorer を使用してアドインを実行できます。 Windows バージョンと Office バージョンのさまざまな組み合わせで使用されるブラウザーの詳細については、「Office アドインで使用されるブラウザー」を[参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

1. プロジェクトが [Yeoman](../develop/yeoman-generator-overview.md) ジェネレーターを使用して Office アドイン ツールで作成されていない場合は、office-addin-dev-settings ツールをインストールする必要があります。 コマンド プロンプトで次のコマンドを実行します。

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. プロジェクトのルートにあるコマンド プロンプトOfficeコマンドで使用するブラウザーを指定します。 プロジェクト `<path-to-manifest>` のルートにある場合は、マニフェストファイル名の相対パスに置き換える。 どちらかまたは `<webview>` で置き換 `ie` える `edge-legacy`。

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    次に例を示します。

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    Webview の種類が IE (またはエッジ レガシ) に設定されているというメッセージがコマンド ラインに表示されます。

1. 完了したら、Office Windows と Office バージョンの組み合わせに既定のブラウザーを使用して次のコマンドを使用して再開Officeを設定します。

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
