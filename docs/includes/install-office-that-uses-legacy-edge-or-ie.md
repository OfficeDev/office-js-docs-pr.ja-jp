Microsoft Edge 従来版 Webview (EdgeHTML) を使用してアドインを実行するバージョンの Office (Microsoft 365 サブスクリプションからダウンロード) または Internet Explorer (Trident) を使用するバージョンをインストールするには、次の手順に従います。

1. 任意の Office アプリケーションで、リボンの [ **ファイル** ] タブを開き、[ **Office アカウント** ] または [アカウント] を選択 **します**。 [**_About host-name]\(ホスト名について_****\) ボタン** (Word についてなど) を選択します。
1. 開いたダイアログで、xx.x.xxxxx.xxxxx の完全なビルド番号を見つけて、そのコピーをどこかに作成します。
1. [Office 展開ツール](https://www.microsoft.com/download/details.aspx?id=49117)をダウンロードします。
1. ダウンロードしたファイルを実行して、ツールを抽出します。 ツールをインストールする場所を選択するように求められます。
1. ツールをインストールしたフォルダー (ファイルがある場所 `setup.exe` ) で、名前 `config.xml` のテキスト ファイルを作成し、次の内容を追加します。

    ```xml
    <Configuration>
      <Add OfficeClientEdition="64" Channel="SemiAnnual" Version="16.0.xxxxx.xxxxx">
        <Product ID="O365ProPlusRetail">
          <Language ID="en-us" />
        </Product>
      </Add>
    </Configuration>
    ```

1. 値を変更します `Version` 。

    - Edge レガシを使用するバージョンをインストールするには、 に変更します `16.0.11929.20946`。
    - Internet Explorer を使用するバージョンをインストールするには、 に変更します `16.0.10730.20348`。

1. 必要に応じて、 の `OfficeClientEdition` 値を に `"32"` 変更して 32 ビット Office をインストールし、Office を `Language ID` 別の言語でインストールするために必要に応じて値を変更します。
1. *管理者として* コマンド プロンプトを開きます。
1. ファイルと `config.xml` ファイルを含むフォルダーに`setup.exe`移動します。
1. 次のコマンドを実行します。

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    このコマンドは Office をインストールします。 このプロセスには数分かかる場合があります。

1. [Office キャッシュをクリアします](../testing/clear-cache.md)。

> [!IMPORTANT]
> インストール後は、Office の自動更新をオフにして、使用を完了する前に使用する Webview を使用しないバージョンに Office が更新されないようにしてください。 **これは、インストールから数分以内に発生する可能性があります。** 次の手順に従ってください。
>
> 1. Office アプリケーションを起動し、新しいドキュメントを開きます。
> 1. リボンの [ **ファイル** ] タブを開き、[ **Office アカウント** ] または [アカウント] を選択 **します**。
> 1. [**製品情報**] 列で、[**更新オプション]** を選択し、[**更新を無効にする] を選択します**。 このオプションを使用できない場合、Office は自動的に更新されないように既に構成されています。

古いバージョンの Office の使用が完了したら、ファイルを編集 `config.xml` し、 を以前にコピーしたビルド番号に変更 `Version` して、新しいバージョンを再インストールします。 次に、管理者コマンド `setup.exe /configure config.xml` プロンプトでコマンドを繰り返します。 必要に応じて、自動更新を再度有効にします。
