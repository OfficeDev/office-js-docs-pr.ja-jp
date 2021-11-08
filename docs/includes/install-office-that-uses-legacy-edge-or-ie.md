Microsoft Edge 従来版 webview (EdgeHTML) を使用してアドインを実行するサブスクリプション Office のバージョンまたは Internet Explorer (Trident) を使用するバージョンをインストールするには、次の手順を使用します。

1. 任意のOfficeアプリケーションで、リボンの [ファイル] タブを開き、[アカウント] または [アカウント] **Officeを** 選択 **します**。 [ホスト **名 _について] ボタン_** (Word についてなど) **を選択します**。
1. 開いたダイアログで、完全な xx.x.xxxxx.xxxxx ビルド番号を見つけて、そのコピーをどこかに作成します。
1. 展開ツールをダウンロード[Officeインストールします](https://www.microsoft.com/download/details.aspx?id=49117)。
1. ツールをインストールしたフォルダー (ファイルがある場所) で、名前を持つテキスト ファイルを作成し、 `setup.exe` `config.xml` 次の内容を追加します。

    ```xml
    <Configuration>
      <Add OfficeClientEdition="64" Channel="SemiAnnual" Version="16.0.xxxxx.xxxxx">
        <Product ID="O365ProPlusRetail">
          <Language ID="en-us" />
        </Product>
      </Add>
    </Configuration>
    ```

1. 値を変更 `Version` します。

    - エッジ レガシを使用するバージョンをインストールするには、に変更します `16.0.11929.20946` 。
    - このバージョンを使用するバージョンをインストールInternet Explorerに変更します `16.0.10730.20348` 。

1. 必要に応じて `OfficeClientEdition` 、32 ビット Office をインストールする値を変更し、必要に応じて値を変更して、Office言語でインストール `"32"` `Language ID` します。
1. 管理者としてコマンド プロンプト *を開きます*。
1. ファイルを含むフォルダー `setup.exe` に `config.xml` 移動します。
1. 次のコマンドを実行します。

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    このコマンドは、Office。 このプロセスには数分かかる場合があります。

1. [キャッシュをOfficeします](../testing/clear-cache.md)。

> [!IMPORTANT]
> インストール後、Office の自動更新をオフにし、Office が使用を完了する前に動作する webview を使用しないバージョンに更新されません。 **これは、インストールから数分以内に発生する可能性があります。** 次の手順に従ってください。
>
> 1. 任意のアプリケーションOfficeし、新しいドキュメントを開きます。
> 1. リボンの **[ファイル]** タブを開き、[アカウント]**または [アカウント] Officeを** 選択 **します**。
> 1. [製品情報 **] 列で** 、[更新オプション] **を選択し**、[更新プログラムの無効化] **を選択します**。 このオプションを使用できない場合は、Office自動的に更新されない構成が既に設定されています。

古いバージョンの Office の使用が完了したら、ファイルを編集し、前にコピーしたビルド番号に変更して、新しいバージョンを再 `config.xml` `Version` インストールします。 次に、管理者 `setup.exe /configure config.xml` のコマンド プロンプトでコマンドを繰り返します。 必要に応じて、自動更新を再び有効にします。
