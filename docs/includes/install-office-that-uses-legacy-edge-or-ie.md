次の手順に従って、Microsoft Edge 従来版 Webview (EdgeHTML) を使用してアドインを実行するサブスクリプション Officeのバージョンまたは Internet Explorer (Trident) を使用するバージョンをインストールします。

1. 任意のOfficeアプリケーションで、リボンの [**ファイル**] タブを開き、**アカウントまたはアカウントOffice** 選択します。 [**_ホスト名_ の概要**] ボタン ([**Word** について] など) を選択します。
1. 開いたダイアログで、完全な xx.x.xxxxx.xxxxx ビルド番号を見つけて、そのコピーをどこかに作成します。
1. [Office 展開ツール](https://www.microsoft.com/download/details.aspx?id=49117)をダウンロードします。
1. ダウンロードしたファイルを実行してツールを抽出します。 ツールをインストールする場所を選択するように求められます。
1. ツールをインストールしたフォルダー (ファイルがある場所 `setup.exe` ) で、名前 `config.xml` を含むテキスト ファイルを作成し、次の内容を追加します。

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

    - Edge Legacy を使用するバージョンをインストールするには、次のように `16.0.11929.20946`変更します。
    - Internet Explorer を使用するバージョンをインストールするには、次のように変更します `16.0.10730.20348`。

1. 必要に応じて、32 ビット Officeをインストールするように`"32"`値`OfficeClientEdition`を変更し、必要に応じて値を`Language ID`変更して、別の言語でOfficeをインストールします。
1. *管理者として* コマンド プロンプトを開きます。
1. ファイルを含むフォルダーに`setup.exe``config.xml`移動します。
1. 次のコマンドを実行します。

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    このコマンドは、Officeをインストールします。 このプロセスには数分かかる場合があります。

1. [Office キャッシュをクリアします](../testing/clear-cache.md)。

> [!IMPORTANT]
> インストール後は、Officeの自動更新をオフにし、使用を完了する前に作業する Web ビューを使用しないバージョンにOfficeが更新されないようにしてください。 **これは、インストールから数分以内に発生する可能性があります。** 次の手順に従ってください。
>
> 1. Officeアプリケーションを起動し、新しいドキュメントを開きます。
> 1. リボンの [**ファイル**] タブを開き、**アカウントまたはアカウントOffice** 選択します。
> 1. [ **Product Information** ] 列で、[ **更新オプション**] を選択し、[ **更新プログラムの無効化**] を選択します。 このオプションを使用できない場合、Officeは自動的に更新されないように既に構成されています。

古いバージョンのOfficeの使用が完了したら、ファイルを編集`config.xml`し、前にコピーしたビルド番号に変更`Version`して、新しいバージョンを再インストールします。 次に、管理者コマンド `setup.exe /configure config.xml` プロンプトでコマンドを繰り返します。 必要に応じて、自動更新を再度有効にします。
