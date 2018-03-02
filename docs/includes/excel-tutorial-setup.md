開発プロジェクトを設定して、このチュートリアルを始めます。 

> [!TIP]
> まだ「[jQuery を使用する Excel アドインのクイック スタート](../quickstarts/excel-quickstart-jquery.md?tabs=visual-studio-code)」をご覧になっていなければ、お読みください。 具体的には、テスト用に Excel アドインをサイドロードする方法をしっかりと理解します。

## <a name="prerequisites"></a>前提条件

このチュートリアルを使用するには、以下のバージョンがインストールされている必要があります。 

- Excel 2016、バージョン 1711 (ビルド 8730.1000 クイック実行) 以降。 このバージョンを入手するには、Office Insider への参加が必要になることがあります。 詳細については、「[Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1)」を参照してください。
- [Node と npm](https://nodejs.org/en/) 
- [Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)

## <a name="setup"></a>セットアップ

1. 「[Excel アドインのチュートリアル](https://github.com/OfficeDev/Excel-Add-in-Tutorial)」で、GitHub リポジトリを複製します。
2. Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。
3. `npm install` コマンドを実行して、package.json ファイルに一覧表示されているツールとライブラリをインストールします。 
4. 「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」の手順を実行して、開発用コンピューターのオペレーティング システムの証明書を信頼します。

