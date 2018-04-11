開発プロジェクトを設定して、このチュートリアルを始めます。 

> [!NOTE]
> このページでは、Word アドインのチュートリアルの個々の手順について説明します。このページに検索エンジンの結果から、またはその他の直接リンクからアクセスした場合は、「[Word アドインのチュートリアル](../tutorials/word-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

> [!TIP]
> 「[最初の Word アドインをビルドする](../quickstarts/word-quickstart.md?tabs=visual-studio-code)」をまだ読んでいない場合は、最初にその記事をご確認ください。 具体的には、テスト用に Word アドインをサイドロードする方法をしっかりと理解します。

## <a name="prerequisites"></a>前提条件

このチュートリアルを使用するには、以下のバージョンがインストールされている必要があります。 

- Word 2016、バージョン 1711 (ビルド 8730.1000 クイック実行) 以降。 このバージョンを入手するには、Office Insider への参加が必要になることがあります。 詳細については、「[Office Insider](https://products.office.com/ja-jp/office-insider?tab=tab-1)」を参照してください。
- [Node と npm](https://nodejs.org/en/) 
- [Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)

## <a name="setup"></a>セットアップ

1. 「[Word アドインのチュートリアル](https://github.com/OfficeDev/Word-Add-in-Tutorial)」で、GitHub リポジトリを複製します。
2. Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。
3. `npm install` コマンドを実行して、package.json ファイルに一覧表示されているツールとライブラリをインストールします。 
4. 「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」の手順を実行して、開発用コンピューターのオペレーティング システムの証明書を信頼します。

