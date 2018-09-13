開発プロジェクトを設定して、このチュートリアルを始めます。 

> [!NOTE]
> このページでは、Excel のアドインのチュートリアルの個々 の手順について説明します。 このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[Excel アドインのチュートリアル](../tutorials/excel-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="prerequisites"></a>前提条件

このチュートリアルを使用するには、以下のバージョンがインストールされている必要があります。 

- Excel 2016、バージョン 1711 (ビルド 8730.1000 クイック実行) 以降。 このバージョンを入手するには、Office Insider への参加が必要になることがあります。 詳細については、「[Office Insider](https://products.office.com/office-insider?tab=tab-1)」を参照してください。
- [Node と npm](https://nodejs.org/en/) 
- [Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)

## <a name="setup"></a>セットアップ

1. 「[Excel アドインのチュートリアル](https://github.com/OfficeDev/Excel-Add-in-Tutorial)」で、GitHub リポジトリを複製します。

2. Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。

3. コマンドを実行して、package.json ファイルに一覧表示されているツールとライブラリをインストールします。`npm install` 

4. 「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」の手順を実行して、開発用コンピューターのオペレーティング システムの証明書を信頼します。

