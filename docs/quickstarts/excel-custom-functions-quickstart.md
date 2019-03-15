---
ms.date: 03/06/2019
description: Excel クイックスタートガイドでのカスタム関数の開発。
title: カスタム関数クイックスタート (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 9dd3e5a99f08ce0b931e705fac3312ab10c19e18
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/14/2019
ms.locfileid: "30632703"
---
# <a name="get-started-developing-excel-custom-functions"></a>Excel カスタム関数の開発を始める

カスタム関数を使用すると、開発者は、JavaScript または Typescript でアドインの一部として定義することによって、Excel に新しい関数を追加できるようになります。 excel ユーザーは、excel の任意のネイティブ関数の場合と同じように、カスタム`SUM()`関数にアクセスできます。

## <a name="prerequisites"></a>前提条件

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

カスタム関数の作成を開始するには、次のツールと関連するリソースが必要です。

- [Node.js](https://nodejs.org/en/) (バージョン 8.0.0 以降)

- [Git バッシュ](https://git-scm.com/downloads) (または別の Git クライアント)

- 最新バージョンの [Yeoman](https://yeoman.io/) と [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)。これらのツールをグローバルにインストールするには、コマンド プロンプトから次のコマンドを実行します。

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > 以前に一度使用したバージョンのジェネレーターをインストールしていた場合でも、パッケージを npm から最新バージョンに更新することをお勧めします。

## <a name="build-your-first-custom-functions-project"></a>最初のカスタム関数プロジェクトを作成する

はじめに、Yeoman ジェネレーターを使って、カスタム関数プロジェクトを作成します。 これにより、カスタム関数のコーディングを開始するための正しいフォルダー構造、ソース ファイル、依存関係によるプロジェクトがセットアップされます。

1. 次のコマンドを実行し、以下のようにプロンプトに応答します。

    ```
    yo office
    ```

    - Choose a project type (プロジェクトの種類を選択): `Excel Custom Functions Add-in project (...)`

    - Choose a script type (スクリプトの種類を選択): `JavaScript`

    - What would you want to name your add-in? (アドインの名前を何にしますか) `stock-ticker`

    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/12-10-fork-cf-pic.jpg)

    Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。

2. 作成したばかりのプロジェクトフォルダーに移動します。

    ```
    cd stock-ticker
    ```

3. このプロジェクトを実行するには、自己署名証明書を信頼する必要があります。 Windows または Mac についての詳細な手順については、「[自己署名証明書を信頼済みルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」を参照してください。  

4. プロジェクトをビルドします。

    ```
    npm run build
    ```

5. Node.js で実行しているローカル Web サーバーを開始します。

    - Windows 版 excel を使用してカスタム関数をテストする場合は、次のコマンドを実行してローカル web サーバーを起動し、Excel を起動して、アドインをサイドロードします。

        ```
         npm run start
        ```
        このコマンドを実行すると、コマンドプロンプトに web サーバーの起動に関する詳細が表示されます。 Excel は、アドインが読み込まれた状態で起動します。 アドインが読み込まれない場合は、手順 3 が正しく完了しているか確認してください。

    - Excel Online を使用してカスタム関数をテストする場合は、次のコマンドを実行してローカル web サーバーを開始します。

        ```
        npm run start-web
        ```

         このコマンドを実行すると、コマンドプロンプトに web サーバーの起動に関する詳細が表示されます。 関数を使用するには、Excel Online で新しいブックを開きます。 このブックでは、アドインを読み込む必要があります。 

        これを行うには、リボンの [**挿入**] タブを選択して、[アドインの**取得**] を選択します。生成された新しいウィンドウで、[**マイアドイン**] タブが表示されていることを確認します。次に、[**個人用アドインの管理 > [個人用**アドインのアップロード] を選択します。 マニフェストファイルを参照してアップロードします。 アドインが読み込まれない場合は、手順3が正しく完了していることを確認してください。

## <a name="try-out-the-prebuilt-custom-functions"></a>あらかじめ用意されているカスタム関数を試してみる

Yeoman ジェネレーターで作成したカスタム関数プロジェクトには、あらかじめ用意されているカスタム関数がいくつか含まれており、**src/customfunctions.js** ファイル内で定義されています。 プロジェクトのルート ディレクトリの **manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。

Excel ブックで、次の手順を`ADD`実行してカスタム関数を試してみます。

1. セルを選択し、 `=CONTOSO`テキストを入力します。 `CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。

2. セルに`CONTOSO.ADD`値`=CONTOSO.ADD(10,200)`を入力し`10` 、 `200` enter キーを押して、数値と入力パラメーターを使用して、関数を実行します。

`ADD` カスタム関数によって、入力パラメーターとして指定した 2 つの数字の合計が計算されます。 「`=CONTOSO.ADD(10,200)`」と入力して Enter キーを押すと、**210** という結果が生成されるはずです。

## <a name="next-steps"></a>次の手順

おめでとうございます。 Excel アドインでカスタム関数が正常に作成されました。 次に、ストリーミングデータ機能を使用して、より複雑なアドインをビルドします。 次のリンクでは、「カスタム関数を使用した Excel アドインのチュートリアル」の次の手順を実行します。

> [!div class="nextstepaction"]
> [Excel カスタム関数アドインのチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a>関連項目

* [カスタム関数の概要](../excel/custom-functions-overview.md)
* [カスタム関数のメタデータ](../excel/custom-functions-json.md)
* [Excel カスタム関数のランタイム](../excel/custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](../excel/custom-functions-best-practices.md)
