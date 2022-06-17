---
ms.date: 06/10/2022
description: Excel カスタム関数開発のためのクイック スタート ガイド。
title: カスタム関数クイック スタート
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: aa44caf014a6d617112a616e96e1c67079c4c385
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091084"
---
# <a name="get-started-developing-excel-custom-functions"></a>Excel カスタム関数の開発を開始する

カスタム関数機能を使用すると、開発者は、アドインの一部としてカスタム関数を JavaScript または TypeScript で定義することによって、新しい関数を Excel に追加できます。 Excel のユーザーは、`SUM()` など、Excel のすべてのネイティブ関数にアクセスするとの同じようにカスタム関数にアクセスできます。

## <a name="prerequisites"></a>前提条件

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Microsoft 365 サブスクリプションに接続されている Office (Office for the web を含む)。

  > [!NOTE]
  > Office をまだお持ちでない場合は、[Microsoft 365 開発者プログラムに参加](https://developer.microsoft.com/office/dev-program)して、開発中に使用できる 90 日間更新可能な無料の Microsoft 365 サブスクリプションを取得できます。

## <a name="build-your-first-custom-functions-project"></a>カスタム関数プロジェクトを初めて作成する

はじめに、Yeoman ジェネレーターを使って、カスタム関数プロジェクトを作成します。 これにより、カスタム関数のコーディングを開始するための正しいフォルダー構造、ソース ファイル、依存関係によるプロジェクトがセットアップされます。

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project`
    - **Choose a script type: (スクリプトの種類を選択)** `JavaScript`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My custom functions add-in`

    :::image type="content" source="../images/yo-office-excel-cf-quickstart.png" alt-text="カスタム関数プロジェクトの Yeoman Office アドイン ジェネレーター コマンドライン インターフェイス プロンプトのスクリーンショット。":::

    Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。

1. Yeoman ジェネレーターによりプロジェクトの作業に関する手順がコマンド ライン内にいくつか示されますが、これらは無視し、引き続き指示に従います。プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd "My custom functions add-in"
    ```

1. プロジェクトをビルドします。

    ```command&nbsp;line
    npm run build
    ```

1. Node.js で実行しているローカル Web サーバーを開始します。 Excel でカスタム関数アドインを試すことができます。 アドインの作業ウィンドウを開くように求められる場合がありますが、これは省略可能です。 カスタム関数はアドインの作業ウィンドウを開かなくても実行できます。

# <a name="excel-on-windows-or-mac"></a>[Windows または Mac 上の Excel](#tab/excel-windows)

Windows または Mac の Excel でアドインをテストするには、次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動し、アドインが読み込まれたときに Excel が開きます。

```command&nbsp;line
npm run start:desktop
```

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

# <a name="excel-on-the-web"></a>[Web 版 Excel](#tab/excel-online)

アドインを Web 版 Excel で試すには、次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します。 "{url}" を、アクセス許可を持っている OneDrive または SharePoint ライブラリ上の Excel ドキュメントの URL に置き換えます。

[!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

---

## <a name="try-out-a-prebuilt-custom-function"></a>既製のカスタム関数を試す

Yeoman ジェネレーター使用して作成したカスタム関数プロジェクトには既製のカスタム関数がいくつか含まれており、これらは **./src/functions/functions.js** ファイル内で定義されています。 カスタム関数はすべて `CONTOSO` 名前空間に属するということは、プロジェクトのルート ディレクトリの **./manifest.xml** ファイルで指定されています。

Excel ブックで次の手順を実行し、`ADD` カスタム関数を試してみてください。

1. セルを選択して、`=CONTOSO` と入力します。`CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。

1. セル内に「`=CONTOSO.ADD(10,200)`」という値を入力して Enter キーを押し、入力パラメーターとして数値「`10`」 と「`200`」を指定して、`CONTOSO.ADD` 関数を実行します。

`ADD` カスタム関数によって、入力パラメーターとして指定した 2 つの数字の合計が計算されます。 「`=CONTOSO.ADD(10,200)`」と入力して Enter キーを押すと、**210** という結果が生成されるはずです。

[!include[Manually register an add-in](../includes/excel-custom-functions-manually-register.md)]

## <a name="next-steps"></a>次の手順

これで、カスタム関数が Excel アドイン内に正常に作成されました。 次は、ストリーミング データ機能を使用してより複雑なアドインを作成してください。 カスタム関数を使用した Excel アドインのチュートリアルの次の手順を確認するには、次のリンクをクリックしてください。

> [!div class="nextstepaction"]
> [Excel カスタム関数アドインのチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web)

## <a name="troubleshooting"></a>トラブルシューティング

クイック スタートを複数回実行すると、問題が発生する場合があります。 Office キャッシュに同じ名前を持つ関数のインスタンスが既に存在する場合、アドインのサイドロード時にエラーが発生します。 `npm run start` を実行する前に [Office キャッシュをクリアする](../testing/clear-cache.md)ことにより、これを防ぐことができます。

:::image type="content" source="../images/custom-function-already-exists-error.png" alt-text="Excelで '関数のインストール中にエラーが発生しました' というタイトルのエラー メッセージが表示されます。これには、'同じ名前を持つカスタム関数が既に存在するため、このアドインはインストールされませんでした' というテキストが含まれます。":::

## <a name="see-also"></a>関連項目

- [カスタム関数の概要](../excel/custom-functions-overview.md)
- [カスタム関数のメタデータ](../excel/custom-functions-json.md)
- [Excel カスタム関数のランタイム](../excel/custom-functions-runtime.md)
- [Visual Studio コードを使用して発行する](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
