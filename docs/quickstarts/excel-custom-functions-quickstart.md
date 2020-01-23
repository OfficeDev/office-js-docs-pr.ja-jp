---
ms.date: 01/16/2020
description: Excel クイックスタートガイドでのカスタム関数の開発。
title: カスタム関数のクイックスタート
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7446de52832a70b30bbe39c71b37e80e006dc62f
ms.sourcegitcommit: 8bce9c94540ed484d0749f07123dc7c72a6ca126
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/22/2020
ms.locfileid: "41265565"
---
# <a name="get-started-developing-excel-custom-functions"></a>Excel カスタム関数の開発を始める

カスタム関数を使用すると、開発者は、JavaScript または Typescript でアドインの一部として定義することによって、Excel に新しい関数を追加できるようになります。 Excel ユーザーは、Excel の任意のネイティブ関数の場合と同じように、カスタム`SUM()`関数にアクセスできます。

## <a name="prerequisites"></a>前提条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* Windows 上の Excel (バージョン1904以降、Office 365 サブスクリプションに接続されている) または web 上の Excel
* Excel カスタム関数は Office on Mac でサポートされています (Office 365 サブスクリプションに接続されています)。また、このチュートリアルへの更新はまもなく公開されます。

>[!NOTE]
>Excel カスタム関数は Office 2019 (1 回限りの購入) ではサポートされていません。

## <a name="build-your-first-custom-functions-project"></a>最初のカスタム関数プロジェクトを作成する

はじめに、Yeoman ジェネレーターを使って、カスタム関数プロジェクトを作成します。 これにより、カスタム関数のコーディングを開始するための正しいフォルダー構造、ソース ファイル、依存関係によるプロジェクトがセットアップされます。

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project`
    - **Choose a script type: (スクリプトの種類を選択)** `JavaScript`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `starcount`

    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/starcountPrompt.png)

    Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。

2. [ごみ箱] ジェネレーターでは、プロジェクトの処理に関するいくつかの命令がコマンドラインに表示されますが、それらは無視して、手順に従って続行します。 プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd starcount
    ```

3. プロジェクトをビルドします。 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > 開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。 `npm run build`の実行後に証明書をインストールするように指示が出された場合は、Yeomanジェネレーターが提供する証明書をインストールする手順に従ってください。

4. Node.js で実行しているローカル Web サーバーを開始します。 Web または Windows 上の Excel でカスタム関数アドインを試すことができます。 アドインの作業ウィンドウを開くように求められる場合がありますが、これはオプションです。 アドインの作業ウィンドウを開かなくても、カスタム関数を実行できます。

# <a name="excel-on-windowstabexcel-windows"></a>[Excel on Windows](#tab/excel-windows)

Windows の Excel でアドインをテストするには、次のコマンドを実行します。 このコマンドを実行すると、ローカル web サーバーが起動し、アドインが読み込まれた状態で Excel が開きます。

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[Excel on the web](#tab/excel-online)

Web 上の Excel でアドインをテストするには、次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します。

```command&nbsp;line
npm run start:web
```

カスタム関数アドインを使用するには、ブラウザー上の Excel で新しいブックを開きます。 このブックでは、次の手順を実行して、アドインをサイドロードします。

1. Excel で、[**挿入**] タブを選択し、[**アドイン**] を選択します。

   ![[個人用アドイン] アイコンが強調表示されている web 上の Excel にリボンを挿入する](../images/excel-cf-online-register-add-in-1.png)
   
2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

---

## <a name="try-out-a-prebuilt-custom-function"></a>あらかじめ用意されているカスタム関数を試す

[ごみ箱] ジェネレーターを使用して作成したカスタム関数プロジェクトには、 **/src/functions/functions.js**ファイル内で定義されているいくつかのあらかじめ用意されたカスタム関数があります。 プロジェクトのルートディレクトリの **./manifest¥ xml**ファイルは、すべてのカスタム関数が`CONTOSO`名前空間に属することを指定します。

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