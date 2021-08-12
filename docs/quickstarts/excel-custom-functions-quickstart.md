---
ms.date: 08/04/2021
description: Excel カスタム関数開発のためのクイック スタート ガイド。
title: カスタム関数クイック スタート
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6c463c494bf3175309226d72d0ca95417a3889b392a4f43035cd5d50263d8fbf
ms.sourcegitcommit: f5d4321763e366a10f2d868fb329dbef5239c830
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "57845607"
---
# <a name="get-started-developing-excel-custom-functions"></a>Excel カスタム関数の開発を開始する

カスタム関数機能により、開発者は、アドインの一部としてカスタム関数を JavaScript または Typescript で定義することによって、新しい関数を Excel に追加できるようになりました。 Excel のユーザーは、`SUM()` など、Excel のすべてのネイティブ関数にアクセスするとの同じようにカスタム関数にアクセスできます。

## <a name="prerequisites"></a>前提条件

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Windows 版 Excel (Microsoft 365 サブスクリプションに接続されている、バージョン 1904 以降) または Excel on the web
- Excel カスタム関数は (Microsoft 365 サブスクリプションに接続されている) Mac 版 Office でサポートされており、このチュートリアルはまもなく更新されます。

>[!NOTE]
>Excel カスタム関数は Office 2019 (1 回限りの購入) ではサポートされていません。

## <a name="build-your-first-custom-functions-project"></a>カスタム関数プロジェクトを初めて作成する

はじめに、Yeoman ジェネレーターを使って、カスタム関数プロジェクトを作成します。 これにより、カスタム関数のコーディングを開始するための正しいフォルダー構造、ソース ファイル、依存関係によるプロジェクトがセットアップされます。

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project`
    - **Choose a script type: (スクリプトの種類を選択)** `JavaScript`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `starcount`

    ![カスタム関数プロジェクトの Yeoman Office アドイン ジェネレーター コマンドライン インターフェイス プロンプトのスクリーンショット。](../images/starcountPrompt.png)

    Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。

1. Yeoman ジェネレーターによりプロジェクトの作業に関する手順がコマンド ライン内にいくつか示されますが、これらは無視し、引き続き指示に従います。プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd starcount
    ```

1. プロジェクトをビルドします。

    ```command&nbsp;line
    npm run build
    ```

1. Node.js で実行しているローカル Web サーバーを開始します。 カスタム関数アドインは Web 版 Excel または Windows 版 Excel で試すことができます。 アドインの作業ウィンドウを開くように求められる場合がありますが、これは省略可能です。 カスタム関数はアドインの作業ウィンドウを開かなくても実行できます。

# <a name="excel-on-windows"></a>[Windows 版 Excel](#tab/excel-windows)

アドインを Windows 版 Excel で試すには、次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動し、アドインが読み込まれた状態で Excel が開きます。

```command&nbsp;line
npm run start:desktop
```

> [!NOTE]
> Office アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。 `npm run start`の実行後に証明書をインストールするように指示が出された場合は、Yeomanジェネレーターが提供する証明書をインストールする手順に従ってください。
    
# <a name="excel-on-the-web"></a>[Web 版 Excel](#tab/excel-online)

アドインを Web 版 Excel で試すには、次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します。

```command&nbsp;line
npm run start:web
```

> [!NOTE]
> Office アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。 `npm run start`の実行後に証明書をインストールするように指示が出された場合は、Yeomanジェネレーターが提供する証明書をインストールする手順に従ってください。

カスタム関数アドインを使用するには、ブラウザー上の Excel で新しいブックを開きます。 このブックで次の手順を実行してアドインをサイドロードします。

1. Excel で、[**挿入**] タブを選択して、[**アドイン**] を選択します。

   ![[個人用アドイン] ボタンが強調表示された Excel on the web の [挿入] リボンのスクリーンショット。](../images/excel-cf-online-register-add-in-1.png)

1. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

1. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

1. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

---

## <a name="try-out-a-prebuilt-custom-function"></a>既製のカスタム関数を試す

Yeoman ジェネレーター使用して作成したカスタム関数プロジェクトには既製のカスタム関数がいくつか含まれており、これらは **./src/functions/functions.js** ファイル内で定義されています。 カスタム関数はすべて `CONTOSO` 名前空間に属するということは、プロジェクトのルート ディレクトリの **./manifest.xml** ファイルで指定されています。

Excel ブックで次の手順を実行し、`ADD` カスタム関数を試してみてください。

1. セルを 1 つ選択し、「`=CONTOSO`」と入力します。 `CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。

1. セル内に「`=CONTOSO.ADD(10,200)`」という値を入力して Enter キーを押し、入力パラメーターとして数値「`10`」 と「`200`」を指定して、`CONTOSO.ADD` 関数を実行します。

`ADD` カスタム関数によって、入力パラメーターとして指定した 2 つの数字の合計が計算されます。 「`=CONTOSO.ADD(10,200)`」と入力して Enter キーを押すと、**210** という結果が生成されるはずです。

## <a name="next-steps"></a>次の手順

これで、カスタム関数が Excel アドイン内に正常に作成されました。 次は、ストリーミング データ機能を使用してより複雑なアドインを作成してください。 カスタム関数を使用した Excel アドインのチュートリアルの次の手順を確認するには、次のリンクをクリックしてください。

> [!div class="nextstepaction"]
> [Excel カスタム関数アドインのチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web)

## <a name="see-also"></a>関連項目

- [カスタム関数の概要](../excel/custom-functions-overview.md)
- [カスタム関数のメタデータ](../excel/custom-functions-json.md)
- [Excel カスタム関数のランタイム](../excel/custom-functions-runtime.md)
