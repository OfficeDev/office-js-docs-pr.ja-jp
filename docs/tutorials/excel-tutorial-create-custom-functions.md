---
title: Excel カスタム関数のチュートリアル
description: このチュートリアルでは、計算の実行、Web データの要求、Web データのストリームが可能なカスタム関数を含む Excel アドインを作成します。
ms.date: 03/23/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 984a2090a70360af4f361bb531190a7ee2a05c4c
ms.sourcegitcommit: 64942cdd79d7976a0291c75463d01cb33a8327d8
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/25/2022
ms.locfileid: "64404731"
---
# <a name="tutorial-create-custom-functions-in-excel"></a>チュートリアル: Excel でのカスタム関数の作成

カスタム関数では、関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。 ユーザーは、Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。 計算のような単純なタスク、または Web からワークシートへのデータのリアルタイム ストリーミングのようなより複雑なタスクを実行するカスタム関数を作成できます。

このチュートリアルの内容:
> [!div class="checklist"]
> - [Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用して、カスタム関数アドインを作成します。
> - あらかじめ用意されているカスタム関数を使用し、単純な計算を実行します。
> - Web からデータを取得するカスタム関数を作成します。
> - Web からデータをリアルタイムでストリーミングするカスタム関数を作成します。

## <a name="prerequisites"></a>前提条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Microsoft 365 サブスクリプションに接続されている Office (Office for the web を含む)。

  > [!NOTE]
  > Office をまだお持ちでない場合は、[Microsoft 365 開発者プログラムに参加](https://developer.microsoft.com/office/dev-program)して、開発中に使用できる 90 日間更新可能な無料の Microsoft 365 サブスクリプションを取得できます。

## <a name="create-a-custom-functions-project"></a>カスタム関数プロジェクトを作成する

 まず、カスタム関数アドインをビルドするコード プロジェクトを作成します。[Office アドインの Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)は、試すことができるいくつかのカスタム関数を使ってプロジェクトをセットアップします。カスタム関数のクイック スタートを既に実行し、プロジェクトを生成している場合は、そのプロジェクトを引き続き使用し、代わりに[この手順](#create-a-custom-function-that-requests-data-from-the-web)に進みます。

> [!NOTE]
> Office Project を再作成すると、Office キャッシュに同じ名前の関数のインスタンスが既に含まれているため、エラーが発生する可能性があります。`npm run start` を実行する前に [Office キャッシュをクリアする](../testing/clear-cache.md)ことにより、これを防ぐことができます。

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project`
    - **Choose a script type: (スクリプトの種類を選択)** `JavaScript`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `starcount`

    :::image type="content" source="../images/starcountPrompt.png" alt-text="カスタム関数プロジェクトの Yeoman Office アドイン ジェネレーター コマンドライン インターフェイス プロンプトのスクリーンショット。":::

    Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd starcount
    ```

1. プロジェクトをビルドします。

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > 開発の最中でも、Office アドインでは HTTP ではなく HTTPS を使用する必要があります。`npm run build` を実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。

1. Node.js で実行しているローカル Web サーバーを開始します。 Excel でカスタム関数アドインを試すことができます。

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

## <a name="try-out-a-prebuilt-custom-function"></a>あらかじめ用意されているカスタム関数を試す

作成したカスタム関数プロジェクトには、あらかじめ用意されているカスタム関数がいくつか含まれており、**./src/functions/functions.js** ファイル内で定義されています。 **./manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。 Excel でカスタム関数にアクセスするには、CONTOSO 名前空間を使用します。

次に、以下の手順を実行し、`ADD` カスタム関数を試してみてください。

1. Excel で、任意のセルに移動し、`=CONTOSO` と入力します。 `CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。

1. セル内で値 `=CONTOSO.ADD(10,200)` を入力して Enter キーを押し、入力パラメーターとして数値 `10` と `200` を指定して、`CONTOSO.ADD` 関数を実行します。

`ADD` カスタム関数によって、指定した 2 つの数字の合計が計算され、**210** という結果が返されます。

[!include[Manually register an add-in](../includes/excel-custom-functions-manually-register.md)]

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Web からデータを要求するカスタム関数を作成する

Web からデータを統合することは、カスタム関数を使用して Excel を拡張する優れた方法です。 次に、特定の Github リポジトリが所有する星の数を示す `getStarCount` という名前のカスタム関数を作成します。

1. **starcount** プロジェクトで **./src/functions/functions.js** ファイルを見つけ、それをコード エディターで開きます。

1. **function.js** で、次のコードを追加します。

    ```JS
    /**
      * Gets the star count for a given Github repository.
      * @customfunction 
      * @param {string} userName string name of Github user or organization.
      * @param {string} repoName string name of the Github repository.
      * @return {number} number of stars given to a Github repository.
      */
      async function getStarCount(userName, repoName) {
        try {
          //You can change this URL to any web request you want to work with.
          const url = "https://api.github.com/repos/" + userName + "/" + repoName;
          const response = await fetch(url);
          //Expect that status code is in 200-299 range
          if (!response.ok) {
            throw new Error(response.statusText)
          }
            const jsonResponse = await response.json();
            return jsonResponse.watchers_count;
        }
        catch (error) {
          return error;
        }
      }
    ```

1. 次のコマンドを実行してプロジェクトを再構築します。

    ```command&nbsp;line
    npm run build
    ```

1. Excel のアドインを再登録するには、次の手順を完了します (Web、Windows または Mac 上の Excel の場合)。 新しい関数を使用するには、次の手順を完了する必要があります。

### <a name="excel-on-windows-or-mac"></a>[Windows または Mac 上の Excel](#tab/excel-windows)

1. Excel を閉じて再び開きます。

1. Excel で、**[挿入]** タブを選択し、**[個人用アドイン]** の右側に配置された下矢印を選択します。

    :::image type="content" source="../images/select-insert.png" alt-text="[個人用アドイン] の下矢印が強調表示された Windows での Excel の [挿入] リボンのスクリーンショット":::

1. 使用可能なアドインのリストから [**開発者向けアドイン**] セクションを見つけ、**starcount** アドインを選択して登録します。

    :::image type="content" source="../images/list-starcount.png" alt-text="[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 上の Excel の [挿入] リボンのスクリーンショット。":::

# <a name="excel-on-the-web"></a>[Web 版 Excel](#tab/excel-online)

1. Excel で、[**挿入**] タブを選択して、[**アドイン**] を選択します。

    :::image type="content" source="../images/excel-cf-online-register-add-in-1.png" alt-text="[個人用アドイン] ボタンが強調表示された Excel on the web の [挿入] リボンのスクリーンショット。":::

1. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

1. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

1. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

1. 新しい関数をお試しください。 セル **B1** で、テキスト **=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")** を入力し、Enter キーを押します。 セル **B1** の結果は [Excel-Custom-Functions Github リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions) に与えられた現在の星の数です。

---

## <a name="create-a-streaming-asynchronous-custom-function"></a>非同期でデータをストリーミングするカスタム関数を作成する

`getStarCount` 関数は、ある時点でリポジトリに存在する星の数を返します。 カスタム関数は、継続的に変更されているデータも返します。 これらの関数は、ストリーミング関数と呼ばれます。 関数を呼び出したセルを参照する `invocation` パラメーターを含める必要があります。 `invocation` パラメーターは、セルのコンテンツをいつでも更新するために使用します。  

次のコード例では、`currentTime` と `clock` という 2 つの関数があることがわかります。 `currentTime` 関数は、ストリーミングを使わない静的な関数です。 日付を表す文字列を返します。 `clock` 関数は、`currentTime` 関数を使用して、Excel 内のセルに毎秒新しい時間を提します。 `invocation.setResult`を使用して Excel セルに時間を配信し、関数のキャンセルを処理する`invocation.onCanceled`を使用します。 

**starcount** プロジェクトには、**./src/functions/functions.js** ファイルに次の 2 つの関数が既に含まれています。

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}
    
/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);
    
  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

関数を試すには、セル **C1** に、テキスト **=CONTOSO.CLOCK()** を入力して、Enter キーを押します。 現在の日付が表示されます。この日付は 1 秒ごとにアップデートされます。 このクロックはループ上の単なるタイマーですが、リアルタイム データの Web 要求を行うより複雑な関数にタイマーを設定するという同じ考え方を使用できます。

## <a name="next-steps"></a>次の手順

おめでとうございます! 新しいカスタム関数プロジェクトを作成し、あらかじめ用意されている関数を試し、Web にデータを要求するカスタム関数を作成し、ストリーミング データであるカスタム関数を作成しました。 次に、共有ランタイムを使用するようにプロジェクトを変更することで、関数が作業ウィンドウを操作しやすくなります。 以下の記事の手順に従ってください。

> [!div class="nextstepaction"]
> [共有ランタイムを使用するようにアドインを構成する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
