---
title: Excel カスタム関数のチュートリアル
description: このチュートリアルでは、計算の実行、Web データの要求、Web データのストリームが可能なカスタム関数を含む Excel アドインを作成します。
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6ce3bbb4f36819599451f6f87db6c6a6f882f5a1
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275609"
---
# <a name="tutorial-create-custom-functions-in-excel"></a>チュートリアル: Excel でのカスタム関数の作成

カスタム関数では、関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。 ユーザーは、Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。 計算のような単純なタスク、または Web からワークシートへのデータのリアルタイム ストリーミングのようなより複雑なタスクを実行するカスタム関数を作成できます。

このチュートリアルの内容:
> [!div class="checklist"]
> * [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用して、カスタム関数アドインを作成します。 
> * あらかじめ用意されているカスタム関数を使用し、単純な計算を実行します。
> * Web からデータを取得するカスタム関数を作成します。
> * Web からデータをリアルタイムでストリーミングするカスタム関数を作成します。

## <a name="prerequisites"></a>前提条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* Windows または Web 上での Excel (Office 365 サブスクリプションに接続されているバージョン 1904 以降)

## <a name="create-a-custom-functions-project"></a>カスタム関数プロジェクトを作成する

 まず、カスタム関数アドインをビルドするコード プロジェクトを作成します。 Office アドインの [Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)は、試すことができるいくつかのカスタム関数を使ってプロジェクトをセットアップします。カスタム関数のクイック スタートをすでに実行し、プロジェクトを生成している場合は、そのプロジェクトを引き続き使用し、代わりに[この手順](#create-a-custom-function-that-requests-data-from-the-web) に進みます。

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]
    
    * **Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project`
    * **Choose a script type: (スクリプトの種類を選択)** `JavaScript`
    * **What would you want to name your add-in?: (アドインの名前を何にしますか)** `starcount`

    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/starcountPrompt.png)
    
    Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

2. プロジェクトのルート フォルダーに移動します。
    
    ```command&nbsp;line
    cd starcount
    ```

3. プロジェクトをビルドします。
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Office アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。 `npm run build`の実行後に証明書をインストールするように指示が出された場合は、Yeomanジェネレーターが提供する証明書をインストールする手順に従ってください。

4. Node.js で実行しているローカル Web サーバーを開始します。 Web または Windows 上の Excel でカスタム関数アドインを試すことができます。

# <a name="excel-on-windows-or-mac"></a>[Windows または Mac 上の Excel](#tab/excel-windows)

Windows または Mac の Excel でアドインをテストするには、次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動し、アドインが読み込まれたときに Excel が開きます。

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-web"></a>[Excel on the web](#tab/excel-online)

ブラウザーの Excel でアドインをテストするには、次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します。

```command&nbsp;line
npm run start:web
```

カスタム関数アドインを使用するには、Excel on the web で新しいブックを開きます。 このブックでアドインをサイドロードするには、次の手順を完了します。

1. Excel で、[**挿入**] タブを選択して、[**アドイン**] を選択します。

   ![[個人用アドイン] のアイコンが強調表示された Web 上の Excel の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)
   
2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a>あらかじめ用意されているカスタム関数を試す

作成したカスタム関数プロジェクトには、あらかじめ用意されているカスタム関数がいくつか含まれており、**./src/functions/functions.js** ファイル内で定義されています。 **./manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。 Excel でカスタム関数にアクセスするには、CONTOSO 名前空間を使用します。

その後、次の手順を実行し、`ADD` カスタム関数を試します。

1. Excel で、任意のセルに移動し、`=CONTOSO` と入力します。 `CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。

2. セル内で値 `=CONTOSO.ADD(10,200)` を入力して Enter キーを押し、入力パラメーターとして数値 `10` と `200` を指定して、`CONTOSO.ADD` 関数を実行します。

`ADD` カスタム関数によって、指定した 2 つの数字の合計が計算され、**210** という結果が返されます。

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Web からデータを要求するカスタム関数を作成する

Web からデータを統合することは、カスタム関数を使用して Excel を拡張する優れた方法です。 次に、特定の Github リポジトリが所有する星の数を示す `getStarCount` という名前のカスタム関数を作成します。

1. **starcount** プロジェクトで **./src/functions/functions.js** ファイルを見つけ、それをコード エディターで開きます。 

2. **function.js** で、次のコードを追加します。 

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

3. 次のコマンドを実行してプロジェクトを再構築します。

    ```command&nbsp;line
    npm run build
    ```

4. Excel のアドインを再登録するには、次の手順を完了します (Web、Windows または Mac 上の Excel の場合)。 新しい関数を使用するには、次の手順を完了する必要があります。

### <a name="excel-on-windows-or-mac"></a>[Windows または Mac 上の Excel](#tab/excel-windows)

1. Excel を閉じて再び開きます。

2. Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 上の Excel の [挿入] リボン](../images/select-insert.png)

3. 使用可能なアドインのリストから [**開発者向けアドイン**] セクションを見つけ、**starcount** アドインを選択して登録します。
    ![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 上の Excel の [挿入] リボン](../images/list-starcount.png)


# <a name="excel-on-the-web"></a>[Excel on the web](#tab/excel-online)

1. Excel で [**挿入**] タブを選択し、[**アドイン**] を選択します。![[個人用アドイン] アイコンが強調表示されている Excel on the web の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)

2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

---

<ol start="5">
<li> 新しい関数をお試しください。 セル <strong>B1</strong> で、テキスト <strong>=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")</strong> を入力し、Enter キーを押します。 セル <strong>B1</strong> の結果は [Excel-Custom-Functions Github リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions) に与えられた現在の星の数です。</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>非同期でデータをストリーミングするカスタム関数を作成する

`getStarCount` 関数は、ある時点でリポジトリに存在する星の数を返します。 カスタム関数は、継続的に変更されているデータも返すことができます。 これらの関数は、ストリーミング関数と呼ばれます。 関数が呼び出されたセルを参照する `invocation` パラメーターを含める必要があります。 `invocation` パラメーターは、セルの内容をいつでも更新するために使用します。  

次のコード例では、`currentTime` と `clock` という 2 つの関数があることがわかります。 `currentTime` 関数は、ストリーミングを使わない静的な関数です。 日付を表す文字列を返します。 `clock` 関数は、`currentTime` 関数を使用して、Excel 内のセルに毎秒新しい時間を提します。 時間を Excel セルに配信するには `invocation.setResult` を使用し、関数がキャンセルされたときに発生する処理には `invocation.onCanceled` を使用します。

1. **starcount**プロジェクトで、次のコードを **./src/functions/functions.js** に追加し、ファイルを保存します。

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

2. 次のコマンドを実行してプロジェクトを再構築します。

    ```command&nbsp;line
    npm run build
    ```

3. Excel のアドインを再登録するには、次の手順を完了します (Web、Windows または Mac 上の Excel の場合)。 新しい関数を使用するには、次の手順を完了する必要があります。 

# <a name="excel-on-windows-or-mac"></a>[Windows または Mac 上の Excel](#tab/excel-windows)

1. Excel を閉じて再び開きます。

2. Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 上の Excel の [挿入] リボン](../images/select-insert.png)

3. 使用可能なアドインのリストから [**開発者向けアドイン**] セクションを見つけ、**starcount** アドインを選択して登録します。
    ![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 上の Excel の [挿入] リボン](../images/list-starcount.png)

# <a name="excel-on-the-web"></a>[Excel on the web](#tab/excel-online)

1. Excel で [**挿入**] タブを選択し、[**アドイン**] を選択します。![[個人用アドイン] アイコンが強調表示されている Excel on the web の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)

2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

--- 

<ol start="4">
<li>新しい関数をお試しください。 セル <strong>C1</strong> にテキスト <strong>=CONTOSO.CLOCK()</strong> を入力し、Enter キーを押します。 現在の日付が表示されます。この日付は 1 秒ごとにアップデートされます。 このクロックはループ上の単なるタイマーですが、リアルタイム データの Web 要求を行うより複雑な関数にタイマーを設定するという同じ考え方を使用できます。</li>
</ol>

## <a name="next-steps"></a>次の手順

おめでとうございます! 新しいカスタム関数プロジェクトを作成し、あらかじめ用意されている関数を試し、Web にデータを要求するカスタム関数を作成し、ストリーミング データであるカスタム関数を作成しました。 次に、共有ランタイムを使用するようにプロジェクトを変更することで、作業ウィンドウの操作が容易になります。 次の記事の手順を実行します。

> [!div class="nextstepaction"]
> [共有ランタイムを使用するようにアドインを構成する](../excel/configure-your-add-in-to-use-a-shared-runtime.md)
