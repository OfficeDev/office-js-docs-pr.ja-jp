---
title: Excel カスタム関数のチュートリアル
description: このチュートリアルでは、計算の実行、Web データの要求、Web データのストリームが可能なカスタム関数を含む Excel アドインを作成します。
ms.date: 07/09/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2832df467f7e155ed026fe7f04837f18a4d2309d
ms.sourcegitcommit: 49af31060aa56c1e1ec1e08682914d3cbefc3f1c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/29/2019
ms.locfileid: "36672874"
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

* Windows 上の Excel (バージョン1904以降、Office 365 サブスクリプションに接続されている) または web

## <a name="create-a-custom-functions-project"></a>カスタム関数プロジェクトを作成する

 まず、カスタム関数アドインをビルドするコード プロジェクトを作成します。 [Office アドイン用の [ごみ箱] ジェネレーター](https://www.npmjs.com/package/generator-office)では、プロジェクトに事前に用意されているカスタム関数を使用してセットアップし、試すことができます。カスタム関数のクイックスタートを既に実行してプロジェクトを生成した場合は、そのプロジェクトを引き続き使用して、[この手順](#create-a-custom-function-that-requests-data-from-the-web)に進んでください。

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

1. 次のコマンドを実行し、以下のようにプロンプトに応答します。
    
    ```command&nbsp;line
    yo office
    ```
    
    * **Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project`
    * **Choose a script type: (スクリプトの種類を選択)** `JavaScript`
    * **What would you want to name your add-in?: (アドインの名前を何にしますか)** `starcount`

    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/starcountPrompt.png)
    
    Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。

2. プロジェクトのルート フォルダーに移動します。
    
    ```command&nbsp;line
    cd starcount
    ```

3. プロジェクトをビルドします。
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > 開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。 `npm run build`の実行後に証明書をインストールするように指示が出された場合は、Yeomanジェネレーターが提供する証明書をインストールする手順に従ってください。

4. Node.js で実行しているローカル Web サーバーを開始します。 Web または Windows 上の Excel でカスタム関数アドインを試すことができます。

# <a name="excel-on-windows-or-mactabexcel-windows"></a>[Windows または Mac 上の Excel](#tab/excel-windows)

Windows または Mac 上の Excel でアドインをテストするには、次のコマンドを実行します。 このコマンドを実行すると、ローカル web サーバーが起動し、アドインが読み込まれた状態で Excel が開きます。

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[Excel on the web](#tab/excel-online)

ブラウザー上の Excel でアドインをテストするには、次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します。

```command&nbsp;line
npm run start:web
```

カスタム関数アドインを使用するには、web 上の Excel で新しいブックを開きます。 このブックでは、次の手順を実行して、アドインをサイドロードします。

1. Excel で、[**挿入**] タブを選択し、[**アドイン**] を選択します。

   ![[個人用アドイン] アイコンが強調表示されている web 上の Excel にリボンを挿入する](../images/excel-cf-online-register-add-in-1.png)
   
2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a>あらかじめ用意されているカスタム関数を試す

作成したカスタム関数プロジェクトには、 **/src/functions/functions.js**ファイル内で定義されたあらかじめ用意されたカスタム関数がいくつか含まれています。 **./manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。 Excel でカスタム関数にアクセスするには、CONTOSO 名前空間を使用します。

その後、次の手順を実行し、`ADD` カスタム関数を試します。

1. Excel で、任意のセルに移動し、`=CONTOSO` と入力します。 `CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。

2. セル内で値 `=CONTOSO.ADD(10,200)` を入力して Enter キーを押し、入力パラメーターとして数値 `10` と `200` を指定して、`CONTOSO.ADD` 関数を実行します。

`ADD` カスタム関数によって、指定した 2 つの数字の合計が計算され、**210** という結果が返されます。

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Web からデータを要求するカスタム関数を作成する

Web からデータを統合することは、カスタム関数を使用して Excel を拡張する優れた方法です。 次に、指定された Github `getStarCount`リポジトリのスター数を示すという名前のカスタム関数を作成します。

1. **Starcount**プロジェクトで、 **/src/functions/functions.js**を見つけて、コードエディターで開きます。 

2. **関数 .js**で、次のコードを追加します。 

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

4. Excel でアドインを再登録するには、次の手順を実行します (web、Windows、または Mac の Excel の場合)。 新しい関数を使用できるようにするには、これらの手順を完了する必要があります。

### <a name="excel-on-windows-or-mactabexcel-windows"></a>[Windows または Mac 上の Excel](#tab/excel-windows)

1. Excel を閉じて再び開きます。

2. Excel で [**挿入**] タブを選択し、[**マイ**アドイン] の右側にある下向き矢印を選択します。 ![[個人用アドイン] 矢印が強調表示されている Windows 上の Excel でのリボンの挿入](../images/select-insert.png)

3. 利用可能なアドインの一覧で、[**開発者用アドイン**] セクションを見つけ、 **starcount**アドインを選択して登録します。
    ![[個人用アドイン] ボックスの一覧で強調表示された Excel カスタム関数アドインを使用して、Excel の Excel にリボンを挿入する](../images/list-starcount.png)


# <a name="excel-on-the-webtabexcel-online"></a>[Excel on the web](#tab/excel-online)

1. Excel で、[**挿入**] タブを選択し、[**アドイン**] を選択します。 ![[個人用アドイン] アイコンが強調表示されている web 上の Excel にリボンを挿入する](../images/excel-cf-online-register-add-in-1.png)

2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

---

<ol start="5">
<li> 新しい関数をお試しください。 セル<strong>B1</strong>に、「CONTOSO」というテキストを入力し<strong>ます。GETSTARCOUNT ("OfficeDev", "Excel-ユーザー定義関数")</strong> 。 enter キーを押します。 セル<strong>B1</strong>の結果は、 [Excel のカスタム機能である Github リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions)に与えられている現在の星数であることがわかります。</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>非同期でデータをストリーミングするカスタム関数を作成する

関数`getStarCount`は、特定の時点でリポジトリにある星の数を返します。 カスタム関数は、絶えず変化するデータを返すこともできます。 これらの関数は、ストリーミング関数と呼ばれます。 これらには、 `invocation`関数が呼び出されたセルを参照するパラメーターを含める必要があります。 `invocation`パラメーターは、セルの内容をいつでも更新するために使用されます。  

次のコードサンプルでは、 `currentTime`と`clock`の2つの関数があることに注意してください。 関数`currentTime`は、ストリーミングを使用しない静的関数です。 日付を文字列として返します。 `clock`関数は、 `currentTime`関数を使用して、Excel のセルに対して2秒ごとに新しい時刻を提供します。 を使用`invocation.setResult`して、Excel セルに時刻を提供`invocation.onCanceled`し、関数がキャンセルされたときに発生する処理を処理します。

1. **Starcount**プロジェクトで、次のコードを **/src/functions/functions.js**に追加し、ファイルを保存します。

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

3. Excel でアドインを再登録するには、次の手順を実行します (web、Windows、または Mac の Excel の場合)。 新しい関数を使用できるようにするには、これらの手順を完了する必要があります。 

# <a name="excel-on-windows-or-mactabexcel-windows"></a>[Windows または Mac 上の Excel](#tab/excel-windows)

1. Excel を閉じて再び開きます。

2. Excel で [**挿入**] タブを選択し、[**マイ**アドイン] の右側にある下向き矢印を選択します。 ![[個人用アドイン] 矢印が強調表示されている Windows 上の Excel でのリボンの挿入](../images/select-insert.png)

3. 利用可能なアドインの一覧で、[**開発者用アドイン**] セクションを見つけ、 **starcount**アドインを選択して登録します。
    ![[個人用アドイン] ボックスの一覧で強調表示された Excel カスタム関数アドインを使用して、Excel の Excel にリボンを挿入する](../images/list-starcount.png)

# <a name="excel-on-the-webtabexcel-online"></a>[Excel on the web](#tab/excel-online)

1. Excel で、[**挿入**] タブを選択し、[**アドイン**] を選択します。 ![[個人用アドイン] アイコンが強調表示されている web 上の Excel にリボンを挿入する](../images/excel-cf-online-register-add-in-1.png)

2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

--- 

<ol start="4">
<li>新しい関数をお試しください。 セル<strong>C1</strong>に、「CONTOSO」というテキストを入力し<strong>ます。CLOCK ())</strong>と入力し、enter キーを押します。 現在の日付が表示され、1秒ごとに更新が流れます。 このクロックはループのタイマーにすぎませんが、リアルタイムデータに対する web 要求を行う、より複雑な関数でタイマーを設定するのと同じ概念を使用できます。</li>
</ol>

## <a name="next-steps"></a>次のステップ

おめでとうございます。 新しいカスタム関数プロジェクトを作成し、あらかじめ作成された関数を試し、web からデータを要求するカスタム関数を作成し、データをストリーム処理するカスタム関数を作成しました。 この関数のデバッグは[、カスタム関数のデバッグ手順](../excel/custom-functions-debugging.md)を使用して実行することもできます。 Excel のカスタム関数に関する詳細については、次の記事にお進みください。

> [!div class="nextstepaction"]
> [Excel でカスタム関数を作成する](../excel/custom-functions-overview.md)
