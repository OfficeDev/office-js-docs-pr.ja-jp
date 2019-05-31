---
title: Excel カスタム関数のチュートリアル
description: このチュートリアルでは、計算の実行、Web データの要求、Web データのストリームが可能なカスタム関数を含む Excel アドインを作成します。
ms.date: 05/16/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 7d4d87a6bb3910c1b46698d5a2ff211ea1bbc6dd
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589175"
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

* Excel on Windows (64 ビットバージョン1810以降) または Excel Online

## <a name="create-a-custom-functions-project"></a>カスタム関数プロジェクトを作成する

 まず、カスタム関数アドインをビルドするコード プロジェクトを作成します。 [Office アドイン用の [ごみ箱] ジェネレーター](https://www.npmjs.com/package/generator-office)では、プロジェクトに事前に用意されているカスタム関数を使用してセットアップし、試すことができます。カスタム関数のクイックスタートを既に実行してプロジェクトを生成した場合は、そのプロジェクトを引き続き使用して、[この手順](#create-a-custom-function-that-requests-data-from-the-web)に進んでください。

1. 次のコマンドを実行し、以下のようにプロンプトに応答します。
    
    ```command&nbsp;line
    yo office
    ```
    
    * **Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project`
    * **Choose a script type: (スクリプトの種類を選択)** `JavaScript`
    * **What would you want to name your add-in?: (アドインの名前を何にしますか)** `stock-ticker`

    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/UpdatedYoOfficePrompt.png)
    
    Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。

2. プロジェクトのルート フォルダーに移動します。
    
    ```command&nbsp;line
    cd stock-ticker
    ```

3. プロジェクトをビルドします。
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > 開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。 `npm run build`の実行後に証明書をインストールするように指示が出された場合は、Yeomanジェネレーターが提供する証明書をインストールする手順に従ってください。

4. Node.js で実行しているローカル Web サーバーを開始します。 カスタム関数アドインは、Windows または Excel Online で Excel で試すことができます。

# <a name="excel-on-windowstabexcel-windows"></a>[Windows 上の Excel](#tab/excel-windows)

Windows の Excel でアドインをテストするには、次のコマンドを実行します。 このコマンドを実行すると、ローカル web サーバーが起動し、アドインが読み込まれた状態で Excel が開きます。

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

Excel Online でアドインをテストするには、次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します。

```command&nbsp;line
npm run start:web
```

カスタム関数アドインを使用するには、Excel Online で新しいブックを開きます。 このブックでは、次の手順を実行して、アドインをサイドロードします。

1. Excel Online で、**[挿入]** タブを選択して、**[アドイン]** を選択します。

   ![[個人用アドイン] アイコンが強調表示された状態で Excel Online にリボンを挿入する](../images/excel-cf-online-register-add-in-1.png)
   
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

Web からデータを統合することは、カスタム関数を使用して Excel を拡張する優れた方法です。 次に、Web API から株価情報を取得し、ワークシートのセルに結果を返す、`stockPrice` というカスタム関数を作成します。 IEX Trading API を使用します。これは無料であり、認証を必要としません。

1. **銘柄**コードプロジェクトで、 **/src/functions/functions.js**を見つけて、コードエディターで開きます。

2. **Js**で、 `increment`関数を見つけて、その関数の後に次のコードを追加します。

    ```js
    /**
    * Fetches current stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @returns {number} The current stock price.
    */
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }
    CustomFunctions.associate("STOCKPRICE", stockPrice);
    ```

    `CustomFunctions.associate` コードは、JavaScript で関数の `id` と `stockPrice` の関数アドレスを関連付けて、Excel により関数を呼び出せるようにします。

3. 次のコマンドを実行してプロジェクトを再構築します。

    ```command&nbsp;line
    npm run build
    ```

4. 次の手順を実行して (Excel on Windows または Excel Online の場合)、Excel でアドインを再登録します。 新しい関数を使用できるようにするには、これらの手順を完了する必要があります。 

# <a name="excel-on-windowstabexcel-windows"></a>[Windows 上の Excel](#tab/excel-windows)

1. Excel を閉じて再び開きます。

2. Excel で [**挿入**] タブを選択し、[**マイ**アドイン] の右側にある下向き矢印を選択します。 ![[個人用アドイン] 矢印が強調表示されている Windows 上の Excel でのリボンの挿入](../images/select-insert.png)

3. 使用可能なアドインの一覧から **[開発者向けアドイン]** セクションを見つけ、**銘柄コード** アドインを選択して登録します。
    ![[個人用アドイン] ボックスの一覧で強調表示された Excel カスタム関数アドインを使用して、Excel の Excel にリボンを挿入する](../images/list-stock-ticker-red.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. Excel Online で **[挿入]** タブを選択し、**[アドイン]** を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)

2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。 

3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。 

4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

---

<ol start="5">
<li> 新しい関数をお試しください。 セル <strong>B1</strong> に <strong>=CONTOSO.STOCKPRICE("MSFT")</strong> と入力し、Enter キーを押します。 セル <strong>B1</strong> の結果が Microsoft の最新株価になっているはずです。</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>非同期でデータをストリーミングするカスタム関数を作成する

`stockPrice` 関数では、特定の時点での株価が返されますが、株価は常に変動するものです。 次に、1000 ミリ秒ごと株価を取得する、`stockPriceStream` という名前のカスタム関数を作成します。

1. **銘柄**コードプロジェクトで、次のコードを **/src/functions/functions.js**に追加し、ファイルを保存します。

    ```js
    /**
    * Streams real time stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @param {CustomFunctions.StreamingInvocation<number>} invocation
    */
    function stockPriceStream(ticker, invocation) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    invocation.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    invocation.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        invocation.onCanceled = () => {
            clearInterval(timer);
        };
    }
    CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
    ```
    
    `CustomFunctions.associate` コードは、JavaScript で関数の `id` と `stockPriceStream` の関数アドレスを関連付けて、Excel により関数を呼び出せるようにします。
    
2. 次のコマンドを実行してプロジェクトを再構築します。

    ```command&nbsp;line
    npm run build
    ```

3. 次の手順を実行して (Excel on Windows または Excel Online の場合)、Excel でアドインを再登録します。 新しい関数を使用できるようにするには、これらの手順を完了する必要があります。 

# <a name="excel-on-windowstabexcel-windows"></a>[Windows 上の Excel](#tab/excel-windows)

1. Excel を閉じて再び開きます。

2. Excel で [**挿入**] タブを選択し、[**マイ**アドイン] の右側にある下向き矢印を選択します。 ![[個人用アドイン] 矢印が強調表示されている Windows 上の Excel でのリボンの挿入](../images/select-insert.png)

3. 使用可能なアドインの一覧から **[開発者向けアドイン]** セクションを見つけ、**銘柄コード** アドインを選択して登録します。
    ![[個人用アドイン] ボックスの一覧で強調表示された Excel カスタム関数アドインを使用して、Excel の Excel にリボンを挿入する](../images/list-stock-ticker-red.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. Excel Online で **[挿入]** タブを選択し、**[アドイン]** を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)

2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

--- 

<ol start="4">
<li>新しい関数をお試しください。 セル <strong>C1</strong> に <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong> と入力し、Enter キーを押します。 株式市場が開いている場合、セル <strong>C1</strong> の結果が継続的に更新され、Microsoft の株価がリアルタイムで反映されます。</li>
</ol>

## <a name="next-steps"></a>次の手順

おめでとうございます。 新しいカスタム関数プロジェクトを作成し、あらかじめ用意されている関数を試し、Web にデータを要求するカスタム関数を作成し、Web からデータをリアルタイムでストリーミングするカスタム関数を作成しました。 この関数のデバッグは[、カスタム関数のデバッグ手順](../excel/custom-functions-debugging.md)を使用して実行することもできます。 Excel のカスタム関数に関する詳細については、次の記事にお進みください。

> [!div class="nextstepaction"]
> [Excel でカスタム関数を作成する](../excel/custom-functions-overview.md)

### <a name="legal-information"></a>法的情報

データは [IEX](https://iextrading.com/developer/) より無料提供されました。 [IEX の利用規約](https://iextrading.com/api-exhibit-a/)をご覧ください。 Microsoft はこのチュートリアルで IEX API を教育目的でのみ使用しています。
