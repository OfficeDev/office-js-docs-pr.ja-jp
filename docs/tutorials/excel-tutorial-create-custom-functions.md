---
title: Excel カスタム関数のチュートリアル (プレビュー)
description: このチュートリアルでは、計算の実行、Web データの要求、Web データのストリームが可能なカスタム関数を含む Excel アドインを作成します。
ms.date: 01/08/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 4ac735e6fc19f13859d07df6cb3d2443e6dfe2fd
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982021"
---
# <a name="tutorial-create-custom-functions-in-excel-preview"></a>チュートリアル: Excel でのカスタム関数の作成 (プレビュー)

カスタム関数では、関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。 ユーザーは、Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。 計算のような単純なタスク、または Web からワークシートへのデータのリアルタイム ストリーミングのようなより複雑なタスクを実行するカスタム関数を作成できます。

このチュートリアルの内容:
> [!div class="checklist"]
> * [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用して、カスタム関数アドインを作成します。 
> * あらかじめ用意されているカスタム関数を使用し、単純な計算を実行します。
> * Web からデータを取得するカスタム関数を作成します。
> * Web からデータをリアルタイムでストリーミングするカスタム関数を作成します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a>前提条件

* [Node.js](https://nodejs.org/en/) (バージョン 8.0.0 以降)

* [Git バッシュ](https://git-scm.com/downloads) (または別の Git クライアント)

* 最新バージョンの [Yeoman](https://yeoman.io/) と [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)。これらのツールをグローバルにインストールするには、コマンド プロンプトから次のコマンドを実行します。

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > 以前に Yeoman ジェネレーターをインストールしている場合でも、npm からパッケージを最新バージョンに更新することをお勧めします。

* Windows 版 Excel (64 ビット バージョン 1810 以降) または Excel Online

* [Office Insider プログラム](https://products.office.com/office-insider)に加入する (**Insider** レベル -- 以前は "Insider Fast" と呼ばれていたもの)

## <a name="create-a-custom-functions-project"></a>カスタム関数プロジェクトを作成する

 まず、カスタム関数アドインをビルドするコード プロジェクトを作成します。 [Yeoman Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用すると、プロジェクトをセットアップして、いくつかの初期カスタム関数を試すことができます。

1. 次のコマンドを実行し、以下のようにプロンプトに応答します。
    
    ```
    yo office
    ```
    
    * Choose a project type (プロジェクトの種類を選択): `Excel Custom Functions Add-in project (...)`
    * Choose a script type (スクリプトの種類を選択): `JavaScript`
    * What would you want to name your add-in? (アドインの名前を何にしますか) `stock-ticker`
    
    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/12-10-fork-cf-pic.jpg)
    
    Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしている Node.js コンポーネントをインストールします。

2. プロジェクト フォルダーに移動します。
    
    ```
    cd stock-ticker
    ```

3. このプロジェクトを実行するために必要な自己署名証明書を信頼します。 Windows または Mac についての詳細な手順については、「[自己署名証明書を信頼済みルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」を参照してください。  

4. プロジェクトをビルドします。
    
    ```
    npm run build
    ```

5. Node.js で実行しているローカル Web サーバーを開始します。 Windows 用 Excel または Excel Online で、カスタム関数アドインを試すことができます。

# <a name="excel-for-windowstabexcel-windows"></a>[Windows 用 Excel](#tab/excel-windows)

次のコマンドを実行します。

```
npm run start
```

このコマンドは、Web サーバーを開始し、カスタム関数アドインを Windows 用 Excel にサイドロードします。

> [!NOTE]
> アドインが読み込まれない場合は、手順 3 が正しく完了しているか確認してください。 **[実行時のログ](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** を、任意のインストールまたは実行時の問題と同様に、アドインの XML マニフェスト ファイルの問題をトラブルシューティングすることもできます。 実行時のログの書き込み`console.log`ステートメントを検索して、問題を解決するためにログ ファイルにします。

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

次のコマンドを実行します。

```
npm run start-web
```

このコマンドは、Web サーバーを開始します。 アドインをサイドロードするには、次の手順を実行します。

<ol type="a">
   <li>Excel Online で、<strong>[挿入]</strong> タブを選択して、<strong>[アドイン]</strong> を選択します。<br/>
   <img src="../images/excel-cf-online-register-add-in-1.png" alt="Insert ribbon in Excel Online with the My Add-ins icon highlighted"></li>
   <li><strong>[マイ アドインの管理]</strong> を選択し、<strong>[マイ アドインのアップロード]</strong> を選択します。</li> 
   <li><strong>[参照...]</strong> を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</li> 
   <li><strong>manifest.xml</strong> ファイルを選択し、<strong>[開く]</strong> を選択し、<strong>[アップロード]</strong> を選択します。</li>
</ol>

> [!NOTE]
> アドインが読み込まれない場合は、手順 3 が正しく完了しているか確認してください。

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a>あらかじめ用意されているカスタム関数を試す

既に作成したカスタム関数のプロジェクトには、ADD と INCREMENT という名前のあらかじめ用意されている 2 つのカスタム機能があります。 これらのあらかじめ用意されている関数のコードは、**src/customfunctions.js** ファイルにあります。 **./manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。 Excel でカスタム関数にアクセスするには、CONTOSO 名前空間を使用します。

その後、次の手順を実行し、`ADD` カスタム関数を試します。

1. Excel で、任意のセルに移動し、`=CONTOSO` と入力します。 `CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。

2. セル内で値 `=CONTOSO.ADD(10,200)` を入力して Enter キーを押し、入力パラメーターとして数値 `10` と `200` を指定して、`CONTOSO.ADD` 関数を実行します。

`ADD` カスタム関数によって、指定した 2 つの数字の合計が計算され、**210** という結果が返されます。

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Web からデータを要求するカスタム関数を作成する

Web からデータを統合することは、カスタム関数を使用して Excel を拡張する優れた方法です。 次に、Web API から株価情報を取得し、ワークシートのセルに結果を返す、`stockPrice` というカスタム関数を作成します。 IEX Trading API を使用します。これは無料であり、認証を必要としません。

1. **銘柄コード**プロジェクトで **src/customfunctions.js** ファイルを見つけ、それをコード エディターで開きます。

2. **customfunctions.js** で、`increment` 関数を見つけ、その関数の直後に次のコードを追加します。

    ```js
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

> [!NOTE]
> In the January Insiders 1901 Build, there is a bug preventing fetch calls from executing which will result in #VALUE!.
> To workaround this please use the [XMLHTTPRequest API](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime#requesting-external-data) to make the web request.

3. In **customfunctions.js**, locate the line `CustomFunctions.associate("INCREMENT", increment);`. Add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctions.associate("STOCKPRICE", stockprice);
    ```

    `CustomFunctions.associate` コードは、JavaScript で関数の `id` と `increment` の関数アドレスを関連付けて、Excel により関数を呼び出せるようにします。

    Excel でカスタム関数を使用できるようにするには、その前にメタデータを使用してそれを記述する必要があります。 以前に `associate` メソッドで使用した `id` を、他のいくつかのメタデータと共に定義する必要があります。


4. **config/customfunctions.json** ファイルを開きます。 '関数' 配列に次の JSON オブジェクトを追加し、ファイルを保存します。

    ```JSON
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

    この JSON は、`stockPrice` 関数、そのパラメーター、それによって返される結果の種類を記述します。

5. 新しい関数を使用できるようにするには、Excel でアドインを再登録します。 

# <a name="excel-for-windowstabexcel-windows"></a>[Windows 用 Excel](#tab/excel-windows)

1. Excel を閉じて再び開きます。

2. Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)

3. 使用可能なアドインの一覧から **[開発者向けアドイン]** セクションを見つけ、**銘柄コード** アドインを選択して登録します。
    ![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 用 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. Excel Online で **[挿入]** タブを選択し、**[アドイン]** を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)

2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。 

3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。 

4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

--- 

<ol start="6">
<li> 新しい関数をお試しください。 セル <strong>B1</strong> に <strong>=CONTOSO.STOCKPRICE("MSFT")</strong> と入力し、Enter キーを押します。 セル <strong>B1</strong> の結果が Microsoft の最新株価になっているはずです。</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>非同期でデータをストリーミングするカスタム関数を作成する

`stockPrice` 関数では、特定の時点での株価が返されますが、株価は常に変動するものです。 次に、1000 ミリ秒ごと株価を取得する、`stockPriceStream` という名前のカスタム関数を作成します。

1. **銘柄コード**プロジェクトで、次のコードを **src/customfunctions.js** に追加し、ファイルを保存します。

    ```js
    function stockPriceStream(ticker, handler) {
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
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }
    
    CustomFunctions.associate("STOCKPRICESTREAM", stockpricestream);
    ```
    
    Excel でカスタム関数を使用できるようにするには、その前にメタデータを使用してそれを記述する必要があります。
    
2. **銘柄コード**プロジェクトで、**config/customfunctions.json** ファイル内の `functions` 配列に次のオブジェクトを追加し、ファイルを保存します。
    
    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

    この JSON は、`stockPriceStream` 関数を記述します。 ストリーミング関数の場合、このコード サンプルで示すように、`options` オブジェクト内で `stream` プロパティと `cancelable` プロパティを `true` に設定する必要があります。

3. 新しい関数を使用できるようにするには、Excel でアドインを再登録します。

# <a name="excel-for-windowstabexcel-windows"></a>[Windows 用 Excel](#tab/excel-windows)

1. Excel を閉じて再び開きます。

2. Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)

3. 使用可能なアドインの一覧から **[開発者向けアドイン]** セクションを見つけ、**銘柄コード** アドインを選択して登録します。
    ![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 用 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)

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

おめでとうございます。 新しいカスタム関数プロジェクトを作成し、あらかじめ用意されている関数を試し、Web にデータを要求するカスタム関数を作成し、Web からデータをリアルタイムでストリーミングするカスタム関数を作成しました。 Excel のカスタム関数に関する詳細については、次の記事にお進みください。

> [!div class="nextstepaction"]
> [Excel でカスタム関数を作成する](../excel/custom-functions-overview.md)

### <a name="legal-information"></a>法的情報

データは [IEX](https://iextrading.com/developer/) より無料提供されました。 [IEX の利用規約](https://iextrading.com/api-exhibit-a/)をご覧ください。 Microsoft はこのチュートリアルで IEX API を教育目的でのみ使用しています。


