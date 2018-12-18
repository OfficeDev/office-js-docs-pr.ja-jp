# <a name="tutorial-create-custom-functions-in-excel"></a>チュートリアル: Excel でのカスタム関数の作成

## <a name="introduction"></a>概要

カスタム関数では、関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。 ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。 ユーザー定義の計算のような単純なタスク、または Web からワークシートへのデータのリアルタイム ストリーミングのようなより複雑なタスクを実行するカスタム関数を作成できます。

このチュートリアルの内容:
> [!div class="checklist"]
> * Yo Office ジェネレーターを使用してカスタム関数プロジェクトを作成する
> * あらかじめ用意されているカスタム関数を使用し、単純な計算を実行する
> * Web からデータを要求するカスタム関数を作成する
> * Web からデータをリアルタイムでストリーミングするカスタム関数を作成する

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a>前提条件

* [Node.js と npm](https://nodejs.org/en/)

* [Git バッシュ](https://git-scm.com/downloads) (または別の Git クライアント)

* [Yeoman](https://yeoman.io/) と [Yo Office ジェネレーター](https://www.npmjs.com/package/generator-office)の最新版。 以上のツールをグローバルにインストールするには、コマンド プロンプトから次のコマンドを実行します。

    ```bash
    npm install -g yo generator-office
    ```

* Windows 版 Excel (バージョン 1810 以降) または Excel Online

* [Office Insider プログラム](https://products.office.com/office-insider)に加入する (**Insider** レベル -- 以前は "Insider Fast" と呼ばれていたもの)

## <a name="create-a-custom-functions-project"></a>カスタム関数プロジェクトを作成する

このチュートリアルでは最初に、Yo Office ジェネレーターを使用し、カスタム関数プロジェクトに必要なファイルを作成します。

1. 次のコマンドを実行し、以下のようにプロンプトに応答します。

    ```bash
    yo office
    ```

    * Choose a project type (プロジェクトの種類を選択): `Excel Custom Functions Add-in project (...)`
    * Choose a script type (スクリプトの種類を選択): `JavaScript`
    * What would you want to name your add-in? (アドインの名前を何にしますか) `stock-ticker`

    ![カスタム関数の Yo Office バッシュ プロンプト](../images/yo-office-cfs-stock-ticker-3.png)

    ウィザードを完了すると、ジェネレーターによってプロジェクト ファイルが作成され、サポート ノード コンポーネントがインストールされます。 プロジェクト ファイルは [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub リポジトリにあります。

2. プロジェクト フォルダーに移動します。

    ```bash
    cd stock-ticker
    ```

3. ローカル Web サーバーを開始します。

    * Windows 版 Excel を使用してカスタム関数をテストする場合、次のコマンドを実行してローカル Web サーバーを開始し、Excel を起動し、アドインをサイドロードします。

        ```bash
        npm run start-desktop
        ```

    * Excel Online を使用してカスタム関数をテストする場合、次のコマンドを実行してローカル Web サーバーを開始します。 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a>あらかじめ用意されているカスタム関数をテストする

Yo Office ジェネレーターで作成したカスタム関数プロジェクトには、あらかじめ用意されているカスタム関数がいくつか含まれており、**src/functions/functions.js** ファイル内で定義されています。 プロジェクトのルート ディレクトリの **manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。

あらかじめ用意されているカスタム関数を使用する前に、Excel でカスタム関数アドインを登録する必要があります。 そのためには、このチュートリアルで使用しているプラットフォームの場合、次の手順を実行します。

* Windows 版 Excel を使用してカスタム関数をテストする場合:

    1. Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)

    2. 使用可能なアドインの一覧から [**開発者向けアドイン**] を見つけ、[**Excel カスタム関数**] アドインを選択して登録します。
        ![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)

* Excel Online を使用してカスタム関数をテストする場合: 

    1. Excel Online で [**挿入**] タブを選択し、[**アドイン**] を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)

    2. [**マイ アドインの管理**] を選択し、[**マイ アドインのアップロード**] を選択します。 

    3. [**参照...**] を選択し、Yo Office ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。 

    4. ファイル **manifest.xml** を選択し、[**開く**] を選択し、[**アップロード**] を選択します。

この時点で、プロジェクトにあらかじめ用意されているカスタム関数が読み込まれており、Excel 内で使用できます。 Excel で次の手順を実行し、`ADD` カスタム関数を試してみてください。

1. セル内に「**=CONTOSO**」と入力します。 `CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。

2. セル内で値 `=CONTOSO.ADD(10,200)` を入力して Enter キーを押し、入力パラメーターとして `10` と `200` を指定して、`CONTOSO.ADD` 関数を実行します。

`ADD` カスタム関数によって、入力パラメーターとして指定した 2 つの数字の合計が計算されます。 「`=CONTOSO.ADD(10,200)`」と入力して Enter キーを押すと、**210** という結果が生成されるはずです。

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Web からデータを要求するカスタム関数を作成する

API に株価を要求し、ワークシートのセルに結果を表示する関数が必要になった場合、どうすればよいでしょうか。 カスタム関数は、Web にデータを非同期で簡単に要求できるように設計されています。

次の手順を実行し、銘柄コード (**MSFT** など) を受け取り、その株価を返す、`stockPrice` という名前のカスタム関数を作成します。 このカスタム関数では、IEX Trading API が使用されます。これは無料であり、認証を必要としません。

1. Yo Office ジェネレーターによって作成された**株価情報** プロジェクトで、ファイル **src/functions/functions.js** を見つけ、それをコード エディターで開きます。

2. 次のコードを **customfunctions.js** に追加し、ファイルを保存します。

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

    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

3. Excel のエンドユーザーがこの新しい関数を使用できるようにするには、この関数について説明するメタデータを指定する必要があります。 Yo Office ジェネレーターによって作成された**株価情報** プロジェクトで、ファイル **src/functions/functions.json** を見つけ、それをコード エディターで開きます。 **src/functions/functions.json** ファイル内の `functions` 配列に次のオブジェクトを追加し、ファイルを保存します。

    この JSON では、`stockPrice` 関数について説明しています。

    ```json
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
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

4. 新しい関数をエンドユーザーが使用できるようにするには、Excel にアドインを登録する必要があります。 このチュートリアルで使用しているプラットフォームの場合、次の手順を実行します。

    * Windows 版 Excel を使用する場合:

        1. Excel を閉じて再び開きます。

        2. Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)

        1. 使用可能なアドインの一覧から [**開発者向けアドイン**] を見つけ、[**Excel カスタム関数**] アドインを選択して登録します。
            ![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)

    * Excel Online を使用する場合: 

        1. Excel Online で [**挿入**] タブを選択し、[**アドイン**] を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)

        2. [**マイ アドインの管理**] を選択し、[**マイ アドインのアップロード**] を選択します。 

        3. [**参照...**] を選択し、Yo Office ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。 

        4. ファイル **manifest.xml** を選択し、[**開く**] を選択し、[**アップロード**] を選択します。

5. それでは、新しい関数を試してみましょう。 セル **B1** にテキスト `=CONTOSO.STOCKPRICE("MSFT")` を入力し、Enter キーを押します。 セル **B1** の結果が Microsoft の最新株価になっているはずです。

## <a name="create-a-streaming-asynchronous-custom-function"></a>非同期でデータをストリーミングするカスタム関数を作成する

作成した `stockPrice` 関数では、特定の時点での株価が返されますが、株価は常に変動するものです。 API からデータをストリーミングし、株価をリアルタイム更新するカスタム関数を作成しましょう。

次の手順を実行し、(前の要求が完了しているという条件で) 1,000 ミリ秒ごとに指定の株価を要求する、`stockPriceStream` という名前のカスタム関数を作成します。 最初の要求が進行中のとき、関数が呼び出されているセルに **#GETTING_DATA** というプレースホルダー値が表示されることがあります。 関数によって値が返されると、そのセルの **#GETTING_DATA** がその値で置換されます。

1. Yo Office ジェネレーターによって作成された**株価情報** プロジェクトで、次のコードを **src/functions/functions.js** に追加し、ファイルを保存します。

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

    CustomFunctionMappings.STOCKPRICESTREAM = stockPriceStream;
    ```

2. Excel のエンドユーザーがこの新しい関数を使用できるようにするには、この関数について説明するメタデータを指定する必要があります。 Yo Office ジェネレーターによって作成された**株価情報** プロジェクトで、**src/functions/functions.json** ファイル内の `functions` 配列に次のオブジェクトを追加し、ファイルを保存します。

    この JSON では、`stockPriceStream` 関数について説明しています。 ストリーミング関数の場合、このコード サンプルで示すように、`options` オブジェクト内で `stream` プロパティと `cancelable` プロパティを `true` に設定する必要があります。

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
                "description": "stock ticker name",
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

3. 新しい関数をエンドユーザーが使用できるようにするには、Excel にアドインを登録する必要があります。 このチュートリアルで使用しているプラットフォームの場合、次の手順を実行します。

    * Windows 版 Excel を使用する場合:

        1. Excel を閉じて再び開きます。
        
        2. Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)

        3. 使用可能なアドインの一覧から [**開発者向けアドイン**] を見つけ、[**Excel カスタム関数**] アドインを選択して登録します。
            ![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)

    * Excel Online を使用する場合: 

        1. Excel Online で [**挿入**] タブを選択し、[**アドイン**] を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)

        2. [**マイ アドインの管理**] を選択し、[**マイ アドインのアップロード**] を選択します。 

        3. [**参照...**] を選択し、Yo Office ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。 

        4. ファイル **manifest.xml** を選択し、[**開く**] を選択し、[**アップロード**] を選択します。

4. それでは、新しい関数を試してみましょう。 セル **C1** にテキスト `=CONTOSO.STOCKPRICESTREAM("MSFT")` を入力し、Enter キーを押します。 株式市場が開いている場合、セル **C1** の結果が継続的に更新され、Microsoft の株価がリアルタイムで反映されます。

## <a name="next-steps"></a>次の手順

このチュートリアルでは、新しいカスタム関数プロジェクトを作成し、あらかじめ用意されている関数を試し、Web にデータを要求するカスタム関数を作成し、Web からデータをリアルタイムでストリーミングするカスタム関数を作成しました。 Excel のカスタム関数に関する詳細については、次の記事にお進みください。 

> [!div class="nextstepaction"]
> [Excel でカスタム関数を作成する](../excel/custom-functions-overview.md)

## <a name="legal-information"></a>法的情報

データは [IEX](https://iextrading.com/developer/) より無料提供されました。 [IEX の利用規約](https://iextrading.com/api-exhibit-a/)をご覧ください。 Microsoft はこのチュートリアルで IEX API を教育目的でのみ使用しています。
