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

 はじめに、Yeoman ジェネレーターを使って、カスタム関数プロジェクトを作成します。 これにより、カスタム関数のコーディングを開始するための正しいフォルダー構造、ソース ファイル、依存関係によるプロジェクトがセットアップされます。

1. 次のコマンドを実行し、以下のようにプロンプトに応答します。

    ```
    yo office
    ```

    * Choose a project type (プロジェクトの種類を選択): `Excel Custom Functions Add-in project (...)`

    * Choose a script type (スクリプトの種類を選択): `JavaScript`

    * What would you want to name your add-in? (アドインの名前を何にしますか) `stock-ticker`

    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/12-10-fork-cf-pic.jpg)

    Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。 プロジェクト ファイルは [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub リポジトリにあります。

2. プロジェクト フォルダーに移動します。

    ```
    cd stock-ticker
    ```

3. このプロジェクトを実行するために必要な自己署名証明書を信頼します。 Windows または Mac についての詳細な手順については、「[自己署名証明書を信頼済みルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」を参照してください。  

4. プロジェクトをビルドします。

    ```
    npm run build
    ```

5. Node.js で実行しているローカル Web サーバーを開始します。

    * Windows 版 Excel を使用してカスタム関数をテストする場合、次のコマンドを実行してローカル Web サーバーを開始し、Excel を起動し、アドインをサイドロードします。

        ```
         npm run start
        ```
        次のコマンドを実行すると、コマンド プロンプトに実行した作業についての詳細が表示され、別の npm ウィンドウが開いてビルドの詳細が表示され、アドインを読み込んだ状態で Excel が起動します。 アドインが読み込まれない場合は、手順 3 が正しく完了しているか確認してください。

    * Excel Online を使用してカスタム関数をテストする場合、次のコマンドを実行してローカル Web サーバーを開始します。

        ```
        npm run start-web
        ```

         次のコマンドを実行すると、別のウィンドウが開いてビルドの詳細が表示されます。 関数を使って、Office Online に新しいブックを開きます。

## <a name="try-out-a-prebuilt-custom-function"></a>あらかじめ用意されているカスタム関数を試す

Yeoman ジェネレーターで作成したカスタム関数プロジェクトには、あらかじめ用意されているカスタム関数がいくつか含まれており、**src/customfunctions.js** ファイル内で定義されています。 プロジェクトのルート ディレクトリの **manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。

Excel の Excel ブックで次の手順を実行し、`ADD` カスタム関数を試してみてください。

1. セル内に **=CONTOSO** と入力します。 `CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。

2. セル内で値 `=CONTOSO.ADD(10,200)` を入力して Enter キーを押し、入力パラメーターとして `10` と `200` を指定して、`CONTOSO.ADD` 関数を実行します。

`ADD` カスタム関数によって、入力パラメーターとして指定した 2 つの数字の合計が計算されます。 「`=CONTOSO.ADD(10,200)`」と入力して Enter キーを押すと、**210** という結果が生成されるはずです。

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Web からデータを要求するカスタム関数を作成する

API に株価を要求し、ワークシートのセルに結果を表示する関数が必要になった場合、どうすればよいでしょうか。 カスタム関数は、Web にデータを非同期で簡単に要求できるように設計されています。

次の手順を実行し、銘柄コード (**MSFT** など) を受け取り、その株価を返す、`stockPrice` という名前のカスタム関数を作成します。 このカスタム関数では、IEX Trading API が使用されます。これは無料であり、認証を必要としません。

1. Yeoman ジェネレーターによって作成された**銘柄コード**プロジェクトで**src/customfunctions.js** ファイルを見つけ、それをコード エディターで開きます。

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

3. In **customfunctions.js**, locate the line`CustomFunctionMappings.INCREMENT = increment;`, add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

4. Excel でこの新しい関数を使用できるようにするには、Excel で関数について説明するメタデータを指定する必要があります。 **config/customfunctions.json** ファイルを開きます。 '関数' 配列に次の JSON オブジェクトを追加し、ファイルを保存します。

    この JSON では、`stockPrice` 関数について説明しています。

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
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

5. 新しい関数をエンドユーザーが使用できるようにするには、Excel にアドインを再登録する必要があります。 このチュートリアルで使用しているプラットフォームの場合、次の手順を実行します。

    * Windows 版 Excel を使用する場合:

        1. Excel を閉じて再び開きます。

        2. Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)

        3. 使用可能なアドインの一覧から **[開発者向けアドイン]** セクションを見つけ、**銘柄コード** アドインを選択して登録します。
            ![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)

    * Excel Online を使用する場合:

        1. Excel Online で [**挿入**] タブを選択し、[**アドイン**] を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)

        2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。 

        3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。 

        4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

6. それでは、新しい関数を試してみましょう。 セル **B1** にテキスト `=CONTOSO.STOCKPRICE("MSFT")` を入力し、Enter キーを押します。 セル **B1** の結果が Microsoft の最新株価になっているはずです。

## <a name="create-a-streaming-asynchronous-custom-function"></a>非同期でデータをストリーミングするカスタム関数を作成する

作成した `stockPrice` 関数では、特定の時点での株価が返されますが、株価は常に変動するものです。 API からデータをストリーミングし、株価をリアルタイム更新するカスタム関数を作成しましょう。

次の手順を実行し、(前の要求が完了しているという条件で) 1,000 ミリ秒ごとに指定の株価を要求する、`stockPriceStream` という名前のカスタム関数を作成します。 最初の要求が進行中のとき、関数が呼び出されているセルに **#GETTING_DATA** というプレースホルダー値が表示されることがあります。 関数によって値が返されると、そのセルの **#GETTING_DATA** がその値で置換られます。

1. Yeoman ジェネレーターによって作成された**銘柄コード** プロジェクトで、次のコードを **src/customfunctions.js** に追加し、ファイルを保存します。

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

2. Excel のユーザーがこの新しい関数を使用できるようにするには、この関数について説明するメタデータを指定します。 Yeoman ジェネレーターによって作成された**銘柄コード** プロジェクトで、**config/customfunctions.json** ファイル内の `functions` 配列に次のオブジェクトを追加し、ファイルを保存します。

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

3. 新しい関数をエンドユーザーが使用できるようにするには、Excel にアドインを再登録する必要があります。 このチュートリアルで使用しているプラットフォームの場合、次の手順を実行します。

    * Windows 版 Excel を使用する場合:

        1. Excel を閉じて再び開きます。
        
        2. Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)

        3. 使用可能なアドインの一覧から **[開発者向けアドイン]** セクションを見つけ、**銘柄コード** アドインを選択して登録します。
            ![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)

    * Excel Online を使用する場合:

        1. Excel Online で [**挿入**] タブを選択し、[**アドイン**] を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)

        2. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

        3. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

        4. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

4. それでは、新しい関数を試してみましょう。 セル **C1** にテキスト `=CONTOSO.STOCKPRICESTREAM("MSFT")` を入力し、Enter キーを押します。 株式市場が開いている場合、セル **C1** の結果が継続的に更新され、Microsoft の株価がリアルタイムで反映されます。

## <a name="next-steps"></a>次の手順

このチュートリアルでは、新しいカスタム関数プロジェクトを作成し、あらかじめ用意されている関数を試し、Web にデータを要求するカスタム関数を作成し、Web からデータをリアルタイムでストリーミングするカスタム関数を作成しました。 Excel のカスタム関数に関する詳細については、次の記事にお進みください。

> [!div class="nextstepaction"]
> [Excel でカスタム関数を作成する](../excel/custom-functions-overview.md)

## <a name="legal-information"></a>法的情報

データは [IEX](https://iextrading.com/developer/) より無料提供されました。 [IEX の利用規約](https://iextrading.com/api-exhibit-a/)をご覧ください。 Microsoft はこのチュートリアルで IEX API を教育目的でのみ使用しています。
