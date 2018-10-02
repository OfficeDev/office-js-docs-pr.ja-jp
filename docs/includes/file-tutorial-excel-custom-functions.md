# <a name="tutorial-create-custom-functions-in-excel"></a>チュートリアル: Excel でカスタム関数を作成します。

## <a name="introduction"></a>概要

カスタム関数を使用すると、JavaScriptでこれらの関数をアドインの一部として定義することにより、Excelに新しい関数を追加できます。 Excel内のユーザーは、Excel の他のネイティブ関数（`SUM()` など）と同様に、カスタム関数にアクセスできます。 ユーザー設定の計算などの単純なタスク、またはWeb からワークシートにリアルタイムデータをストリーミングするなど、より複雑なタスクを実行するカスタム関数を作成することができます。

このチュートリアルでは、以下の操作を実行します。
> [!div class="checklist"]
> * Yo Office ジェネレーターを使用してカスタム関数プロジェクトを作成します。
> * 作成済みのカスタム関数を使用して、単純な計算を実行するには
> * Web サイトからデータを要求するカスタム関数を作成します。
> * Web サイトからのリアルタイムのデータをストリームするカスタム関数を作成します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a>前提条件

* [Node および npm](https://nodejs.org/en/)

* [Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)

*  [Yeoman](http://yeoman.io/) および [ Yo Office  ジェネレーター](https://www.npmjs.com/package/generator-office)の最新バージョンです。 グローバルにこれらのツールをインストールするには、コマンド プロンプトを使用して次のコマンドを実行します。

    ```bash
    npm install -g yo generator-office
    ```

* Windows 版 Excel  (ビルド 10827 またはそれ以降) またはExcel Online

* [Office 内部からプログラムに参加します。](https://products.office.com/office-insider)

## <a name="create-a-custom-functions-project"></a>カスタム関数プロジェクトを作成します。

Yo Office ジェネレーターを使用するカスタム関数プロジェクトに必要なファイルを作成するこのチュートリアルを開始するでしょう。

1. 次のコマンドを実行し、以下のプロンプトに応答します。

    ```bash
    yo office
    ```

    * プロジェクト タイプを選択してください `Excel Custom Functions Add-in project (...)`
    * スクリプト タイプを選択してください `JavaScript`
    * アドインの名前を何にしますか？ `stock-ticker`

    ![Yo Office bashは、カスタム関数のプロンプトを表示します。](../images/yo-office-cfs-stock-ticker-3.png)

    ウィザードを完了すると、ジェネレーターがプロジェクト ファイルを作成し、ノードのサポート コンポーネントをインストールします。

2. プロジェクト フォルダーに移動します。

    ```bash
    cd stock-ticker
    ```

3. ローカル web サーバーを起動します。

    * Windows 版 Excel のテストに使用する場合、ローカルの web サーバーを起動するのには次のコマンドを実行、カスタム関数は Excel、およびアドインが sideload を起動します。

        ```bash
        npm start
        ```

    * Excel Online を使用してカスタム関数をテストする場合は、次のコマンドを実行してローカルWebサーバーを起動します。 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a>作成済みのカスタム関数を試してみてください。

Yo Office ジェネレーターを使用して作成したカスタム関数プロジェクトには、 **src/customfunction.js** ファイル内で定義されたいくつか作成済みのカスタム関数が含まれています。 プロジェクトのルートディレクトリにある **manifest.xml** ファイルは、すべてのカスタム関数が`CONTOSO` 名前空間に属することを指定します。

作成済みのカスタム関数のいずれかを使用する前に、Excelでカスタム関数アドインを登録する必要があります。 このチュートリアルで使用するプラットフォームの手順を完了するようにします。

* カスタム関数をテストするには 、Windows 版 Excel を使用します。


    1. Excel では、 **[挿入]** タブを選択し、 **[アドイン]** の右にある下向き矢印を選択します。![ [個人用アドイン]の矢印が強調表示された状態で windows 版 Excel にリボンを挿入します。
       ](../images/excel-cf-register-add-in-1b.png)

    2. 利用可能なアドインの一覧で、 **開発者アドイン** のセクションを検索し、 **Excelカスタム関数** アドインを選択して登録します。
        ![Excel カスタム関数アドインを[個人用アドイン] リストで強調表示して、windows 版 Excel にリボンを挿入します。](../images/excel-cf-register-add-in-2.png)

* カスタム関数をテストするには 、Excel Online を使用します。 

    1. Excel Online で** [挿入]** ]タブを選択し、** [アドイン]** を選択します。![  [個人用アドイン]アイコンを強調表示して Excel Online でリボンを挿入します。](../images/excel-cf-online-register-add-in-1.png)

    2.  ** [個人用アドインの管理] ** を選択し、 ** [個人用アドインのアップロード] **を選択します。 

    3.   ** [参照... ]**  を選択し、Yo Officeジェネレータが作成したプロジェクトのルートディレクトリに移動します。 

    4.  ** Manifest.xml** ファイルを選択して ** 開く** を選択し、** アップロード** を選択します。

この時点で、プロジェクトの作成済みのユーザー定義関数がロードされて Excel 内で使用できます。 Excelで次の手順を実行して、`ADD` カスタム関数を試してみてください。

1. セル内には、 **=CONTOSO**を入力します。 オートコンプリートメニューには、 `CONTOSO` 名前空間内のすべての関数の一覧が表示されます。

2. セルに次の値を指定し、enter キーを押して、`CONTOSO.ADD` 数値 `10` および`200` 入力パラメータとして関数を実行します。

    ```
    =CONTOSO.ADD(10,200)
    ```

 `ADD` カスタム関数は、入力パラメーターとして指定されている 2 つの数値の合計を計算します。 `=CONTOSO.ADD(10,200)` を入力すると、Enterキーを押した後にセル内に**210** という結果が表示されます。


## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Web サイトからデータを要求するカスタム関数を作成します。

APIからの在庫の価格を要求し、その結果をワークシートのセルに表示する機能が必要な場合はどうなりますか？ カスタム関数は、Webサイトから非同期にデータを簡単に要求できるように設計されています。

`stockPrice` というカスタム関数を作成し、株価表示（例：**MSFT** ）を受け取り、その株式の価格を返す次の手順を実行します。 このカスタム関数は、無料で認証を必要としない IEX 取引 APIを使用します。

1. Yo Office ジェネレーターが作成した **株価表示** プロジェクトでは、ファイル **src/customfunctions.js** を検索し、コード エディターで開きます。

2. 次のコードを **customfunctions.js** に追加して、ファイルを保存します。

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

3. Excelでエンドユーザーがこの新しい機能を利用できるようにする前に、この機能を説明するメタデータを指定する必要があります。 Yo Office ジェネレーターが作成した **株価表示** プロジェクトでは、ファイル **config/customfunctions.json** を検索し、コード エディターで開きます。 次のオブジェクトを **config/customfunctions.json**  ファイル内の   `functions` 配列に追加し、ファイルを保存します。

    このJSONは、 `stockPrice` 関数を説明します。

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

4. エンドユーザーが新しい機能を利用できるようにするには、アドインをExcelに再登録する必要があります。 このチュートリアルで使用しているプラットフォームの次の手順を実行します。

    * Windows 版 Excel の場合

        1. Excelを終了し、再度Excelを開きます。

        2. Excel では、 **[挿入]** タブを選択し、 **[アドイン]** の右にある下向き矢印を選択します。![ [個人用アドイン]の矢印が強調表示された状態で windows 版 Excel にリボンを挿入します。
           ](../images/excel-cf-register-add-in-1b.png)

        1. 利用可能なアドインの一覧で、 **開発者アドイン** のセクションを検索し、 **Excelカスタム関数** アドインを選択して登録します。
            ![Excel カスタム関数アドインを[個人用アドイン] リストで強調表示して、windows 版 Excel にリボンを挿入します。](../images/excel-cf-register-add-in-2.png)

    * Excel Onlineを使用している場合 

        1. Excel Online で** [挿入]** ]タブを選択し、** [アドイン]** を選択します。![  [個人用アドイン]アイコンを強調表示して Excel Online でリボンを挿入します。](../images/excel-cf-online-register-add-in-1.png)

        2.  ** [個人用アドインの管理] ** を選択し、 ** [個人用アドインのアップロード] **を選択します。 

        3.   ** [参照... ]**  を選択し、Yo Officeジェネレータが作成したプロジェクトのルートディレクトリに移動します。 

        4.  ** Manifest.xml** ファイルを選択して ** 開く** を選択し、** アップロード** を選択します。

5. ここで、新しい機能を試してみましょう。 セル **B1**に、 `=CONTOSO.STOCKPRICE("MSFT")` というテキストを入力して、enter キーを押します。 セル**B1** の結果は、Microsoft 株式の1株当たりの現在の株価であることがわかります。

## <a name="create-a-streaming-asynchronous-custom-function"></a>ストリーミング非同期のカスタム関数を作成します。

 `stockPrice` 関数を作成した時点で株式の価格を返しますが、株価は常に変化しています。
 株価のリアルタイム更新を取得するために、APIからデータをストリーミングするカスタム関数を作成しましょう。

 `stockPriceStream` という名前のカスタム関数を作成するには、次の手順を実行して、1000ミリ秒ごとに指定した在庫の価格を要求します（前の要求が完了している場合）。 最初の要求が進行中に、関数が呼び出されているセルのプレースホルダ値 **#GETTING_DATA** が表示される場合があります。 関数によって値が返されると、 **#GETTING_DATA** はセル内の値に置き換えられます。

1. Yo Office ジェネレーターが作成した **株価表示** プロジェクトでは、 **src/customfunctions.js** に次のコードを追加し、ファイルを保存します。

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

2. Excelでエンドユーザーがこの新しい機能を利用できるようにする前に、この機能を説明するメタデータを指定する必要があります。 Yo Office ジェネレーターが作成した **株価表示** プロジェクトで、 `functions` 内の**config/customfunctions.json** ファイルを開き、ファイルを保存します。

    このJSONは、 `stockPriceStream` 関数を説明します。 ストリーミング機能の場合、`stream` プロパティとプロパティ`cancelable` は、次のコード例に示すように、`true`   オブジェクト 内 `options` に設定する必要があります。

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

3. エンドユーザーが新しい機能を利用できるようにするには、アドインをExcelに再登録する必要があります。 このチュートリアルで使用しているプラットフォームの次の手順を実行します。

    * Windows 版 Excel の場合

        1. Excelを終了し、再度Excelを開きます。
        
        2. Excel では、 **[挿入]** タブを選択し、 **[アドイン]** の右にある下向き矢印を選択します。![ [個人用アドイン]の矢印が強調表示された状態で windows 版 Excel にリボンを挿入します。
           ](../images/excel-cf-register-add-in-1b.png)

        3. 利用可能なアドインの一覧で、 **開発者アドイン** のセクションを検索し、 **Excelカスタム関数** アドインを選択して登録します。
            ![Excel カスタム関数アドインを[個人用アドイン] リストで強調表示して、windows 版 Excel にリボンを挿入します。](../images/excel-cf-register-add-in-2.png)

    * Excel Onlineを使用している場合 

        1. Excel Online で** [挿入]** ]タブを選択し、** [アドイン]** を選択します。![  [個人用アドイン]アイコンを強調表示して Excel Online でリボンを挿入します。](../images/excel-cf-online-register-add-in-1.png)

        2.  ** [個人用アドインの管理] ** を選択し、 ** [個人用アドインのアップロード] **を選択します。 

        3.   ** [参照... ]**  を選択し、Yo Officeジェネレータが作成したプロジェクトのルートディレクトリに移動します。 

        4.  ** Manifest.xml** ファイルを選択して ** 開く** を選択し、** アップロード** を選択します。

4. ここで、新しい機能を試してみましょう。 セル**C1** に `=CONTOSO.STOCKPRICESTREAM("MSFT")`  というテキストを入力して、enter キーを押します。 株式市場が開いている Microsoft の株 1 株のリアルタイムの価格を反映するようにセル **C1** の結果が常に更新されているはずです。

## <a name="next-steps"></a>次の手順

このチュートリアルでは、作成済みの機能を試して、新しいカスタム関数プロジェクトが完成しました。Webサイトからデータを要求し、Webサイトからリアルタイムデータをストリームするカスタム関数を作成しました。
 Excelのカスタム関数の詳細については、次の記事を参照してください。 

> [!div class="nextstepaction"]
> [Excel でカスタム関数を作成する](../excel/custom-functions-overview.md)

## <a name="legal-information"></a>法的情報

 [IEX](https://iextrading.com/developer/)が無料で提供するデータです。  [IEX の使用条件](https://iextrading.com/api-exhibit-a/)を表示します。 このチュートリアルで Microsoft が IEX API を使用するのは、教育目的でのみです。

