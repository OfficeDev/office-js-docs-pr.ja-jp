# <a name="create-custom-functions-in-excel-preview"></a>Excel でのカスタム関数の作成 (プレビュー)

カスタム関数（ユーザー定義関数 UDF と同様のもの）を使用すると、開発者はアドインを使用して任意の JavaScript 関数を Excel に追加できます。 ユーザーは、Excel の他のネイティブ関数（`=SUM()` など）と同様に、カスタム関数にアクセスできます。 この記事では、Excel でカスタム関数を作成する方法について説明します。

次の図は、エンドユーザーがカスタム関数をセルに挿入する方法を示しています。 1 組の数字に 42 を加える関数。

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

同じカスタム関数のコードは次のとおりです。

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

カスタム機能は、Windows、Mac、および Excel Online の開発者プレビューで利用できるようになりました。 以下の手順に従って試してみましょう。

1. Office（Windows では build 9325、Mac では 13.329）をインストールし、 [Office Insider](https://products.office.com/office-insider) プログラムに参加します。 （最新のビルドを入手するだけでは不十分であることに注意してください。Insider プログラムに参加するまでは、どのビルドでも機能が無効になります）
2. [Yo Office](https://github.com/OfficeDev/generator-office) を使用して Excel カスタム関数のアドインを作成し、[プロジェクトの README.md](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) の指示に従って Excel でアドインを起動し、コードを変更してデバッグします。
3. 任意のセルに `=CONTOSO.ADD42(1,2)` を入力し、**Enter** を押してカスタム関数を実行します。

この記事の末尾にある **既知の問題** のセクションを参照してください。このセクションには、カスタム関数の現在の制約が記載されており、時間の経過に従って更新されます。

## <a name="learn-the-basics"></a>基本操作の説明

複製されたサンプル リポジトリには、次のファイルが表示されます。

- **./src/customfunctions.js** カスタム関数のコードが含まれています (`ADD42` 関数の上の単純なコード例を参照してください)。
- **./config/customfunctions.json** カスタム関数について Excel に通知する登録 JSON が含まれています。 登録すると、ユーザーがセルに入力するときに表示される使用可能な関数のリストにカスタム関数が表示されます。
- **./index.html** JS ファイルへの &lt;Script&gt; 参照を提供します。 このファイルでは、Excel の UI は表示されません。
- **./manifest.xml** HTML、JavaScript、および JSON ファイルの場所を Excel に通知します。また、アドインと共にインストールされているすべてのカスタム関数の名前空間も指定します。

### <a name="json-file-configcustomfunctionsjson"></a>JSON ファイル (./config/customfunctions.json)

customfunctions.json の以下のコードは、同じ `ADD42` 関数のメタデータを指定します。

> [!NOTE]
> この例で使用されていないオプションを含むJSONファイルの詳細な参照情報は、「[カスタム関数登録 JSON](custom-functions-json.md)」 にあります。。

この例では、以下のことに注意してください。

- カスタム関数は1つしかないので、 `functions` ARRAY のメンバーも1つです。
- `name` プロパティは関数名を定義します。 前に示したアニメーションGIFのように、名前空間（`CONTOSO`）は、Excel オートコンプリート メニューの関数名の前に付加されます。 このプレフィックスは、後述するアドインマニフェストで定義されます。 プレフィックスと関数名はピリオドで区切られ、慣例では接頭辞と関数名は大文字です。 カスタム関数を使用するには、ユーザーが名前空間に続けて関数の名前（`ADD42` ）をセルに入力します。この場合、 `=CONTOSO.ADD42` です。 プレフィックスは、所属する会社やアドインの識別子として使用することが想定されています。 
- Excel のオートコンプリート メニュー `description` 表示されます。
- ユーザーが関数のヘルプを要求すると、Excel は作業ウィンドウを開き、`helpUrl` に指定された URL にある Web ページを表示します。
- `result`プロパティは、関数が Excel に返す情報の種類を指定します。 子のプロパティは `"string"`、 `"number"`、または `"boolean"` ができます。。`type` プロパティは `scalar` または `matrix` （指定された`type` の値の2次元配列）とすることができます。`dimensionality`
- 配列は、 関数に渡される各パラメーターのデータの種類を *順番に* 指定します。`parameters` と `description` 子のプロパティは Excel intellisense で使用されます。`name` と `dimensionality` 子のプロパティは上記で説明した `result` プロパティの子プロパティと同じです。`type`
- プロパティを使用すると、Excel がいつどのようにして関数を実行するかについてのいくつかの側面をカスタマイズできます。`options` これらのオプションについての詳細がこの記事の後半にあります。

```js
    {
        "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
        "functions": [
            {
                "name": "ADD42", 
                "description":  "adds 42 to the input numbers",
                "helpUrl": "http://dev.office.com",
                "result": {
                    "type": "number",
                    "dimensionality": "scalar"
                },
                "parameters": [
                    {
                        "name": "number 1",
                        "description": "the first number to be added",
                        "type": "number",
                        "dimensionality": "scalar"
                    },
                    {
                        "name": "number 2",
                        "description": "the second number to be added",
                        "type": "number",
                        "dimensionality": "scalar"
                    }
                ],
                "options": {
                    "sync": true
                }
            }
        ]
    }
```

> [!NOTE]
> カスタム関数は、ユーザーが最初にアドインを実行したときに登録されます。 その後、同じユーザーに対して、すべてのブック（アドインが最初に実行されたものだけでなく）で関数を使用できます。

JSON ファイルのサーバー設定では、カスタム関数が Excel Online で正しく作動するために [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) が有効になっていなければなりません。


### <a name="manifest-file-manifestxml"></a>マニフェスト ファイル (manifest.xml)


以下は、作成した関数を Excel が実行できるようにアドインのマニフェストに組み込んだ `<ExtensionPoint>` および `<Resources>` マークアップの例です。 このマークアップについて、次の点に注意してください。

- 要素とそれに対応するリソース ID は、関数で JavaScript ファイルの場所を指定します。`<Script>`
- 要素とそれに対応するリソース ID は、アドインの HTML ページの場所を指定します。`<Page>` HTML ページには、JavaScript ファイル（customfunctions.js）を読み込む `<Script>` タグが含まれています。 HTML ページは非表示のページであり、UI に表示されることはありません。
- 要素とそれに対応するリソース ID は、JSON ファイルの場所を指定します。`<Metadata>`
- 要素および対応するリソース ID は、アドインのすべてのカスタム関数のプレフィックスを指定します。`<Namespace>`


```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="residjs" />
                    </Script>
                    <Page>
                        <SourceLocation resid="residhtml"/>
                    </Page>
                    <Metadata>
                        <SourceLocation resid="residjson" />
                    </Metadata>
                    <Namespace resid="residNS" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="residjson" DefaultValue="http://127.0.0.1:8080/customfunctions.json" />
            <bt:Url id="residjs" DefaultValue="http://127.0.0.1:8080/customfunctions.js" />
            <bt:Url id="residhtml" DefaultValue="http://127.0.0.1:8080/customfunctions.html" />
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="residNS" DefaultValue="CONTOSO" />
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>

```

## <a name="initializing-custom-functions"></a>カスタム関数の初期化

コードは、使用する前にカスタム関数の機能を初期化する必要があります。 初期化は、HTML ファイル （customfunctions.html）の &lt;Script&gt; タグ、または JavaScript ファイル（customfuntions.js）のトップで実行できます。 カスタム関数のプレビュー中に、初期化のための 2 つの構文を選択できます。 リポジトリ内の HTML ファイルは、次の構文を使用します。

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

次の構文も使用できます。

```js
Office.Preview.StartCustomFunctions();
```

## <a name="handling-errors"></a>エラーの処理
カスタム関数のエラー処理は、[一般的な Excel JavaScript API のエラー処理](./excel-add-ins-error-handling.md) と同じです。 一般的に、エラー処理には `.catch` を使用します。 次のコードは、`.catch` の例を示しています。 

```js
function getComment(x) {
    var url = "https://jsonplaceholder.typicode.com/comments/" + x; //this delivers a section of lorem ipsum from the jsonplaceholder API
    return fetch(url)
        .then(function (data) {
            return data.json();
        })
        .then((json) => {
            return json.body;
        })
        .catch(function (error) {
            throw error;
        })
}
```

## <a name="synchronous-and-asynchronous-functions"></a>同期関数と非同期関数

上記の `ADD42` 関数は Excel （JSON ファイルのオプション `"sync": true` を使用して指定したもの ）と同期しています。 同期関数は、Excel と同じプロセスで実行され、マルチスレッド計算中に並行して実行されるため、高速なパフォーマンスを提供します。   

一方、カスタム関数が Web からデータを取得する場合は、Excel と非同期でなければなりません。 非同期関数は以下を実行する必要があります。

1. JavaScript Promise を Excel に返します。
3. コールバック関数を使用して Promise を最終値で解決します。

次のコードは、温度計の温度を取得する非同期カスタム関数の例を示しています。 は、XHR を使用して温度 Web サービスを呼び出す、ここでは指定されていない仮想関数であることにご注意ください。`sendWebRequest`

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

非同期関数は、 Excelが最終結果を待つ間、セルに `GETTING_DATA` 一時的エラーを表示します。 ユーザーは、結果を待つ間、スプレッドシートの他の部分と通常通りやりとりすることができます。

> [!NOTE]
> カスタム関数は既定では非同期です。 同期として関数を指定するには、登録 JSON ファイル内のカスタム関数の `options` プロパティでオプション `"sync": true` を設定してください。

## <a name="streamed-functions"></a>ストリーム関数

非同期関数をストリーミングできます。 カスタムのストリーム関数を使用すると、Excel やユーザーが再計算を要求するのを待たずに、時間の経過に従ってセルに繰り返しデータを出力できます。 次の例は、1 秒おきに結果に数値を追加するカスタム関数です。 このコードについては、次の点に注意してください。

- Excel は、`setResult`コールバックを使用して自動的に新しい値を表示します。
- 最終的なパラメータ `caller` は登録コードでは指定されず、Excel ユーザーが関数を入力するときにオートコンプリート メニューに表示されません。 これは、関数のデータを Excel に渡してセルの値を更新するために使用される `setResult` コールバック関数を含むオブジェクトです。
- Excel が `caller` オブジェクト内の `setResult` 関数を渡すには、関数登録の際に、登録 JSON ファイル内のカスタム関数の `options` プロパティでオプション `"stream": true` を設定して、ストリーミングのサポートを宣言する必要があります。

```js
function incrementValue(increment, caller){
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a>キャンセル

ストリーム関数と非同期関数をキャンセルできます。 関数呼び出しのキャンセルは、帯域幅の使用量、作業メモリ、および CPU の負荷を減らすために重要です。 Excel では、次のような状況で関数の呼び出しをキャンセルします。

- ユーザーが関数を参照するセルを編集または削除する。
- 関数の引数 (入力) の 1 つが変更される。 この場合、キャンセルに加えて新しい関数の呼び出しがトリガーされます。
- ユーザーは手動で再計算をトリガーします。上記の場合と同様に、キャンセルに加えて新しい関数の呼び出しがトリガーされます。

すべてのストリーミング関数に対してキャンセル ハンドラを実装することが *必須* です。 非同期の非ストリーミング関数は、キャンセル可能にもキャンセル不可にもでき、ご自分で決定できます。 同期機能はキャンセルすることはできません。

関数をキャンセル可能にするには、登録 JSON ファイル内のカスタム関数の `options` プロパティでオプション `"cancelable": true` を設定してください。

次のコードでは、前述の例にキャンセルを実装しています。 このコードでは、`caller` オブジェクトに `onCanceled` 関数が含まれており、キャンセル可能な各カスタム関数ごとに定義する必要があります。

```js
function incrementValue(increment, caller){ 
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);

    caller.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a>状態の保存と共有

非同期カスタム関数では、JavaScript のグローバル変数にデータを保存できます。 後続の呼び出しでは、カスタム関数はこれらの変数に保存されている値を使用できます。 保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を追加する場合に便利です。 たとえば、同じ Web リソースへの追加呼び出しを避けるために、呼び出しから返されたデータを Web リソースに保存することができます。

次のコードは、 状態をグローバルで保存する前述の温度ストリーミング関数の実装を示しています。 このコードについては、次の点に注意してください。

- `refreshTemperature` は、1 秒おきに特定の温度計の温度を読み取るストリーム関数です。 新しい温度は `savedTemperatures` 変数に保存されますが、セルの値を直接更新することはありません。 ワークシート・セルから直接呼び出されません。*したがって、JSON ファイルには登録されません *
- `streamTemperature` 1 秒おきにセルに表示される温度値を更新します。また、 `savedTemperatures` 変数をデータソースとして使用します。 JSON ファイルに登録し、すべて大文字で `STREAMTEMPERATURE` という名前をつける必要があります。
- ユーザーは、Excel UI の複数のセルから `streamTemperature` を呼び出すことができます。 呼び出すたびに、同じ `savedTemperatures` 変数からデータを読み取ります。

```js
var savedTemperatures;

function streamTemperature(thermometerID, caller){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
     }

     function getNextTemperature(){
         caller.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
     }
     getNextTemperature();
}

function refreshTemperature(thermometerID){
     sendWebRequest(thermometerID, function(data){
         savedTemperatures[thermometerID] = data.temperature;
     });
     setTimeout(function(){
         refreshTemperature(thermometerID);
     }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

> [!NOTE]
> 同期関数（JSON ファイル内のオプション `"sync": true` で指定されたもの）は、Excel がマルチスレッド計算中にそれらを並行して行うため、状態を共有できません。 アドインの同期関数が各セッションで同じ JavaScript コンテキストを共有するため、非同期関数のみが状態を共有できます。

## <a name="working-with-ranges-of-data"></a>データの範囲を使用する

カスタム関数は、データ範囲をパラメーターとして受け取ったり、カスタム関数からデータ範囲を返したりすることができます。

たとえば、関数が Excel に格納されている数値の範囲から 2 番目に高い値を返すとします。 次の関数は、パラメーター `values` を取ります。これは `Excel.CustomFunctionDimensionality.matrix` パラメーター型です。 この関数の登録 JSON では、パラメータの `type` プロパティを `matrix` に設定するよう注意してください。

```js
function secondHighest(values){ 
     var highest = values[0][0], secondHighest = values[0][0];
     for(var i = 0; i < values.length; i++){
         for(var j = 1; j < values[i].length; j++){
             if(values[i][j] >= highest){
                 secondHighest = highest;
                 highest = values[i][j];
             }
             else if(values[i][j] >= secondHighest){
                 secondHighest = values[i][j];
             }
         }
     }
     return secondHighest;
 }
```

ご覧のとおり、範囲は JavaScript で行配列の配列（2次元配列など）として処理されます。

## <a name="known-issues"></a>既知の問題

- ヘルプの URL とパラメーターの説明。Excel ではまだ使用されていません。
- カスタム機能は現在、モバイル クライアント用の Excel では使用できません。
- 現在、アドインは、非同期関数カスタム関数を実行するために非表示ブラウザ プロセスに依存しています。 カスタム関数をより高速にし、使用メモリを少なくするために、今後 JavaScript はいくつかのプラットフォームで直接実行されるようになります。 さらに、マニフェストの `<Page>` 要素によって参照される HTML ページは、Excel が JavaScript を直接実行するようになるため、ほとんどのプラットフォームで不要になります。 この変更に備えるため、カスタム関数が Web ページ DOM を使用しないことを徹底してください。 Web にアクセスするためにサポートされているホスト API は、GET または POST を使用する [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) および [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) になります。
- 揮発性関数（スプレッドシート内の無関係なデータが変更されたときに自動的に再計算する関数）はまだサポートされていません。
- デバッグは、Excel for Windows の非同期関数に対してのみ有効です。
- Office 365 管理ポータルと AppSource による展開はまだ有効になっていません。
- Excel Online のカスタム関数は、一定期間使用しないとセッション中に機能しなくなることがあります。 ブラウザページを更新（F5）し、カスタム関数を再入力して機能を復元します。

## <a name="changelog"></a>変更ログ

- **2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開
- **2017 年 11 月 20 日**: ビルド 8801 以降を使用する場合の互換性バグを修正
- **2017 年 11 月 28 日**: 非同期関数のキャンセルのサポートを公開 (ストリーミング機能の変更が必要)
- **2018 年 5 月 7 日**：Mac、Excel Online、およびインプロセスで実行される同期関数のサポートを公開
