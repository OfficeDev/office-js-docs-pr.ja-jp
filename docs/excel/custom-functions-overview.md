# <a name="create-custom-functions-in-excel-preview"></a>Excel でのカスタム関数の作成 (プレビュー)

カスタム関数 (ユーザー定義関数、つまり UDF に似ている) を使用すると、開発者はアドインを使用して任意の JavaScript 関数を Excel に追加できます。 ユーザーは Excel の他のネイティブ関数 (= SUM() など) のようなカスタム関数にアクセスできるようになります。 この記事では、Excel でカスタム関数を作成する方法について説明します。

Excel でのカスタム関数がどのようなものかを以下に示します。

<img src="../../images/custom-function.gif" width="579" height="383" />

数値のペアに 42 を追加するサンプル カスタム関数のコードを示します。

```js
function add42 (a, b) {
    return a + b + 42;
}
```

カスタム関数がプレビューで利用できるようになりました。 以下の手順に従って試してみましょう。

1.  [Office Insider](https://products.office.com/ja-JP/office-insider) プログラムに参加して、コンピューターに、カスタム関数に必要な Excel 2016 のバージョン (バージョン 16.8711 以降) をインストールします。
2.  [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) リポジトリを複製し、*README.md* の指示に従って Excel でアドインを開始してください。
3.  任意のセルに `=CONTOSO.ADD42(1,2)` を入力し、**Enter** を押してカスタム関数を実行します。
4.  質問がある場合は、Stack Overflow で [office-js](https://stackoverflow.com/questions/tagged/office-js) タグを付けて質問してください。

この記事の末尾にある既知の問題のセクションを参照してください。このセクションには、カスタム関数の現在の制限が記載されており、時間の経過に従って更新されます。

## <a name="learn-the-basics"></a>基本操作の説明


複製されたサンプル リポジトリには、次のファイルが表示されます。

-   *customfunctions.js*。内容は次のとおりです。

    -   Excel に追加するカスタム関数のコード。
    -   カスタム関数を Excel に接続するための登録コード。 登録すると、ユーザーがセルに入力するときに表示される使用可能な関数のリストにカスタム関数が表示されます。
-   *customfunctions.html*。*customfunctions.js* への &lt;Script &gt;参照を提供します。 このファイルでは、Excel の UI は表示されません。
-   *manifest.xml*。カスタム関数の実行に必要な HTML および JS ファイルの場所を Excel に通知します。

### <a name="javascript-file-customfunctionsjs"></a>JavaScript ファイル (*customfunctions.js*)

customfunctions.js の以下のコードは、カスタム関数 `add42` を宣言してから、その関数を Excel に登録します。

```js
function add42 (a, b) {
    return a + b + 42;
}

Excel.Script.customFunctions["CONTOSO"]["ADD42"] = {
    call: add42,
    description: "Adds 42 to the sum of two numbers",
    helpUrl: "https://www.contoso.com/help.html",
    result: {
        resultType: Excel.CustomFunctionValueType.number,
        resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    parameters: [{
        name: "num 1",
        description: "The first number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    {
        name: "num 2",
        description: "The second number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    }],
    options:{ batch: false, stream: false }
};

Excel.run(function(ctx) {
    ctx.workbook.customFunctions.addAll();
});
```

カスタム関数の**登録**では `Excel.Script.customFunctions["CONTOSO"]["ADD42"]` コード ブロックを使用します。 Excel の関数を登録するには、次のパラメーターが必要です。

-   プレフィックスと関数名:`Excel.Script.customFunctions` の最初の値がプレフィックスです (この場合、CONTOSO がプレフィックスです)。 `Excel.Script.customFunctions` の2番目の値が関数名です (この場合、ADD42 が関数名です)。 Excel では、プレフィックスと関数名はピリオドで区切られます。カスタム関数を使用するには、関数のプレフィックス (CONTOSO) を関数の名前 (ADD42) と組み合わせ、セルに `=CONTOSO.ADD42` を入力します。 プレフィックスと関数名には、規則により、大文字を使用します。 プレフィックスは、アドインの識別子として使用することが想定されています。
-   `call`:呼び出す JavaScript 関数を定義します (たとえば、`add42`)。 JavaScript 関数の名前は、Excel で登録する名前と一致する必要はありません。
-   `description`:Excel のオートコンプリート メニューに説明が表示されます。
-   `helpUrl`:ユーザーが関数のヘルプを要求すると、Excel は作業ウィンドウを開き、この URL にある Web ページを表示します。
-   `result`:関数が返す情報の種類を Excel に定義します。

    -   `resultType`:関数は、`"string"` か `"number"` (日付と通貨にも使用されます) のいずれかを返すことができます。 詳細については、「[カスタム関数の列挙型](../../reference/excel/customfunctionsenumerations.md)」を参照してください。
    -   `resultDimensionality`:関数は、単一の (`"scalar"`) 値または値の `"matrix"` のいずれかを返すことができます。 値の行列を返すとき、関数は配列を返します。各配列要素は値の行を表す別の配列です。 詳細については、「[カスタム関数の列挙型](../../reference/excel/customfunctionsenumerations.md)」を参照してください。 次の例では、カスタム関数から 3 行 2 列の値の行列を返します。

```js
return [["first","row"],["second","row"],["third","row"]];
```

-   カスタム関数では、入力として引数を取る場合があります。 カスタム関数に渡される引数は、*parameters* プロパティで指定されます。 定義内のパラメーターの順序は、JavaScript 関数の順序と一致する必要があります。 各パラメーターには、以下のプロパティを定義します。

    -   `name`:パラメーターを表すために Excel に表示される文字列。
    -   `description`:パラメーターの詳細情報のために表示される文字列。
    -   `valueType`:`"number"` または `"string"`。これは前述の resultType プロパティと同様です。
    -   `valueDimensionality`:`"scalar"` の値、または値の `"matrix"`。これは前述の resultDimensionality プロパティと同様です。 マトリックス型のパラメーターを使用すると、ユーザーは単一のセルより大きな範囲を選択することができます。

-   `options`: 特別な種類のカスタム関数を有効にします。詳細については、この記事の後半で説明します。

`Excel.Script.customFunctions` を使用して定義したすべての関数の登録を完了するには、`CustomFunctions.addAll()` を呼び出してください。

カスタム関数は、登録後はユーザーのすべてのブック (アドインが最初に実行されたブックだけでなく) で使用できます。 関数は、ユーザーが入力を開始すると、オートコンプリート メニューに表示されます。

### <a name="manifest-file-manifestxml"></a>マニフェスト ファイル (*manifest.xml*)

manifest.xml の次の例では、Excel が関数のコードを検索することができます。

```xml

<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">

    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="scriptURL" />
                        <!— Required. The Developer Preview does not use the Script element.-->
                    </Script>
                    <Page>
                        <SourceLocation resid="pageURL"/>
                    </Page>
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>

    <Resources>
        <bt:Urls>
            <bt:Url id="scriptURL" DefaultValue="https://www.contoso.com/addin/customfunctions.js" />
            <bt:Url id="pageURL" DefaultValue="https://www.contoso.com/addin/customfunctions.html" />
        </bt:Urls>
    </Resources>

</VersionOverrides>

```

前述のコードでは、次のものを指定します。

-   &lt; `Script` &gt;要素。必須ですが開発者向けプレビューでは使用されません。
-   &lt; `Page` &gt;要素。アドインの HTML ページにリンクします。 HTML ページには、カスタム関数と登録コードを含む JavaScript ファイル (*customfunctions.js*) への &lt;Script&gt; 参照が含まれています。 HTML ページは非表示のページであり、UI に表示されることはありません。

## <a name="asynchronous-functions"></a>非同期関数

カスタム関数が Web からデータを取得する場合は、フェッチするために非同期呼び出しを行う必要があります。 外部 web サービスを呼び出すときは、カスタム関数は以下を実行する必要があります。

1.   JavaScript Promise を Excel に返します。
2.   外部のサービスを呼び出す http 要求を行います。
3.   `setResult` コールバックによってプロミスを解決します。 `setResult` が値を Excel に送信します。

次のコードは、温度計の温度を取得するカスタム関数の例を示しています。

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult, setError){
        sendWebRequestExample(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a>ストリーム関数

カスタムのストリーム関数を使用すると、Excel やユーザーが再計算を要求するのを待たずに、時間の経過に従ってセルに繰り返しデータを出力できます。 たとえば、次のコードの `incrementValue` カスタム関数は、1 秒おきに結果に数値を追加し、Excel は `setResult` コールバックを使用して自動的に新しい値を表示します。 `incrementValue` で使用されている登録コードを参照するには、*customfunctions.js* ファイルをお読みください。

```js
function incrementValue(increment, setResult){ 
     var result = 0;
     setInterval(function(){
         result += increment;
         setResult(result);
    }, 1000);
}
```

ストリーム関数の場合、最終的なパラメーター `setResult` は登録コードでは指定されず、Excel ユーザーが関数を入力するときにオートコンプリート メニューに表示されません。 これは、関数のデータを Excel に渡してセルの値を更新するために使用されるコールバック関数です。 Excel が `setResult` 関数を渡すには、関数登録時にパラメーター `stream` を `true` に設定してストリーミングのサポートを宣言する必要があります。

## <a name="saving-state"></a>状態の保存

カスタム関数では、JavaScript のグローバル変数にデータを保存できます。 後続の呼び出しでは、カスタム関数はこれらの変数に保存されている値を使用できます。 保存された状態は、ユーザーが同じカスタム関数の複数のインスタンスを入力し、相互にデータを共有する必要がある場合に便利です。 たとえば、同じ Web リソースへの追加呼び出しを避けるために、呼び出しから返されたデータを Web リソースに保存することができます。

次のコードは、`savedTemperatures` 変数を使用して状態を保存する前述の温度ストリーミング関数の実装を示しています。 このコードは、次の概念を示しています。

-   **データを保存する。** `refreshTemperature` は、1 秒おきに特定の温度計の温度を読み取るストリーム関数です。 新しい温度は、savedTemperatures 変数に保存されます。

-   **保存されたデータを使用する。** `streamTemperature` は Excel UI に表示される温度値を 1 秒おきに更新します。 温度は `savedTemperature` から読み取られ、`setResult` によって Excel UI に送信されます。 ユーザーは、Excel UI の複数のセルから `streamTemperature` を呼び出すことができます。 `streamTemperature` を呼び出すたびに `savedTemperatures` からデータが読み取られます。

> この場合は、`streamTemperature` をカスタム関数として Excel に登録します。

```js
var savedTemperatures{};

function streamTemperature(thermometerID, setResult){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID);
     }

     function getNextTemperature(){
         setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
     }
     getNextTemperature();
}

function refreshTemperature(thermometerID){
     sendWebRequestExample(thermometerID, function(data){
         savedTemperatures[thermometerID] = data.temperature;
     });
     setTimeout(function(){
         refreshTemperature(thermometerID);
     }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="working-with-ranges-of-data"></a>データの範囲を使用する

カスタム関数は、データ範囲をパラメーターとして受け取ったり、カスタム関数からデータ範囲を返したりすることができます。

たとえば、関数が Excel に格納されている温度値の範囲から2番目に高い温度を返すとします。 次の関数は、パラメーター `temperatures` を取ります。これは `Excel.CustomFunctionDimensionality.matrix` パラメーター型です。

```js
function secondHighestTemp(temperatures){ 
     var highest = -273, secondHighest = -273;
     for(var i = 0; i < temperatures.length; i++){
         for(var j = 0; j < temperatures[i].length; j++){
             if(temperatures[i][j] <= highest){
                 secondHighest = highest;
                 highest = temperatures[i][j];
             }
             else if(temperatures[i][j] <= secondHighest){
                 secondHighest = temperatures[i][j];
             }
         }
     }
     return secondHighest;
 }
```

## <a name="known-issues"></a>既知の問題

次の機能は、開発者向けプレビューで、まだサポートされていません。

-   バッチ処理。複数の呼び出しを同一の関数に集約し、パフォーマンスを向上できます。

-   取り消し。ストリーミング機能が不要になったとき (ユーザーがセルをクリアしたときなど) に通知します。 現在、関数ではセルに新しい値を書き込むのを止める時期を判断することができません。

-   ヘルプの URL とパラメーターの説明は Excel ではまだ使用されていません。

-   カスタム機能を使用する Office ストアまたはOffice 365 一元展開にアドインを公開する。

-   カスタム機能は、Mac 上の Excel、Excel for iOS、Excel Online では使用できません。

-   現在、アドインは、カスタム機能を実行するための隠しブラウザー プロセスに依存しています。 カスタム関数をより高速にし、使用メモリを少なくするために、今後 JavaScript はいくつかのプラットフォームで直接実行されるようになります。 また、マニフェストの &lt;Page&gt; 要素によって参照される HTML ページは、Excel が JavaScript を直接実行するようになれば、ほとんどのプラットフォームで不要になります。 この変更に備えるため、カスタム関数が Web ページ DOM を使用しないことを徹底してください。
