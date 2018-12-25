---
ms.date: 12/21/2018
description: JavaScript を使用して Excel でカスタム関数を作成する。
title: Excel でのカスタム関数の作成 (プレビュー)
ms.openlocfilehash: 8f30ee32168147b8beeb6e60372cd631237ce993
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433041"
---
# <a name="create-custom-functions-in-excel-preview"></a>Excel でのカスタム関数の作成 (プレビュー)

開発者は、カスタム関数を使用して関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。 ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。 この記事では、Excel でカスタム関数を作成する方法について説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

次の図は、エンドユーザーが Excel ワークシートのセルにカスタム関数を挿入する様子を示します。 `CONTOSO.ADD42` カスタム関数は、関数への入力パラメーターとしてユーザーが指定した数値のペアに 42 を追加するように設計されています。

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

`ADD42` カスタム関数は次のコードにより定義されます。

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。

## <a name="components-of-a-custom-functions-add-in-project"></a>カスタム関数 アドイン プロジェクトのコンポーネント

[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Excel のカスタム関数アドイン プロジェクトを作成する場合、ジェネレーターが作成するプロジェクトに以下のようなファイルが表示されます。

| ファイル | ファイル形式 | 説明 |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>または<br/>**./src/functions/functions.ts** | JavaScript<br/>または<br/>TypeScript | カスタム関数を定義するコードが含みます。 |
| **./src/functions/functions.json** | JSON | カスタム関数を定義し、Excel に関数を登録してエンドユーザーが使用できるようにするためのメタデータを含みます。 |
| **./src/functions/functions.html** | HTML | カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。 |
| **./manifest.xml** | XML | アドイン内のすべてのカスタム関数の名前空間と、この表で前述した JavaScript、JSON、HTML ファイルの位置を指定します。 |

次のセクションでは、これらのファイルに関する詳細について説明します。

### <a name="script-file"></a>スクリプト ファイル

スクリプト ファイル (Yo Office ジェネレーターが作成するプロジェクト内の **./src/customfunctions.js** または **/src/customfunctions.ts**) には、カスタム関数を定義して、カスタム関数の名前を [JSON メタデータ ファイル](#json-metadata-file)のオブジェクトにマップするコードが含まれています。 

例えば、次のコードはカスタム関数 `add` と `increment` を定義し、両方の関数のマッピング情報を指定します。  `add` 関数は、`id` プロパティの値が **ADD** の JSON メタデータ ファイル内のオブジェクトにマップされ、`increment` 関数は、`id` プロパティの値が **INCREMENT** のメタデータ ファイル内のオブジェクトにマップされます。 JSON メタデータ ファイル内のオブジェクトへのスクリプト ファイル内関数名のマッピングの詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)」を参照してください。

```js
function add(first, second){
  return first + second;
}

function increment(incrementBy, callback) {
  var result = 0;
  var timer = setInterval(function() {
    result += incrementBy;
    callback.setResult(result);
  }, 1000);

  callback.onCanceled = function() {
    clearInterval(timer);
  };
}

// map `id` values in the JSON metadata file to the JavaScript function names
CustomFunctionMappings.ADD = add;
CustomFunctionMappings.INCREMENT = increment;
```

### <a name="json-metadata-file"></a>JSON メタデータ ファイル 

カスタム関数のメタデータ ファイル (Yo Office ジェネレーターが作成するプロジェクトでは **./config/customfunctions.json**) は、Excel がカスタム関数の登録し、エンドユーザーが利用できるようするために必要な情報を提供します。 カスタム関数は、ユーザーがアドインを初めて実行するときに登録されます。 その後は、同じユーザーに対しては、(アドインが最初に実行されたワークブック内のみでなく) すべてのワークブック内で利用が可能になります。

> [!TIP]
> JSON ファイルをホストするサーバーでは、カスタム関数を Excel Online で正しく作動させるために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効に設定する必要があります。

**functions.json** の次のコードは、`add` 関数のメタデータと上述の `increment` 関数を指定します。 このコード サンプルに続く表では、JSON オブジェクト内の個別のプロパティについての詳細情報を提供します。 JSON メタデータ ファイル内の `id` と `name` 各プロパティーの値の指定に関する詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)」を参照してください。

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "first",
          "description": "first number to add",
          "type": "number",
          "dimensionality": "scalar"
        },
        {
          "name": "second",
          "description": "second number to add",
          "type": "number",
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      "description": "Periodically increment a value",
      "helpUrl": "http://www.contoso.com",
      "result": {
          "type": "number",
          "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "increment",
            "description": "Amount to increment",
            "type": "number",
            "dimensionality": "scalar"
        }
    ],
    "options": {
        "cancelable": true,
        "stream": true
      }
    }
  ]
}
```

次の表は、JSON メタデータ ファイルに通常格納されているプロパティの一覧表示です。 JSON メタデータ ファイルの詳細については、「[カスタム関数のメタデータ](custom-functions-json.md)」を参照してください。

| プロパティ  | 説明 |
|---------|---------|
| `id` | 関数の一意の ID です。 この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。 |
| `name` | Excel でエンド ユーザーに表示される関数の名前です。 Excel では、この関数名は [XML マニフェスト ファイル](#manifest-file)で指定されているカスタム関数の名前空間でプレフィックスされます。 |
| `helpUrl` | ユーザーがヘルプを要求したときに表示されるページの URL です。 |
| `description` | 関数について説明します。 この値は、関数が Excel 内のオートコンプリート メニューで選択された項目となっている場合に、ツールヒントとして表示されます。 |
| `result`  | 関数が返す情報の種類を定義するオブジェクトです。 このオブジェクトに関する詳細情報については [result](custom-functions-json.md#result) を参照してください。 |
| `parameters` | 関数の入力パラメーターを定義する配列です。 このオブジェクトに関する詳細情報については [parameters](custom-functions-json.md#parameters) を参照してください。 |
| `options` | Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 このプロパティの使用方法の詳細については、[ストリーム関数](#streaming-functions)および[関数のキャンセル](#canceling-a-function)を参照してください。 |

### <a name="manifest-file"></a>マニフェスト ファイル

カスタム関数 (Yo Office ジェネレーターが作成するプロジェクトでは **./manifest.xml**) を定義するアドインの XML マニフェスト ファイルは、アドイン内のすべてのカスタム関数の名前空間と、 JavaScript、JSON、および HTML の場所を指定します。 次の XML マークアップでは、`<ExtensionPoint>` と `<Resources>` カスタム関数を有効にするアドインのマニフェストに含める必要がある要素の例を示します。  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
        <Hosts>
            <Host xsi:type="Workbook">
                <AllFormFactors>
                    <ExtensionPoint xsi:type="CustomFunctions">
                        <Script>
                            <SourceLocation resid="Contoso.Functions.Script.Url" />
                        </Script>
                        <Page>
                            <SourceLocation resid="Contoso.Functions.Page.Url"/>
                        </Page>
                        <Metadata>
                            <SourceLocation resid="Contoso.Functions.Metadata.Url" />
                        </Metadata>
                        <Namespace resid="Contoso.Functions.Namespace" />
                    </ExtensionPoint>
                </AllFormFactors>
            </Host>
        </Hosts>
        <Resources>
            <bt:Images>
                <bt:Image id="Contoso.tpicon_16x16" DefaultValue="https://localhost:3000/assets/icon-16.png" />
                <bt:Image id="Contoso.tpicon_32x32" DefaultValue="https://localhost:3000/assets/icon-32.png" />
                <bt:Image id="Contoso.tpicon_80x80" DefaultValue="https://localhost:3000/assets/icon-80.png" />
            </bt:Images>
            <bt:Urls>
                <bt:Url id="Contoso.Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js" />
                <bt:Url id="Contoso.Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json" />
                <bt:Url id="Contoso.Functions.Page.Url" DefaultValue="https://localhost:3000/dist/functions.html" />
                <bt:Url id="Contoso.Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
            </bt:Urls>
            <bt:ShortStrings>
                <bt:String id="Contoso.Functions.Namespace" DefaultValue="CONTOSO" />
            </bt:ShortStrings>
        </Resources>
    </VersionOverrides>
```

> [!NOTE]
> Excel の関数は、XML マニフェスト ファイルで指定された名前空間が接頭辞として付加されます。 関数の名前空間は、関数名の前に付けられ、ピリオドで区切られます。 例えば、Excel ワークシートのセル内で、`ADD42` 関数を呼び出すためには、`=CONTOSO.ADD42` と入力します。これは、`CONTOSO` が名前空間で、`ADD42` が JSON ファイルで指定された関数の名前だからです。 名前空間は、会社またはアドインの識別子としての使用を目的としています。 名前空間にはアルファベットとピリオドのみを含めることが出来ます。

## <a name="functions-that-return-data-from-external-sources"></a>外部ソースからデータを返す関数

カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。

1. JavaScript Promise を Excel に返します。

2. コールバック関数を使用して Promise を最終値で解決します。

カスタム関数は、Excel での最終結果を待つ間、`#GETTING_DATA` という一時的な結果をセルに表示します。 ユーザーは、結果を待つ間もワークシートの残りの部分を通常通り操作することができます。

次のコード例では、`getTemperature()` カスタム関数が温度計の現在の温度を取得します。  `sendWebRequest` は、[XHR](custom-functions-runtime.md#xhr-example) を使用して温度 Web サービスを呼び出す仮想の関数 (ここでは指定なし) であることに留意してください。

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a>ストリーミング関数

ストリーム カスタム関数を使用すると、セルに繰り返しデータを長期的に出力でき、ユーザーが再計算を明示的に要求することは特に必要ありません。 以下のコード サンプルは、毎秒ごとに結果に数値を追加するカスタム関数です。 このコードについては、次の点に注意してください。

- Excel は、`setResult` コールバックを使用して自動的に新しい値を表示します。

- 2 番目の入力パラメーターの `handler` は、[オートコンプリート] メニューから関数が選択された場合、Excel のエンドユーザーに表示されません。

- `onCanceled` コールバックは、関数がキャンセルされた場合に実行される関数を定義します。 すべてのストリーム関数には、このようなキャンセル ハンドラーの実装が必要です。 詳細については、「[関数をキャンセルする](#canceling-a-function)」を参照してください。

```js
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}
```

JSON メタデータ ファイルでストリーミング関数にメタデータを指定する場合には、`options` オブジェクト内のプロパティ`"cancelable": true` および `"stream": true` を以下の例のように設定する必要があります。

```json
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a>関数をキャンセルする

状況によっては、帯域幅の消費量、作業メモリ、UPC への負荷を軽減するために、ストリーム カスタム関数の実行をキャンセルする必要があります。 Excel では、次のような状況で関数の実行をキャンセルします。

- ユーザーが、関数を参照するセルを編集または削除した場合。

- 関数の引数 (入力) の 1 つが変更されたとき。 この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。

- ユーザーが手動で再計算をトリガーしたとき。 この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。

関数をキャンセルする機能を有効にするには、JavaScript 関数内にキャンセル ハンドラーを実装し、関数を記述するJSONのメタデータの `options` オブジェクト内のプロパティ `"cancelable": true` を指定する必要があります。 この記事の前のセクションのコード サンプルに、これらの手法の例が示されています。

## <a name="saving-and-sharing-state"></a>状態の保存と共有

カスタム関数は、グローバル JavaScript 変数にデータを保存でき、以降の呼び出しで使用することができます。 保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を呼び出す場合に便利です。 たとえば、同じ Web リソースへの追加呼び出しを避けるために、呼び出しから返されたデータを Web リソースに保存することができます。

次のコード サンプルでは、状態をグローバルに保存する温度ストリーミング関数の実装を示します。 このコードについては、次の点に注意してください。

- `streamTemperature` 関数がセルに表示される温度の値を毎秒更新し、`savedTemperatures` 変数をデータ ソースとして使用します。

- `streamTemperature` はストリーム関数であるため、その関数がキャンセルされたときに実行されるキャンセル ハンドラーを実装します。

- ユーザーが `streamTemperature` 関数を Excel の複数のセルから呼び出す場合、`streamTemperature` 関数は実行のたびに、同じ `savedTemperatures` 変数からのデータを読み取ります。 

- `refreshTemperature` 関数は、特定の温度計の温度を毎秒読み取り、結果を `savedTemperatures` 変数に格納します。 `refreshTemperature` 関数は、Excel でエンド ユーザーには公開されないので、JSON ファイルに登録する必要はありません。

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
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

## <a name="working-with-ranges-of-data"></a>データの範囲を使用する

カスタム関数は、データの範囲を入力パラメーターとして受け入れることができ、また、データの範囲を返すこともできます。 JavaScript では、データの範囲は 2 次元配列として表されます。

例えば、関数が Excel に保存されている数値の範囲から 2 番目に大きい値を返すとします。 次の関数は、`Excel.CustomFunctionDimensionality.matrix` 型の `values` パラメーターを受け入れます。 なお、この関数の JSON メタデータでは、パラメーターの`type`プロパティを`matrix` と設定します。

```js
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
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

## <a name="discovering-cells-that-invoke-custom-functions"></a>カスタム関数を呼び出すセルを検出する

カスタム関数を使用すると、範囲の書式設定、キャッシュされた値の表示、およびを `caller.address` を使用しての値の調整を行うこともでき、カスタム関数を呼び出すセルを検出することができます。 次のシナリオの一部で `caller.address` を使用します。

- 範囲の書式設定: [AsyncStorage](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data)で情報を格納するセルのキーとして `caller.address` を使用します。 Excel で [onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet#oncalculated) を使用して`AsyncStorage` からキーを読み込みます。
- キャッシュされた値を表示させる: 関数がオフラインで使用される場合、`onCalculated` を使用して `AsyncStorage` に格納されているキャッシュされた値を表示します。
- 調整: `caller.address` を使用して元のセルを検出し、処理が発生している場所での調整を行えます。

セルのアドレスに関する情報は、関数の JSON メタデータ ファイルで `requiresAddress` が`true` とマークされている場合にのみ公開されます。 これの例を次のサンプルに示します。

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

セルのアドレスを検索するために、スクリプト ファイル (**./src/customfunctions.js**または **./src/customfunctions.ts**) に `getAddress` 関数を追加する必要があります。 この関数は、次のサンプルで示される `parameter1` のようなパラメーターを受け取ることができます。 最後のパラメーターは常に `invocationContext` で、これはJSON メタデータ ファイルで `requiresAddress` が `true` とマークされているときに Excel が返すセルの位置が格納されているオブジェクトのことです。

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

既定では、`getAddress` 関数が返す値は次の形式に従います: `SheetName!CellNumber`。 たとえば、ある関数が Expenses という名前のシートのセル B2 から呼び出される場合の戻り値は `Expenses!B2` になります。

## <a name="handling-errors"></a>エラーの処理

カスタム関数を定義するアドインをビルドする場合は、実行時エラーを考慮して、エラー処理ロジックを含めるようにします。 カスタム関数のエラー処理は、[全体的な Excel の JavaScript API のエラー処理](excel-add-ins-error-handling.md)と同じです。 次のコード サンプルでは、`.catch` がコード内で以前に発生したエラーを処理します。

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;

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

## <a name="known-issues"></a>既知の問題

- ヘルプの URL とパラメーターの説明。Excel ではまだ使用されていません。
- カスタム関数は現在、モバイル クライアント用の Excel では使用できません。
- 揮発性関数 (スプレッドシート内の無関係なデータが変更されたときに自動的に再計算する関数) はまだサポートされていません。
- Office 365 管理ポータルと AppSource による展開は、まだ有効になっていません。
- Excel Onlineでのカスタム関数は、一定期間動作していないと、セッション中に停止することがあります。 ブラウザーのページを更新 (F5) し、機能を復元するカスタム関数を再入力します。
- Windows 版 Excel で複数のアドインが実行されている場合、ワークシートのセル内に **#GETTING_DATA** という一時的な結果が表示されることがあります。 その場合には、Excel のウィンドウをすべて閉じ、Excel を再起動します。
- 今後、カスタム関数向けのデバッグ ツールが利用できるようになる可能性があります。 それまでは、F12 開発者ツールを使用して Excel Online をデバッグすることができます。 詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md)」を参照してください。
- 32 ビット版の Office 365 *December* インサイダー バージョン 1901 (ビルド 11128.20000) では、カスタム関数が正常に動作しない可能性があります。 https://github.com/OfficeDev/Excel-Custom-Functions/blob/december-insiders-workaround/excel-udf-host.win32.bundle でファイルをダウンロードして、このバグを回避できる場合があります。 それから、"C:\Program Files (x86)\Microsoft Office\root\Office16" フォルダーにそれをコピーします。

## <a name="changelog"></a>変更ログ

- **2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開*
- **2017 年 11 月 20 日**: ビルド 8801 以降を使用する場合の互換性バグを修正
- **2017 年 11 月 28 日**: 非同期関数のキャンセルのサポートを公開* (ストリーミング機能の変更が必要)
- **2018 年 5 月 7 日**: Mac、Excel Online、およびインプロセスで実行される同期関数へのサポートを公開*
- **2018 年 9 月 20日**: JavaScript ランタイムのカスタム関数へのサポートを公開。 詳細については、「[Excel カスタム関数のランタイム](custom-functions-runtime.md)」をご覧ください。
- **2018 年 10 月 20 日**: [10 月の Insider ビルド](https://support.office.com/ja-JP/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24)では、カスタム関数は、 Windows デスクトップ用およびオンライン用の[カスタム定義メタデータ](custom-functions-json.md)で 'id' パラメーターが必要になりました。 Mac では、このパラメーターは無視します。


\* は、[Office Insider](https://products.office.com/office-insider) チャンネル (旧称 "Insider Fast") 

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [チュートリアル: Excel でカスタム関数を作成します。](excel-tutorial-custom-functions.md)
