---
ms.date: 10/17/2018
description: JavaScript を使用して Excel でカスタム関数を作成します。
title: Excel でのカスタム関数の作成 (プレビュー)
ms.openlocfilehash: cc06664a0acb582344448ceec1ec36319d1c3b4c
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640100"
---
# <a name="create-custom-functions-in-excel-preview"></a>Excel でカスタム関数を作成する (プレビュー)

カスタム関数とは、開発者がアドインの一部として、JavaScriptでこれらの関数を定義してExcelに追加する新しい関数です。Excel内のユーザーは、Excel の他のネイティブ関数（ `SUM()` など）と同様に、カスタム関数にアクセスできます。 この資料ではExcelでカスタム関数を作成する方法について説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

次の図は、Excel ワークシートのセルにカスタム関数を挿入する、エンド ユーザーを示します。 `CONTOSO.ADD42` カスタム関数は、関数への入力パラメーターとしてユーザーが指定した数値のペアに 42 を追加するように設計されています。

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

次のコードは、`ADD42` カスタム関数を定義します。

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。

## <a name="components-of-a-custom-functions-add-in-project"></a>カスタム関数アドインプロジェクトのコンポーネント

[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office) を使用して Excel カスタム関数アドイン プロジェクトを作成する場合は、ジェネレーターが作成するプロジェクトに以下のようなファイルが表示されます。

| ファイル | ファイル形式 | 説明 |
|------|-------------|-------------|
| **./src/customfunctions.js**<br/>または<br/>**./src/customfunctions.ts** | JavaScript<br/>または<br/>TypeScript | カスタム関数を定義するコードを含みます。 |
| **./config/customfunctions.json** | JSON | カスタム関数を定義し、Excel に関数を登録してエンドユーザーが使用できるようにするためのメタデータを含みます。 |
| **./index.html** | HTML | カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。 |
| **./manifest.xml** | XML | アドイン内のすべてのカスタム関数の名前空間と、この表で前述した JavaScript、JSON、HTML ファイルの位置を指定します。 |

次のセクションでは、これらのファイルに関する詳細について説明します。

### <a name="script-file"></a>スクリプト ファイル 

スクリプト ファイル (Yo Office ジェネレーターが作成するプロジェクト内の **./src/customfunctions.js** または **/src/customfunctions.ts**) には、カスタム関数を定義して、カスタム関数の名前を [JSON メタデータ ファイル](#json-metadata-file)のオブジェクトにマップするコードが含まれています。 

例えば、次のコードでカスタム関数 `add` と `increment` を定義し、両方の関数のマッピング情報を指定します。 `add` 関数は、JSON メタデータ ファイル内のオブジェクトにマップされ、 この場所に`id` プロパティの値が**追加**されます。`increment` 関数は、メタデータ ファイル内のオブジェクトにマップされ、この場所に`id` プロパティの値が**インクリメント**します。JSON メタデータ ファイル内のオブジェクトへのスクリプト ファイル内関数名のマッピングの詳細については、 [カスタム関数のベスト ・ プラクティス](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) を参照してください。

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

カスタム関数のメタデータ ファイル (Yo Office ジェネレーターが作成するプロジェクトでは **./config/customfunctions.json** ) は、Excel がカスタム関数の登録を要求し、エンドユーザーが利用できるよう、情報を提供します。カスタム関数は、ユーザーがアドインを初めて実行するときに登録されます。その後は、同じユーザーに対しては、（アドインが最初に実行されたワークブック内のみでなく）すべてのワークブック内で利用が可能になります。

> [!TIP]
> JSON ファイルをホストするサーバーは、カスタム関数が Excel Online で正しく作動するために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS)  を有効に設定する必要があります。

**Customfunctions.json** の次のコードは、`add` 関数のメタデータと `increment` 、上述の関数を指定します。このコード サンプルを基にした表は、この JSON オブジェクト内の個別のプロパティについての詳細情報を提供します。`id`の値とJSON のメタデータ ファイル内の`name`プロパティの指定に関する詳細については、 [ベスト プラクティスのカスタム関数](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) を参照してください。

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

次の表は、通常、JSON メタデータ ファイルに格納されているプロパティの一覧表示です。JSON メタデータ ファイルの詳細については、 [カスタム関数のメタデータ](custom-functions-json.md)を参照してください。

| プロパティ  | 説明 |
|---------|---------|
| `id` | 関数の一意の ID です。 この ID は、英数字とピリオドのみを含めることができ、設定された後、変更してはいけません。 |
| `name` | Excel でエンドユーザーに表示される関数の名前です。Excel では、この関数名の前に、[ [XML マニフェスト ファイル](#manifest-file) ]で指定されているカスタム関数の名前空間が接頭辞として付されます。 |
| `helpUrl` | ユーザーがヘルプを要求したときに表示されるページの URL です。 |
| `description` | 関数について説明します。この値は、関数が Excel 内のオートコンプリート メニューで選択された項目となっている場合に、ツールヒントとして表示されます。 |
| `result`  | 関数によって返される情報の種類を定義するオブジェクト。`type` 子プロパティの値は、 **文字列**、 **数値**、または **ブール値**を使用できます。子プロパティの値は、 `dimensionality` **スカラー** または **マトリックス** を使用できます (指定された `type`の値の2 次元配列)。 |
| `parameters` | 関数の入力パラメーターを定義する配列。 Excel の intelliSense に `name` と `description` 子のプロパティが表示されます。 `type` 子プロパティの値には、[ **文字列**]、[ **数値**]、または [ **ブール値**]を使用できます。`dimensionality` 子プロパティの値には、[ **スカラー** ] または [ **マトリックス** ] を使用できます (指定された `type`の値の2次元配列)。 |
| `options` | Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。 このプロパティの使用方法の詳細については、この記事で後述する[ [ストリーム関数](#streaming-functions) ] および [ [関数のキャンセル](#canceling-a-function) ]を参照してください。 |

### <a name="manifest-file"></a>マニフェスト ファイル

カスタム関数 (Yo Office ジェネレーターが作成するプロジェクトでは[**./manifest.xml** ]) を定義するアドインの XML マニフェスト ファイルは、アドインとJavaScript、JSON、および HTML のロケーション内のすべてのカスタム関数の名前空間を指定します。次の XML マークアップでは、 `<ExtensionPoint>` と `<Resources>` カスタム関数を有効にするアドインのマニフェストに含める必要がある要素の一例を示します。  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="JS-URL" /> <!--resid points to location of JavaScript file-->
                    </Script>
                    <Page>
                        <SourceLocation resid="HTML-URL"/> <!--resid points to location of HTML file-->
                    </Page>
                    <Metadata>
                        <SourceLocation resid="JSON-URL" /> <!--resid points to location of JSON file-->
                    </Metadata>
                    <Namespace resid="namespace" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="JSON-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.json" /> <!--specifies the location of your JSON file-->
            <bt:Url id="JS-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.js" /> <!--specifies the location of your JavaScript file-->
            <bt:Url id="HTML-URL" DefaultValue="http://127.0.0.1:8080/index.html" /> <!--specifies the location of your HTML file-->
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. Can only contain alphanumeric characters and periods.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> Excel の関数は、XML マニフェスト ファイルで指定された名前空間が接頭辞として付加されます。関数の名前空間は、関数名の前に付けられ、ピリオドで区切られます。例えば、Excel ワークシートのセル内で、`ADD42` 関数を呼び出すためには  、 `=CONTOSO.ADD42`を入力します。これは、`CONTOSO` が、名前空間であり、`ADD42` が JSON ファイルで指定された関数の名前であるためです。名前空間は、会社またはアドインの識別子としての使用を目的としています。名前空間にはアルファベットとピリオドのみを含めることが出来ます。

## <a name="functions-that-return-data-from-external-sources"></a>外部ソースからデータを返す関数

カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。

1. JavaScript Promise を Excel に返す。

2. コールバック関数を使用して Promise を最終値で解決する。

カスタム関数は、 Excelが `#GETTING_DATA` セルの最終結果を待っている間、一時的な結果を表示します。ユーザーは、結果待機中も通常はワークシートの残りの部分を操作することができます。

次のコード例は、 現在の温度計の温度を取得する `getTemperature()` カスタム関数です。 `sendWebRequest` は、温度 web サービスを呼び出す [XHR](custom-functions-runtime.md#xhr-example) を使用した仮想関数 (ここでは指定なし) であることに留意してください。

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

ストリーム カスタム関数を使用すると、セルに繰り返しデータを長期的に出力でき、ユーザーが再計算を要求することは特に必要ありません。 以下のコード サンプルは、毎秒結果に数値を追加するカスタム関数です。 このコードについては、次の点に注意してください。

- Excel は、`setResult`コールバックを使用して自動的に新しい値を表示します。

- 2 番目の入力パラメーター`handler` は、[オートコンプリート] メニューから関数を選択したとき Excel でエンドユーザーに表示されません。

- `onCanceled` コールバックは、関数がキャンセルされた場合に実行される関数を定義します。 すべてのストリーム関数には、このようなキャンセル ハンドラーの実装が必要です。 詳細については、 [ [関数をキャンセルする](#canceling-a-function) ] を参照してください。

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

JSON メタデータ ファイルでストリーミング関数にメタデータを指定する場合には、以下の例のように、プロパティ`"cancelable": true` および `"stream": true` （ `options` オブジェクト内）を設定する必要があります。

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

状況によっては、帯域幅の消費量、作業メモリ、UPC への負荷を縮小するために、ストリーム カスタム関数の実行をキャンセルする必要もあります。 Excel では、次のような関数の実行をキャンセルします。

- ユーザーが、関数を参照するセルを編集または削除した場合。

- 関数の引数 (入力) のいずれかが変更されたとき。この例では、キャンセルの後、新しい関数の呼び出しがトリガーされます。

- ユーザーが手動で再計算をトリガーしたとき。この例では、キャンセルの後、新しい関数の呼び出しがトリガーされます。

関数をキャンセルする機能を有効にするには、JavaScript 関数内にキャンセル ハンドラーを実装し、関数を記述するJSONのメタデータの`options` オブジェクト内のプロパティ`"cancelable": true`を指定する必要があります。この記事の前のセクションのコード サンプルに、これらの手法の例が示されています。

## <a name="saving-and-sharing-state"></a>状態の保存と共有

カスタム関数は、グローバル JavaScript 変数にデータを保存でき、以降の呼び出しで使用することができます。 保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を呼び出す場合に便利です。 たとえば、Web リソースへの呼び出しから返されたデータを保存しておけば、同じ Web リソースへ繰り返し呼び出しを行わなくて済みます。

次のコード サンプルでは、グローバル状態を保存する温度ストリーミング関数の実装を示します。このコードについては、以下のことにに留意してください。

- `streamTemperature` 関数が 毎秒セルに表示される温度の値を更新し、 `savedTemperatures` 変数をデータ ソースとして使用します。

- `streamTemperature` は、ストリーム関数であるため、その関数がキャンセルされたときに実行されるキャンセル ハンドラーを実装します。

- ユーザーが `streamTemperature` 関数を Excel の複数のセルから呼び出す場合、 `streamTemperature` 関数は実行のたびに、同じ `savedTemperatures` 変数からのデータを読み取ります。 

- `refreshTemperature` 関数は、毎秒特定の温度計の温度を読み取り、結果を `savedTemperatures` 変数に格納します。 `refreshTemperature` 関数は、Excel でのエンド ユーザーには公開されないので、JSON ファイルに登録する必要はありません。

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

カスタム関数は、入力パラメーターとしてのデータの範囲を受け入れる、もしくは、データの範囲を返すことがあります。JavaScript では、データの範囲は、2 次元配列として表されます。

例えば、関数がExcel に保存されている数値の範囲から 2 番目に大きい値を返すとします。次の関数は、種類`Excel.CustomFunctionDimensionality.matrix`のパラメータ`values`を受け取ります。この関数の JSON のメタデータには、パラメーターの `type` プロパティを `matrix`に設定するよう留意してください。

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

## <a name="handling-errors"></a>エラーの処理

カスタム関数を定義するアドインをビルドする場合は、ランタイムエラーを考慮するためのエラー処理 ロジックを含めるようにしてください。カスタム関数のエラー処理は、 [大規模な Excel の JavaScript API のエラー処理](excel-add-ins-error-handling.md)と同じです。次のコード サンプルでは、 `.catch`がコード内で以前に発生したエラーを処理します。

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
- 揮発性関数（スプレッドシート内の無関係なデータが変更されたときに自動的に再計算する関数）はまだサポートされていません。
- Office 365 管理ポータルと AppSource による展開は、まだ有効になっていません。
- Excel Onlineでのカスタム関数は、一定期間動作していないと、セッション中に停止することがあります。ブラウザーのページを更新 (F5) し、機能を復元するカスタム関数を再入力します。
- WindowsのExcelで複数のアドイン Excel for Windows で実行されている場合は、 **#GETTING_DATA** の一時的な結果がワークシートのセル内に表示されることがあります。その場合には、Excel のウィンドウをすべて閉じ、Excel を再起動します。
- カスタム関数向けのデバッグ ツールが、将来利用できるようになります。それまでは、F12 開発者ツールを使用してExcel Onlineでデバッグすることができます。詳細は、[カスタム関数のベスト プラクティス](custom-functions-best-practices.md)を参照してください。

## <a name="changelog"></a>変更ログ

- **2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開*
- **2017 年 11 月 20 日**: ビルド 8801 以降を使用しているユーザー向けに互換性バグを修正
- **2017 年 11 月 28 日**: 非同期関数のキャンセルへのサポートを公開* (ストリーム関数への変更が必要)
- **2018 年 5 月 7 日**: Mac、Excel Online、およびインプロセスで実行される同期関数へのサポートを公開*
-  **2018 年 9 月 20日**:  JavaScript ランタイムのカスタム関数へのサポートを公開されました。詳細については、 [「Excel のカスタム関数ランタイム」](custom-functions-runtime.md)をご覧ください。

\*  [Office Insider](https://products.office.com/office-insider) チャネル (旧称 '' Insider Fast ")

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [Excel カスタム関数のチュートリアル](excel-tutorial-custom-functions.md)
