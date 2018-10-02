---
ms.date: 09/27/2018
description: JavaScript を使用して Excel でカスタム関数を作成します。
title: Excel でカスタム関数を作成する (プレビュー)
ms.openlocfilehash: 98e418f843f6f5574088cea9c7393afc4a42060b
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348802"
---
# <a name="create-custom-functions-in-excel-preview"></a>Excel でカスタム関数を作成する (プレビュー)

JavaScript で関数をアドインの一部として定義することにより、開発者はカスタム関数を使用して Excel に新しい関数を追加することができます。Excel 内のユーザーは、`SUM()` などの Excel のネイティブ関数にアクセスするのと同様に、カスタム関数にアクセスできます。この記事では、Excel でカスタム関数を作成する方法について説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

次の図では、エンド ユーザーが Excel ワークシートのセルにカスタム関数を挿入する例を示します。 `CONTOSO.ADD42` カスタム関数は、ユーザーが関数への入力パラメーターとして指定する数値ペアに、42 を足すように設計されています。

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

次のコードは、`ADD42` カスタム関数を定義します。

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。

## <a name="components-of-a-custom-functions-add-in-project"></a>カスタム関数アドイン プロジェクトのコンポーネント

[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Excel カスタム関数アドイン プロジェクトを作成する場合は、ジェネレーターが作成するプロジェクトに以下のようなファイルが表示されます。

| ファイル | ファイル形式 | 説明 |
|------|-------------|-------------|
| **./src/customfunctions.js**<br/>または<br/>**./src/customfunctions.ts** | JavaScript<br/>または<br/>TypeScript | カスタム関数を定義するコードが含まれています。 |
| **./config/customfunctions.json** | JSON | カスタム関数を説明するメタデータが含まれており、Excel でカスタム関数を登録してエンドユーザーが使用できるようにします。 |
| **./index.html** | HTML | カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。 |
| **./manifest.xml** | XML | アドイン内のすべてのカスタム関数の名前空間と、このテーブルで前に一覧表示した JavaScript、JSON、HTML ファイルの位置を指定します。 |

次のセクションでは、これらのファイルの詳細についてを説明します。

### <a name="script-file"></a>スクリプト ファイル 

スクリプト ファイル (Yo Office ジェネレーターが作成するプロジェクト内の **./src/customfunctions.js** または **/src/customfunctions.ts**) には、カスタム関数を定義して、カスタム関数の名前を [JSON メタデータ ファイル](#json-metadata-file)のオブジェクトにマップするコードが含まれています。 

たとえば、以下のコードでは、`add` と `increment` というカスタム関数を定義して、次に両方の関数のマッピング情報を指定します。 `add` 関数は、`id` プロパティの値が **ADD** である JSON メタデータ ファイルのオブジェクトにマップされ、`increment` 関数は、`id` プロパティの値が **INCREMENT** であるメタデータ ファイルのオブジェクトにマップされます。 スクリプト ファイルの関数名を JSON メタデータ ファイルのオブジェクトにマップする方法の詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)」を参照してください。

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

カスタム関数のメタデータ ファイル (Yo Office ジェネレーターが作成するプロジェクト内の **./config/customfunctions.json**) は、Excel でカスタム関数を登録してエンドユーザーが使用できるようにするのに必要な情報を示しています。 カスタム関数は、ユーザーがはじめてアドインを実行したときに登録されます。 その後、その同じユーザーは、最初にアドインが実行されたブックだけでなく、すべてのブックでそれらのカスタム関数を使用できるようになります。

> [!TIP]
> JSON ファイルをホストするサーバーのサーバー設定では、カスタム関数が Excel Online で正しく作動するために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効にする必要があります。

以下の **customfunctions.json** のコードでは、この記事で前述した `add` 関数と `increment` 関数のメタデータを指定します。 このコード サンプルの次の表では、この JSON オブジェクト内の個々のプロパティについての詳細情報を示しています。 JSON メタデータ ファイルの `id` および `name` プロパティの値を指定する方法の詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)」を参照してください。

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

以下の表では、通常 JSON メタデータ ファイルに格納されているプロパティを一覧表示しています。 JSON メタデータ ファイルの詳細については、「[カスタム関数のメタデータ](custom-functions-json.md)」を参照してください。

| プロパティ  | 説明 |
|---------|---------|
| `id` | 関数の一意の ID です。 設定後は、この ID は変更しないでください。 |
| `name` | Excel でエンド ユーザーに対して表示される関数の名前です。 Excel では、 [XML マニフェスト ファイル](#manifest-file)で指定されているカスタム関数の名前空間が、関数名に接頭辞として付きます。 |
| `helpUrl` | ユーザーがヘルプを要求したときに表示されるページの URL です。 |
| `description` | 関数が実行することについて説明します。 この値は、関数が Excel 内のオートコンプリート メニューで選択された項目となっている場合に、ツールヒントとして表示されます。 |
| `result`  | 関数によって返される情報の種類を定義するオブジェクトです。 `type` 子プロパティには、**文字列**、**数値**、または**ブール値**を使用できます。 `dimensionality` 子プロパティの値には、**スカラー**または**マトリックス** (指定された `type` の値の 2 次元配列) が使用できます。 |
| `parameters` | 関数の入力パラメーターを定義する配列。 `name` および `description` 子プロパティが Excel intelliSense に表示されます。 `type` 子プロパティ値には、**文字列**、**数値**、または**ブール値**を使用できます。 `dimensionality` 子プロパティの値には、**スカラー**または **マトリックス** (指定された `type` の値の 2 次元配列) が使用できます。 |
| `options` | Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。 このプロパティの使用方法の詳細については、この記事で後述する「[ストリーム関数](#streamed-functions)」および「[関数のキャンセル](#canceling-a-function)」を参照してください。 |

### <a name="manifest-file"></a>マニフェスト ファイル

カスタム関数 (Yo Office ジェネレーターが作成するプロジェクト内の **./manifest.xml**) は、アドイン内のすべてのカスタム関数の名前空間と、JavaScript、JSON、HTML ファイルの位置を指定します。 以下の XML マークアップでは、カスタム関数を有効にするためにアドインのマニフェストに含める必要のある `<ExtensionPoint>` および `<Resources>` 要素の例を示します。  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. -->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> Excel 内の関数の先頭には、XML マニフェスト ファイルで指定される名前空間が追加されます。 関数の名前空間は関数名の前に配置され、それらはピリオドで区切られます。 たとえば、Excel ワークシートのセル内の関数 `ADD42` を呼び出すには、`=CONTOSO.ADD42` と入力します。これは、CONTOSO が名前空間で、`ADD42` が JSON ファイルで指定された関数の名前であるからです。 名前空間は、所属する会社またはアドインの識別子として使用することを想定しています。 

## <a name="functions-that-return-data-from-external-sources"></a>外部ソースからデータを返す関数

カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。

1. JavaScript Promise を Excel に返す。

2. コールバック関数を使用して Promise を最終値で解決します。

カスタム関数は、Excel が最終結果を待つ間、セルに `#GETTING_DATA` の一時的な結果を表示します。 ユーザーは、カスタム関数が結果を待つ間、ワークシートの他の部分を通常通り操作することができます。

以下のコード サンプルでは、`getTemperature()` カスタム関数が温度計の現在の温度を取得します。 `sendWebRequest` は [XHR](custom-functions-runtime.md#xhr) を使用して温度 Web サービスを呼び出す仮想関数 (ここでは説明していません) であることに注意してください。

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a>ストリーム関数

ストリーム カスタム関数を使用すると、時間の経過とともにセルに繰り返しデータを出力でき、ユーザーが再計算を要求することは特に必要ありません。 以下のコード サンプルは、1 秒おきに結果に数値を追加するカスタム関数です。 このコードについては、次の点に注意してください。

- Excel は、`setResult`コールバックを使用して自動的に新しい値を表示します。

- 2 番目のパラメーター `handler` は、[オートコンプリート] メニューから関数を選択する場合には、エンドユーザーに対して表示されません。

- `onCanceled` コールバックは、関数がキャンセルされた場合に実行される関数を定義します。 すべてのストリーム関数に対して、このようなキャンセル ハンドラーを実装する必要があります。 詳細については、 「[関数のキャンセル](#canceling-a-function)」を参照してください。 

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

JSON メタデータ ファイルでストリーム関数にメタデータを指定する場合には、以下の例に示すように、`options` オブジェクトにプロパティ `"cancelable": true` および `"stream": true` を設定する必要があります。

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

## <a name="canceling-a-function"></a>関数のキャンセル

状況によっては、帯域幅の消費量、作業メモリ、UPC への負荷を減らすために、ストリーム カスタム関数の実行をキャンセルする必要がある場合もあります。 Excel は、以下のような状況では関数の実行をキャンセルします。

- ユーザーが、関数への参照があるセルを編集または削除した場合。

- 関数の引数 (入力) のいずれかが変更された場合。 この場合、キャンセルに続いて新しい関数の呼び出しがトリガーされます。

- ユーザーが手動で再計算をトリガーした場合。 この場合、キャンセルに続いて新しい関数の呼び出しがトリガーされます。

関数をキャンセルする機能を有効にするには、JavaScript 関数内にキャンセル ハンドラーを実装して、関数を記述する JSON メタデータの `options` オブジェクト内にプロパティ `"cancelable": true` を指定する必要があります。 この記事の前のセクションのコード サンプルは、これらの手法の例を示しています。

## <a name="saving-and-sharing-state"></a>状態の保存と共有

カスタム関数では、JavaScript のグローバル変数にデータを保存できます。 後続の呼び出しでは、カスタム関数はこれらの変数に保存されている値を使用できます。 保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を追加する場合に便利です。 たとえば、Web リソースへの呼び出しから返されたデータを保存しておけば、同じ Web リソースへ繰り返し呼び出しを行わなくて済みます。

以下のコード サンプルは、 状態をグローバルで保存する温度ストリーミング関数の実装を示しています。 このコードについては、次の点に注意してください。

- `refreshTemperature` は、1 秒おきに特定の温度計の温度を読み取るストリーム関数です。 新しい温度は `savedTemperatures` 変数に保存されますが、セルの値を直接更新することはありません。 ワークシート・セルから直接呼び出されません。*したがって、JSON ファイルには登録されません *

- `streamTemperature` 1 秒おきにセルに表示される温度値を更新します。また、 `savedTemperatures` 変数をデータ ソースとして使用します。 JSON ファイルに登録し、すべて大文字で `STREAMTEMPERATURE` という名前をつける必要があります。

- ユーザーは、Excel UI の複数のセルから `streamTemperature` を呼び出すことができます。 呼び出すたびに、同じ `savedTemperatures` 変数からデータを読み取ります。

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
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

## <a name="working-with-ranges-of-data"></a>データの範囲を使用する

カスタム関数は、入力パラメーターとしてデータの範囲を受け取ることができます。または、データの範囲を返すこともできます。 JavaScript では、データの範囲は、2 次元配列として表されます。

たとえば、関数が Excel に格納されている数値の範囲から 2 番目に高い値を返すとします。 以下の関数が、タイプ `Excel.CustomFunctionDimensionality.matrix` のものである `values` パラメーターを受け取ります。 この関数の JSON メタデータでは、パラメーターの `type` プロパティを `matrix` に設定するように注意してください。

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

カスタム関数を定義するアドインをビルドする場合には、実行時エラーに対処するエラー処理ロジックを含めるようにしてください。 カスタム関数のエラー処理は、[一般的な Excel JavaScript API のエラー処理](excel-add-ins-error-handling.md) と同じです。 以下のコード サンプルでは、`.catch` がコード内で発生するエラーを処理します。

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
- Office 365 管理ポータルと AppSource による展開はまだ有効になっていません。
- Excel Online のカスタム関数は、一定期間使用しないとセッション中に機能しなくなることがあります。 ブラウザページを更新（F5）し、カスタム関数を再入力して機能を復元します。
- Excel for Windows で実行されている複数のアドインがある場合には、ワークシートのセル内に **#GETTING_DATA** の一時的な結果が表示される場合があります。 すべての Excel ウィンドウを閉じて、Excel を再起動します。
- 将来的には、カスタム関数用のデバッグ ツールが利用可能となる可能性があります。 それまでは、F12 開発者ツールを使用して Excel オンラインでデバッグできます。 詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md)」を参照してください。

## <a name="changelog"></a>変更ログ

- **2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開*
- **2017 年 11 月 20 日**: ビルド 8801 以降を使用しているユーザー向けに互換性バグを修正
- **2017 年 11 月 28 日**: 非同期関数のキャンセルへのサポートを公開* (ストリーム関数への変更が必要)
- **2018 年 5 月 7 日**: Mac、Excel Online、およびインプロセスで実行される同期関数へのサポートを公開*
- **2018 年 9 月 20日**: JavaScript 実行時のカスタム関数へのサポートを公開 詳細については、「[Excel カスタム関数のランタイム](custom-functions-runtime.md)」を参照してください。

\* Office Insiders チャネル対象

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [Excel カスタム関数のチュートリアル](excel-tutorial-custom-functions.md)