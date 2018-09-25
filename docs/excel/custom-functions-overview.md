---
ms.date: 09/20/2018
description: JavaScript を使用して Excel でカスタム関数を作成します。
title: Excel でのカスタム関数の作成 (プレビュー)
ms.openlocfilehash: b214329fe50955d0f39d50f674152f475ca24b4d
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/25/2018
ms.locfileid: "25005044"
---
# <a name="create-custom-functions-in-excel-preview"></a>Excel でのカスタム関数の作成 (プレビュー)

開発者はカスタム関数を使用すれば、アドインの一部として新しい関数を定義して、これらの関数を追加することができます。 Excel 内のユーザーは、Excel の他のネイティブ関数 (`SUM()` など) と同様に、カスタム関数にアクセスできます。 この記事では、Excel でカスタム関数を作成する方法について説明します。

次の図では、エンド ユーザーが Excel ワークシートのセルにカスタム関数を挿入する例を示します。 カスタム関数は、ユーザーが関数への入力パラメーターとして指定する数値ペアに、42 を足すように設計されています。`CONTOSO.ADD42`

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

次のコードは、`ADD42` カスタム関数を定義します。

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

カスタム機能は、Windows、Mac、および Excel Online の開発者プレビューで利用できるようになりました。 これらを試すには、以下の手順を行います。

1. Office (Windows はビルド 10827、Mac は 13.329) をインストールし、 [Office Insider](https://products.office.com/office-insider) プログラムに参加します。 カスタム関数へのアクセスを取得するには、Office Insider プログラムに参加する必要があります。現時点では、Office Insider プログラムのメンバーでない限り、カスタム関数はすべての Office のビルド間で無効となっています。

2. [Yo Office](https://github.com/OfficeDev/generator-office) を使用して Excel カスタム関数のアドイン プロジェクトを作成し、[OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) の指示に従ってプロジェクトを使用します。

3. Excel ワークシートの任意のセルに `=CONTOSO.ADD42(1,2)` と入力し、**Enter** キーを押してカスタム関数を実行します。

> [!NOTE]
> この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。

## <a name="learn-the-basics"></a>基本操作の説明

[Yo Office](https://github.com/OfficeDev/generator-office) を使用して作成したカスタム関数プロジェクトに、以下のファイルが表示されます。

| ファイル | ファイル形式 | 説明 |
|------|-------------|-------------|
| **./src/customfunctions.js** | JavaScript | カスタム関数を定義するコードが含まれています。 |
| **./config/customfunctions.json** | JSON | カスタム関数について説明し、エンドユーザーが使用可能なように、Excel で関数を登録できるようにするメタデータが含まれています。 |
| **./index.html** | HTML | カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。 |
| **./manifest.xml** | XML | アドイン内のすべてのカスタム関数の名前空間と、このテーブルで前に一覧表示した JavaScript、JSON、HTML ファイルの位置を指定します。 |

### <a name="manifest-file-manifestxml"></a>マニフェスト ファイル (./manifest.xml)

カスタム関数を定義するアドイン用の XML マニフェスト ファイルでは、アドイン内のすべてのカスタム関数の名前空間と、JavaScript、JSON、HTML ファイルの位置を定義します。 次の XML マークアップの例では、Excel がカスタム関数を実行できるようにするための、アドインのマニフェストに含める必要のある `<ExtensionPoint>` および `<Resources>` 要素の例を示します。  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. For example, a function named "ADD42" is invoked as `=CONTOSO.ADD42` in Excel.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> Excel 内の関数は、XML マニフェスト ファイルで指定される名前空間の先頭に追加されます。 関数の名前空間は関数名の前に配置され、それらはピリオドで区切られます。 たとえば、Excel ワークシートのセル内の関数 `ADD42()` を呼び出すには、`=CONTOSO.ADD42` と入力します。これは、CONTOSO が名前空間で、`ADD42` が JSON ファイルで指定された関数の名前であるからです。 名前空間は、所属する会社またはアドインの識別子として使用することを想定しています。 

### <a name="json-file-configcustomfunctionsjson"></a>JSON ファイル (./config/customfunctions.json)

カスタム関数のメタデータ ファイルは、Excel がカスタム関数を登録し、エンドユーザーが使用できるようにするために必要とする情報を提供します。 カスタム関数は、ユーザーがはじめてアドインを実行したときに登録されます。 その後、その同じユーザーは、最初にアドインが実行されたブックだけでなく、すべてのブックでそれらのカスタム関数を使用できるようになります。

> [!TIP]
> JSON ファイルをホストするサーバーのサーバー設定では、カスタム関数が Excel Online で正しく作動するために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効にする必要があります。

以下の **customfunctions.json** のコードでは、この記事で前述した `ADD42` 関数のメタデータを指定します。 このメタデータでは、関数の名前、説明、戻り値、入力パラメーターその他を定義します。 このコード サンプルの次の表では、この JSON オブジェクト内の個々のプロパティについての詳細情報を示しています。

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [                {
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
        }
    ]
}
```

以下の表では、通常 JSON メタデータ ファイルに格納されているプロパティを一覧表示しています。 前の例で使用されていないオプションを含む、JSON メタデータ ファイルの詳細情報については、「[カスタム関数のメタデータ](custom-functions-json.md)」を参照してください。

| プロパティ  | 説明 |
|---------|---------|
| `id` | 関数の一意の ID です。 設定後は、この ID は変更しないでください。 |
| `name` | ユーザーがセルに数式を入力した際に、オートコンプリート メニューに表示される関数の名前です。 オートコンプリート メニューでは、XML マニフェスト ファイルで指定されるカスタム関数の名前空間が、この値に接頭辞としてつきます。 |
| `helpUrl` | ユーザーがヘルプを要求したときに表示されるページの Url です。 |
| `description` | 関数が実行することについて説明します。 この値は、関数が Excel 内のオートコンプリート メニューで選択された項目となっている場合に、ツールヒントとして表示されます。 |
| `result`  | 関数によって返される情報の種類を定義するオブジェクトです。 子プロパティには、**文字列**、**数値**、または**ブール値**を使用できます。`type` `dimensionality` 子プロパティの値には、**スカラー**または**マトリックス** (指定された `type` の値の 2 次元配列) が使用できます。 |
| `parameters` | 関数の入力パラメーターを定義する配列。 `name` および `description` 子プロパティが Excel intelliSense に表示されます。 および `dimensionality` 子プロパティは、この表で前述した `result` オブジェクトの子プロパティと同じです。`type` |
| `options` | Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。 このプロパティの使用方法の詳細については、この記事で後述する「[ストリーム関数](#streamed-functions)」および「[キャンセル](#canceling-a-function)」を参照してください。 |

## <a name="functions-that-return-data-from-external-sources"></a>外部ソースからデータを返す関数

カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。

1. JavaScript Promise を Excel に返します。

2. コールバック関数を使用して Promise を最終値で解決します。

カスタム関数は、Excel が最終結果を待つ間、セルに `#GETTING_DATA` の一時的な結果を表示します。 ユーザーは、カスタム関数が結果を待つ間、ワークシートの他の部分を通常通り操作することができます。

以下のコード サンプルでは、`getTemperature()` カスタム関数が温度計の現在の温度を取得します。 `sendWebRequest` は XHR を使用して温度 Web サービスを呼び出す仮想関数 (ここでは説明していません) であることに注意してください。

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

- 最終的なパラメーター `handler` は登録コードでは指定されず、Excel ユーザーが関数を入力するときにオートコンプリート メニューに表示されません。 これは、関数のデータを Excel に渡してセルの値を更新するために使用される `setResult` コールバック関数を含むオブジェクトです。

- Excel が `handler` オブジェクトの `setResult` 関数を渡すには、関数の登録の際に、JSON メタデータ ファイル内のカスタム関数の `options` プロパティでオプション `"stream": true` を設定して、ストリーミングへのサポートを宣言する必要があります。

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="canceling-a-function"></a>関数をキャンセルする

状況によっては、帯域幅の消費量、作業メモリ、UPC への負荷を減らすために、ストリーム カスタム関数の実行をキャンセルする必要がある場合もあります。 Excel は、以下のような状況では関数の実行をキャンセルします。

- ユーザーが、関数への参照があるセルを編集または削除した場合。

- 関数の引数 (入力) のいずれかが変更された場合。 この場合、キャンセルに続いて新しい関数の呼び出しがトリガーされます。

- ユーザーが手動で再計算をトリガーする。 この場合、キャンセルに続いて新しい関数の呼び出しがトリガーされます。

> [!NOTE]
> すべてのストリーミング関数に対してキャンセル ハンドラを実装することが 必須 です。

関数をキャンセル可能にするには、JSON メタデータ ファイルのカスタム関数の `options` プロパティで、オプション `"cancelable": true` を設定してください。

以下のコードは、前述したのと同じ `incrementValue` 関数を示していますが、今回はキャンセル ハンドラが実装されています。 この例では、`incrementValue` 関数がキャンセルされたときに `clearInterval()` が実行されます。

```js
function incrementValue(increment, handler){
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);

    handler.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a>状態の保存と共有

カスタム関数では、JavaScript のグローバル変数にデータを保存できます。 後続の呼び出しでは、カスタム関数はこれらの変数に保存されている値を使用できます。 保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を追加する場合に便利です。 たとえば、同じ Web リソースへの追加呼び出しを避けるために、呼び出しから返されたデータを Web リソースに保存することができます。

次のコード サンプルは、 状態をグローバルで保存する前述の温度ストリーミング関数の実装を示しています。 このコードについては、次の点に注意してください。

- `refreshTemperature` は、1 秒おきに特定の温度計の温度を読み取るストリーム関数です。 新しい温度は `savedTemperatures` 変数に保存されますが、セルの値を直接更新することはありません。 ワークシート・セルから直接呼び出されません。*したがって、JSON ファイルには登録されません *

- `streamTemperature` 1 秒おきにセルに表示される温度値を更新します。また、 `savedTemperatures` 変数をデータソースとして使用します。 JSON ファイルに登録し、すべて大文字で `STREAMTEMPERATURE` という名前をつける必要があります。

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
    let url = "https://yourhypotheticalapi/comments/" + x;

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
- カスタム機能は現在、モバイル クライアント用の Excel では使用できません。
- 揮発性関数（スプレッドシート内の無関係なデータが変更されたときに自動的に再計算する関数）はまだサポートされていません。
- Office 365 管理ポータルと AppSource による展開はまだ有効になっていません。
- Excel Online のカスタム関数は、一定期間使用しないとセッション中に機能しなくなることがあります。 ブラウザページを更新（F5）し、カスタム関数を再入力して機能を復元します。
- Excel for Windows で実行されている複数のアドインがある場合には、ワークシートのセル内に **#GETTING_DATA** の一時的な結果が表示される場合があります。 すべての Excel ウィンドウを閉じて、Excel を再起動します。
- 将来的には、カスタム関数用のデバッグ ツールが利用可能となる可能性があります。 それまでは、F12 開発者ツールを使用して Excel オンラインでデバッグできます。 詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md)」を参照してください。

## <a name="changelog"></a>変更ログ

- **2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開*
- **2017 年 11 月 20 日**: ビルド 8801 以降を使用する場合の互換性バグを修正
- **2017 年 11 月 28 日**: 非同期関数のキャンセルへのサポートを公開* (ストリーム関数への変更が必要)
- **2018 年 5 月 7 日**: Mac、Excel Online、およびインプロセスで実行される同期関数へのサポートを公開*
- **2018 年 9 月 20日**: JavaScript 実行時のカスタム関数へのサポートを公開 詳細については、「[Excel カスタム関数のランタイム](custom-functions-runtime.md)」を参照してください。

\* Office Insiders チャネル対象

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
