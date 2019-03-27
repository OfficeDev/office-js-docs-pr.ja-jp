---
ms.date: 03/19/2019
description: JavaScript を使用して Excel でカスタム関数を作成する。
title: Excel でのカスタム関数の作成 (プレビュー)
localization_priority: Priority
ms.openlocfilehash: ac3410267da415c4d567092da2e653fcffd10b72
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870451"
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
| **./src/customfunctions.js**<br/>または<br/>**./src/customfunctions.ts** | JavaScript<br/>または<br/>TypeScript | カスタム関数を定義するコードが含みます。 |
| **./config/customfunctions.json** | JSON | カスタム関数を定義し、Excel に関数を登録してエンドユーザーが使用できるようにするためのメタデータを含みます。 |
| **./index.html** | HTML | カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。 |
| **./manifest.xml** | XML | アドイン内のすべてのカスタム関数の名前空間と、この表で前述した JavaScript、JSON、HTML ファイルの位置を指定します。 |

次のセクションでは、これらのファイルに関する詳細について説明します。

### <a name="script-file"></a>スクリプト ファイル

スクリプト ファイル (Yo Office ジェネレーターが作成するプロジェクト内の **./src/customfunctions.js** または **./src/customfunctions.ts**) には、カスタム関数を定義して、カスタム関数の名前を [JSON メタデータ ファイル](#json-metadata-file)のオブジェクトにマップするコードが含まれています。 

たとえば、次のコードはカスタム関数 `add` と `increment` を定義し、両方の関数の関連付け情報を指定します。 `add` 関数は、`id` プロパティの値が **ADD** の JSON メタデータ ファイル内のオブジェクトに関連付けられ、`increment` 関数は、`id` プロパティの値が **INCREMENT** のメタデータ ファイル内のオブジェクトに関連付けられます。 JSON メタデータ ファイル内のオブジェクトへのスクリプト ファイル内関数名の関連付けの詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#associating-function-names-with-json-metadata)」を参照してください。

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

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
 CustomFunctions.associate("INCREMENT", increment);
```

### <a name="json-metadata-file"></a>JSON メタデータ ファイル

カスタム関数のメタデータ ファイル (Yo Office ジェネレーターが作成するプロジェクトでは **./config/customfunctions.json**) は、Excel がカスタム関数の登録し、エンドユーザーが利用できるようするために必要な情報を提供します。 カスタム関数は、ユーザーがアドインを初めて実行するときに登録されます。 その後は、同じユーザーに対しては、(アドインが最初に実行されたワークブック内のみでなく) すべてのワークブック内で利用が可能になります。

> [!TIP]
> JSON ファイルをホストするサーバーでは、カスタム関数を Excel Online で正しく作動させるために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効に設定する必要があります。

**customfunctions.json** の次のコードは、`add` 関数のメタデータと上述の `increment` 関数を指定します。 このコード サンプルに続く表では、JSON オブジェクト内の個別のプロパティについての詳細情報を提供します。 JSON メタデータ ファイル内の `id` と `name` 各プロパティーの値の指定に関する詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#associating-function-names-with-json-metadata)」を参照してください。

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
| `options` | Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 このプロパティの使用方法の詳細については、「[ストリーム関数](#streaming-functions)」および「[関数のキャンセル](#canceling-a-function)」を参照してください。 |

### <a name="manifest-file"></a>マニフェスト ファイル

カスタム関数 (Yo Office ジェネレーターが作成するプロジェクトでは **./manifest.xml**) を定義するアドインの XML マニフェスト ファイルは、アドイン内のすべてのカスタム関数の名前空間と、 JavaScript、JSON、および HTML の場所を指定します。 次の XML マークアップでは、`<ExtensionPoint>` と `<Resources>` カスタム関数を有効にするアドインのマニフェストに含める必要がある要素の例を示します。  

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://localhost:8081/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://localhost:8081/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://localhost:8081/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> Excel の関数は、XML マニフェスト ファイルで指定された名前空間が接頭辞として付加されます。 関数の名前空間は、関数名の前に付けられ、ピリオドで区切られます。 例えば、Excel ワークシートのセル内で、`ADD42` 関数を呼び出すためには、`=CONTOSO.ADD42` と入力します。これは、`CONTOSO` が名前空間で、`ADD42` が JSON ファイルで指定された関数の名前だからです。 名前空間は、会社またはアドインの識別子としての使用を目的としています。 名前空間にはアルファベットとピリオドのみを含めることが出来ます。

## <a name="declaring-a-volatile-function"></a>揮発性関数の宣言

[揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)とは、関数のいずれの引数にも変更がない場合でも、値が刻々と変化する関数のことです。 これらの関数は、Excel が再計算するたびに再計算を行います。 たとえば、`NOW` 関数を呼び出すセルがあるとします。 `NOW` が呼び出される度に、現在の日付と時刻を自動的に返します。

Excel には、`RAND` や `TODAY` などの組み込み揮発性関数がいくつか含まれています。 Excel の揮発性関数の完全なリストは、「[揮発性および非揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)」を参照してください。

カスタム関数を使用すると独自の揮発性関数を作成することができ、日時、時間、乱数、およびモデルを処理するときに役立つ場合があります。 たとえば、モンテカルロ シミュレーションでは、最適なソリューションを決定するにはランダムな入力値の生成が必要です。

関数を揮発性であると宣言するには、次のコードで示されるように、JSON メタデータファイルの関数で、`options` オブジェクトに`"volatile": true` を追加します。 関数で `"streaming": true`と`"volatile": true` の両方をマークすることはできません。両方とも `true` とマークされている場合、揮発性のオプションは無視されます。

```json
{
 "id": "TOMORROW",
  "name": "TOMORROW",
  "description":  "Returns tomorrow’s date",
  "helpUrl": "http://www.contoso.com",
  "result": {
      "type": "string",
      "dimensionality": "scalar"
  },
  "options": {
      "volatile": true
  }
}
```

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

## <a name="coauthoring"></a>共同編集

Excel Online と Excel for Windows で Office 365 サブスクリプションを利用している場合、ドキュメントの共同編集を行うことができ、カスタム関数を使用できます。 ブックでカスタム関数を使用している場合、仕事仲間はカスタム関数のアドインを読み込むように要求されます。 双方がアドインを読み込むと、共同編集によりカスタム関数は結果を共有します。

共同編集の詳細については、「[Excel での共同編集](/office/vba/excel/concepts/about-coauthoring-in-excel)」を参照してください。

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

## <a name="determine-which-cell-invoked-your-custom-function"></a>カスタム関数が呼び出したセルを特定する

場合によっては、カスタム関数が呼び出したセルのアドレスを取得する必要が生じます。 これは、次の種類のシナリオで役立ちます。

- 範囲の書式設定: [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data) で情報を格納するキーとしてセル アドレスを使用します。 Excel で [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) を使用して`AsyncStorage` からキーを読み込みます。
- キャッシュされた値を表示させる: 関数がオフラインで使用される場合、`onCalculated` を使用して `AsyncStorage` に格納されているキャッシュされた値を表示します。
- 調整: セル アドレスを使用して元のセルを検出し、処理が発生している場所での調整を行えます。

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

## <a name="known-issues"></a>既知の問題

既知の問題については、[Excel カスタム関数についての GitHub のレポート](https://github.com/OfficeDev/Excel-Custom-Functions/issues)を参照してください。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [カスタム関数の変更ログ](custom-functions-changelog.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
