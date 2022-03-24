---
title: カスタム関数の JSON メタデータを手動で作成Excel
description: カスタム関数の JSON メタデータを定義し、Excel ID と名前のプロパティを関連付ける。
ms.date: 12/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: 0ef786be050f081c7186f157b6add7a4aca5b04a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744554"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a>カスタム関数の JSON メタデータを手動で作成する

カスタム関数の概要[](custom-functions-overview.md)に関する記事で説明したように、カスタム関数プロジェクトには、JSON メタデータ ファイルとスクリプト (JavaScript または TypeScript) ファイルの両方を含め、関数を登録して使用できる必要があります。 カスタム関数は、ユーザーが初めてアドインを実行した後、すべてのブックで同じユーザーが使用できる場合に登録されます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

独自の JSON ファイルを作成する代わりに、可能な場合は JSON 自動生成を使用することをお勧めします。 自動生成はユーザー エラーの発生が少なく、スキャ `yo office` フォールディング されたファイルには既にこれが含まれます。 JSDoc タグと JSON 自動生成プロセスの詳細については、「カスタム関数の JSON メタデータを自動生成 [する」を参照してください](custom-functions-json-autogeneration.md)。

ただし、カスタム関数プロジェクトを最初から作成できます。 このプロセスでは、次の操作を行う必要があります。

- JSON ファイルを書き込む。
- マニフェスト ファイルが JSON ファイルに接続されていることを確認します。
- 関数を登録するために、 `id` `name` スクリプト ファイル内の関数とプロパティを関連付ける。

次の図は、スキャフォールディング ファイルの使用と JSON の `yo office` 最初からの書き込みとの違いを説明しています。

![アドインに Yeoman ジェネレーターを使用して、Office JSON を記述する場合の違いをイメージします。](../images/custom-functions-json.png)

> [!NOTE]
> アドインに [Yeoman](../develop/yeoman-generator-overview.md) ジェネレーターを使用しない場合は、XML マニフェスト ファイルのセクションを通じて、作成した JSON `<Resources>` ファイルにマニフェストを接続Officeしてください。

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>メタデータの作成とマニフェストへの接続

プロジェクトに JSON ファイルを作成し、関数のパラメーターなど、プロジェクト内の関数に関する詳細を提供します。 関数プロパティ[の完全な一覧](#json-metadata-example)[については、](#metadata-reference)次のメタデータ例とメタデータ参照を参照してください。

次の例と同様に、XML マニフェスト ファイルでセクション内の JSON `<Resources>` ファイルを参照してください。

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## <a name="json-metadata-example"></a>JSON メタデータの例

次の例では、カスタム関数を定義するアドインの JSON メタデータ ファイルの内容を示します。 この例の後に続くセクションでは、JSON の例に含まれる個々のプロパティの詳細について説明します。

```json
{
  "allowCustomDataForDataTypeAny": true, // This property is currently only available in public preview.
  "allowErrorForDataTypeAny": true,
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com/help",
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
      "id": "GETDAY",
      "name": "GETDAY",
      "description": "Get the day of the week",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE",
      "description": "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "increment",
          "description": "the number to be added each time",
          "type": "number",
          "dimensionality": "scalar"
        }
      ],
      "options": {
        "stream": true,
        "cancelable": true
      }
    },
    {
      "id": "SECONDHIGHEST",
      "name": "SECONDHIGHEST",
      "description": "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "range",
          "description": "the input range",
          "type": "number",
          "dimensionality": "matrix"
        }
      ]
    }
  ]
}
```

> [!NOTE]
> 完全なサンプル JSON ファイルは、リポジトリのコミット履歴Excel [OfficeDev/GitHub-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) で使用できます。 プロジェクトが JSON を自動的に生成するために調整されたので、手書き JSON の完全なサンプルは、以前のバージョンのプロジェクトでのみ使用できます。

## <a name="metadata-reference"></a>メタデータ参照

### <a name="allowcustomdatafordatatypeany-preview"></a>allowCustomDataForDataTypeAny (プレビュー)

> [!NOTE]
> プロパティ`allowCustomDataForDataTypeAny`は現在パブリック プレビューで使用できます。このプロパティは、パブリック プレビュー OfficeとWindows。 プレビュー機能は変更されることがあります。運用環境での使用は意図されていません。 試用はテスト環境と開発環境に限定することをお勧めします。 運用環境やビジネス上重要なドキュメントでプレビュー機能を使用しないでください。
>
> Windows の Office でこのプロパティを試すには、Excel ビルド番号が 16.0.14623.20002 以上である必要があります。 この機能を使用するには、[Office Insider プログラム](https://insider.office.com/)に参加して、**ベータ チャネル** Insider レベルを選択する必要があります。 詳細については、「[Office Insider プログラムに参加する](https://insider.office.com/join/windows)」を参照してください。

プロパティ `allowCustomDataForDataTypeAny` はブール型です。 この値を設定すると、 `true` カスタム関数はデータ型をパラメーターとして受け入れ、値を返します。 詳細については、「カスタム関数 [とデータ型」を参照してください](custom-functions-data-types-concepts.md)。

他のほとんどの JSON メタデータ プロパティとは異なり、 `allowCustomDataForDataTypeAny` トップ レベルのプロパティであり、サブプロパティは含まれています。 このプロパティを書式設定する方法 [の例については](#json-metadata-example) 、上記の JSON メタデータ コード サンプルを参照してください。

### <a name="allowerrorfordatatypeany"></a>allowErrorForDataTypeAny

プロパティ `allowErrorForDataTypeAny` はブール型です。 値を設定すると、カスタム `true` 関数は入力値としてエラーを処理できます。 に設定されている場合、型を持`any``any[][]`つすべてのパラメーターまたは入力値としてエラーを受け`allowErrorForDataTypeAny`入れることができます`true`。 既定値は `allowErrorForDataTypeAny` です `false`。

> [!NOTE]
> 他の JSON メタデータ プロパティとは異なり、 `allowErrorForDataTypeAny` トップ レベルのプロパティであり、サブプロパティは含まれています。 このプロパティを書式設定する方法 [の例については](#json-metadata-example) 、上記の JSON メタデータ コード サンプルを参照してください。

### <a name="functions"></a>functions

`functions` プロパティは、カスタム関数オブジェクトの配列です。 次の表に、各オブジェクトのプロパティを示します。

| プロパティ      | データ型 | 必須 | 説明                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | いいえ       | Excel でエンド ユーザーに表示される関数の説明です。 たとえば、「**華氏の値を摂氏に変換する**」です。                                                            |
| `helpUrl`     | 文字列    | いいえ       | 関数に関する情報を提供する URL です  (作業ウィンドウに表示されます)。たとえば、`http://contoso.com/help/convertcelsiustofahrenheit.html` です。                      |
| `id`          | 文字列    | はい      | 関数の一意の ID です。 この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。                                            |
| `name`        | 文字列    | はい      | Excel でエンド ユーザーに表示される関数の名前です。 このExcel、この関数名には、XML マニフェスト ファイルで指定されたカスタム関数名前空間のプレフィックスが付けられている。 |
| `options`     | オブジェクト    | いいえ       | Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 詳細については、[options](#options) に関する説明を参照してください。                                                          |
| `parameters`  | 配列     | はい      | 関数の入力パラメーターを定義する配列です。 詳細については [、パラメーター](#parameters) を参照してください。                                                                             |
| `result`      | object    | はい      | 関数が返す情報の種類を定義するオブジェクトです。 詳細については、[result](#result) に関する説明を参照してください。                                                                 |

### <a name="options"></a>options

`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 次の表に、`options` オブジェクトのプロパティを示します。

| プロパティ          | データ型 | 必須                               | 説明 |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | ブール   | いいえ<br/><br/>既定値は、`false` です。  | `true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `CancelableInvocation` ハンドラーが呼び出されます。 キャンセル可能な関数は、通常、1 つの結果を返す非同期関数でのみ使用され、データ要求の取り消しを処理する必要があります。 関数は、プロパティとプロパティの両方を `stream` 使用 `cancelable` することはできません。 |
| `requiresAddress` | ブール   | いいえ <br/><br/>既定値は、`false` です。 | 場合 `true`は、カスタム関数は、それを呼び出したセルのアドレスにアクセスできます。 呼 `address` び出しパラメーター [のプロパティには](custom-functions-parameter-options.md#invocation-parameter) 、カスタム関数を呼び出したセルのアドレスが含まれます。 関数は、プロパティとプロパティの両方を `stream` 使用 `requiresAddress` することはできません。 |
| `requiresParameterAddresses` | ブール   | いいえ <br/><br/>既定値は、`false` です。 | 場合 `true`は、カスタム関数は、関数の入力パラメーターのアドレスにアクセスできます。 このプロパティは、result オブジェクトのプロパティと`dimensionality`組み合[](#result) `dimensionality` `matrix`わせて使用する必要があります。 詳細 [については、「パラメーターのアドレスを検出する](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) 」を参照してください。 |
| `stream`          | ブール   | いいえ<br/><br/>既定値は、`false` です。  | `true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。 このオプションは、株価などの急速に変化するデータ ソースに便利です。 この関数には、`return` ステートメントは含めないようにする必要があります。 代わりに、結果の値は `StreamingInvocation.setResult` コールバック メソッドの引数として渡されます。 詳細については、「ストリーミング機能 [を作成する」を参照してください](custom-functions-web-reqs.md#make-a-streaming-function)。 |
| `volatile`        | ブール   | いいえ <br/><br/>既定値は、`false` です。 | 場合`true`は、数式の依存値が変更Excelではなく、計算が再計算されるごとに関数が再計算されます。 関数は、プロパティとプロパティの両方を `stream` 使用 `volatile` することはできません。 プロパティと `stream` プロパティの `volatile` 両方がに設定されている場合 `true`、揮発性プロパティは無視されます。 |

### <a name="parameters"></a>parameters

`parameters` プロパティは、パラメーター オブジェクトの配列です。 次の表に、各オブジェクトのプロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  いいえ |  パラメーターの説明です。 これは、ユーザーのExcelに表示IntelliSense。  |
|  `dimensionality`  |  文字列  |  いいえ  |  (配列以外 `scalar` の値) または (2 次元配列) `matrix` のいずれかである必要があります。  |
|  `name`  |  文字列  |  はい  |  パラメーターの名前です。 この名前は、ExcelのIntelliSense。  |
|  `type`  |  文字列  |  いいえ  |  パラメーターのデータ型です。 、、`boolean`、`number``string`、を使用`any`すると、前の 3 つの種類の任意のを使用できます。 このプロパティを指定しない場合、データ型の既定値は `any`. |
|  `optional`  | ブール | いいえ | `true` の場合、パラメーターは省略可能です。 |
|`repeating`| ブール | いいえ | If `true`, parameters populate from a specified array. 関数のすべての繰り返しパラメーターは、定義によって省略可能なパラメーターと見なされます。  |

### <a name="result"></a>result

`result` オブジェクトは、この関数が返す情報の種類を定義します。 次の表に、`result` オブジェクトのプロパティを示します。

| プロパティ         | データ型 | 必須 | 説明                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | いいえ       | (配列以外 `scalar` の値) または (2 次元配列) `matrix` のいずれかである必要があります。 |
| `type` | 文字列    | いいえ       | 結果のデータ型。 、、`boolean`、 `number``string``any` または (これは、前の 3 つの種類の任意の使用を可能にする) を指定できます。 このプロパティを指定しない場合、データ型の既定値は `any`. |

## <a name="associating-function-names-with-json-metadata"></a>関数名を JSON メタデータに関連付ける

関数が正しく動作するには、関数 `id` のプロパティを JavaScript 実装に関連付ける必要があります。 関連付けがある場合は、関数が登録されないので、この関数を使用Excel。 次のコード サンプルは、メソッドを使用して関連付けを行う方法を示 `CustomFunctions.associate()` しています。 このサンプルではカスタム関数 `add` を定義し、それを `id` プロパティ値が **ADD** の、JSON メタデータ ファイル内のオブジェクトに関連付けます。

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

次の JSON は、前のカスタム関数 JavaScript コードに関連付けられている JSON メタデータを示しています。

```json
{
  "functions": [
    {
      "description": "Add two numbers",
      "id": "ADD",
      "name": "ADD",
      "parameters": [
        {
          "description": "First number",
          "name": "first",
          "type": "number"
        },
        {
          "description": "Second number",
          "name": "second",
          "type": "number"
        }
      ],
      "result": {
        "type": "number"
      }
    }
  ]
}
```

JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。

- JSON のメタデータ ファイルにそれぞれの `id` プロパティには、英数字とピリオドのみが含まれています。

- JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。 すなわち、メタデータ ファイル内の 2 つの関数オブジェクトは同じ `id` 値であってはいけません。

- 対応する JavaScript 関数の名前に関連付けられた後では、JSON のメタデータ ファイル内の `id` プロパティの値を変更しないでください。 JSON のメタデータ ファイル内の `name` プロパティを更新することによって Excel でエンド ユーザーに表示される関数の名前を変更することができます。しかし、確立された後は、 `id` プロパティの値を決して変更しないでください。

- JavaScript ファイルで、各関数の後に使用するカスタム関数の関連 `CustomFunctions.associate` 付けを指定します。

次のサンプルは、前の JavaScript コード サンプルで定義されている関数に対応する JSON メタデータを示しています。 プロパティ `id` の `name` 値は大文字で、カスタム関数を記述する場合のベスト プラクティスです。 独自の JSON ファイルを手動で準備し、自動生成を使用しない場合にのみ、この JSON を追加する必要があります。 自動生成の詳細については、「カスタム関数の [JSON メタデータの自動生成」を参照してください](custom-functions-json-autogeneration.md)。

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="next-steps"></a>次の手順

関数に[名前を付けるベスト](custom-functions-naming.md) プラクティスを説明するか、[](custom-functions-localize.md)前に説明した手書き JSON メソッドを使用して関数をローカライズする方法について説明します。

## <a name="see-also"></a>関連項目

- [カスタム関数用の JSON メタデータの自動生成](custom-functions-json-autogeneration.md)
- [カスタム関数パラメーター のオプション](custom-functions-parameter-options.md)
- [Excel でカスタム関数を作成する](custom-functions-overview.md)
