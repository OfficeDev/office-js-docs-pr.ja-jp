---
ms.date: 05/06/2020
description: Excel でカスタム関数の JSON メタデータを定義し、関数 id と name プロパティを関連付けます。
title: Excel のカスタム関数のメタデータ
localization_priority: Normal
ms.openlocfilehash: 2f044db54d795e221d4b69eb054e639a7d220c10
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609304"
---
# <a name="custom-functions-metadata"></a>カスタム関数のメタデータ

カスタム関数の[概要](custom-functions-overview.md)の記事で説明されているように、カスタム関数プロジェクトには、JSON メタデータファイルとスクリプト (JavaScript または TypeScript) の両方を含める必要があります。関数を登録するには、このファイルを使用できるようにします。 ユーザーが初めてアドインを実行したときに、すべてのブックの同じユーザーがそのアドインを使用できるようになると、カスタム関数が登録されます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Json ファイルを作成する代わりに、可能な場合は JSON の自動生成を使用することをお勧めします。 自動生成は、ユーザーエラーが発生しにくく、 `yo office` スキャフォールディングファイルに既に含まれています。 JSDoc comment JSON ファイル生成のプロセスの詳細については、「[カスタム関数の json メタデータの生成](custom-functions-json-autogeneration.md)」を参照してください。

ただし、カスタム関数プロジェクトを最初から作成できます。そのためには、次のことを行う必要があります。

- JSON ファイルを作成します。
- マニフェストファイルが JSON ファイルに接続されていることを確認します。
- 関数を `id` `name` 登録するために、スクリプトファイルの関数とプロパティを関連付けます。

次の図は、スキャフォールディングファイルを使用すること `yo office` と、JSON を一から作成することの違いについて説明しています。
![Yo Office を使用して独自の JSON を作成することとの違いの画像](../images/custom-functions-json.png)

> [!NOTE]
> ジェネレーターを使用しない `<Resources>` 場合は、XML マニフェストファイルのセクションを使用して、作成した JSON ファイルにマニフェストを必ず接続してください `yo office` 。

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>メタデータの作成とマニフェストへの接続

プロジェクトで JSON ファイルを作成し、関数のパラメーターなど、関数に関するすべての詳細情報を提供します。 関数プロパティの完全なリストについては、[次のメタデータの例](#json-metadata-example)と[メタデータリファレンス](#metadata-reference)を参照してください。

次の例に示すように、XML マニフェストファイルで、セクション内の JSON ファイルを参照するようにし `<Resources>` ます。

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
> 完全なサンプル JSON ファイルは、 [Officedev/Excel-カスタム機能](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json)GitHub リポジトリのコミット履歴で入手できます。 JSON を自動的に生成するようにプロジェクトが調整されているため、手書きの JSON の完全なサンプルは、プロジェクトの以前のバージョンでのみ使用できます。

## <a name="metadata-reference"></a>メタデータリファレンス

### <a name="functions"></a>functions

`functions` プロパティは、カスタム関数オブジェクトの配列です。 次の表に、各オブジェクトのプロパティを示します。

| プロパティ      | データ型 | 必須 | 説明                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | いいえ       | Excel でエンド ユーザーに表示される関数の説明です。 たとえば、「**華氏の値を摂氏に変換する**」です。                                                            |
| `helpUrl`     | string    | いいえ       | 関数に関する情報を提供する URL です  (作業ウィンドウに表示されます)。たとえば、`http://contoso.com/help/convertcelsiustofahrenheit.html` です。                      |
| `id`          | 文字列    | はい      | 関数の一意の ID です。 この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。                                            |
| `name`        | 文字列    | はい      | Excel でエンド ユーザーに表示される関数の名前です。 Excel では、この関数名は、XML マニフェストファイルで指定されているカスタム関数の名前空間でプレフィックスされます。 |
| `options`     | オブジェクト    | いいえ       | Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 詳細については、[options](#options) に関する説明を参照してください。                                                          |
| `parameters`  | 配列     | はい      | 関数の入力パラメーターを定義する配列です。 詳細については、「 [parameters](#parameters) 」を参照してください。                                                                             |
| `result`      | object    | はい      | 関数が返す情報の種類を定義するオブジェクトです。 詳細については、[result](#result) に関する説明を参照してください。                                                                 |

### <a name="options"></a>options

`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 次の表に、`options` オブジェクトのプロパティを示します。

| プロパティ          | データ型 | 必須                               | 説明                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| :---------------- | :-------- | :------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cancelable`      | ブール   | いいえ<br/><br/>既定値は、`false` です。  | `true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `CancelableInvocation` ハンドラーが呼び出されます。 通常、取り消し可能な関数は、1つの結果を返す非同期関数で、データの要求のキャンセルを処理する必要がある場合にのみ使用されます。 関数は、ストリーミングと取り消しの両方にすることはできません。 詳細については、「[ストリーミング機能を作成する](custom-functions-web-reqs.md#make-a-streaming-function)」の最後の方にあるメモを参照してください。 |
| `requiresAddress` | ブール   | いいえ <br/><br/>既定値は、`false` です。 | の場合は、カスタム関数を `true` 呼び出したセルのアドレスにカスタム関数からアクセスできます。 カスタム関数を呼び出したセルのアドレスを取得するには、カスタム関数で context を使用します。 カスタム関数は、streaming と requiresAddress の両方として設定することはできません。 このオプションを使用する場合、' 呼び ' パラメーターは、オプションで渡された最後のパラメーターである必要があります。                                              |
| `stream`          | ブール   | いいえ<br/><br/>既定値は、`false` です。  | `true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。 このオプションは、株価などの急速に変化するデータ ソースに便利です。 この関数には、`return` ステートメントは含めないようにする必要があります。 代わりに、結果の値は `StreamingInvocation.setResult` コールバック メソッドの引数として渡されます。 詳細については、「[ストリーミング関数](custom-functions-web-reqs.md#make-a-streaming-function)」を参照してください。                                                                                                                                                                |
| `volatile`        | ブール   | いいえ <br/><br/>既定値は、`false` です。 | <br /><br /> 場合は、 `true` 数式の依存値が変更されたときだけでなく、Excel が再計算するたびに関数が再計算されます。 関数は、ストリーミングと揮発性の両方にすることはできません。 `stream` と `volatile` の両方のプロパティが `true` に設定されている場合は、揮発性のオプションが無視されます。                                                                                                                                                                                                                                                                                             |

### <a name="parameters"></a>parameters

`parameters` プロパティは、パラメーター オブジェクトの配列です。 次の表に、各オブジェクトのプロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  いいえ |  パラメーターの説明です。 これは、Excel の IntelliSense に表示されます。  |
|  `dimensionality`  |  string  |  いいえ  |  **スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。  |
|  `name`  |  文字列  |  はい  |  パラメーターの名前です。 この名前は、Excel の IntelliSense に表示されます。  |
|  `type`  |  文字列  |  いいえ  |  パラメーターのデータ型です。 **boolean**、**number**、**string**、または **any** が可能です。ここでは、前の 3 種類のいずれかを使用できます。 このプロパティが指定されていない場合、データ型の既定は **any** です。 |
|  `optional`  | ブール | いいえ | `true` の場合、パラメーターは省略可能です。 |
|`repeating`| ブール | いいえ | If は `true` 、指定された配列からパラメーターを設定します。 すべての繰り返しパラメーターは、定義によって省略可能なパラメーターとして扱われることに注意してください。  |

### <a name="result"></a>result

`result` オブジェクトは、この関数が返す情報の種類を定義します。 次の表に、`result` オブジェクトのプロパティを示します。

| プロパティ         | データ型 | 必須 | 説明                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | いいえ       | **スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。 |

## <a name="associating-function-names-with-json-metadata"></a>関数名を JSON メタデータに関連付ける

関数が正しく動作するには、関数のプロパティを JavaScript 実装に関連付ける必要があり `id` ます。 関連付けがあることを確認してください。そうしないと、関数は登録されず、Excel で使用できません。 次のコードサンプルは、メソッドを使用して関連付けを行う方法を示して `CustomFunctions.associate()` います。 このサンプルではカスタム関数 `add` を定義し、それを `id` プロパティ値が **ADD** の、JSON メタデータ ファイル内のオブジェクトに関連付けます。

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

次の JSON は、以前のカスタム関数 JavaScript コードに関連付けられている JSON メタデータを示しています。

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

- JavaScript ファイルで、各関数の後に、カスタム関数の関連付けを指定し `CustomFunctions.associate` ます。

次のサンプルは、JavaScript コード サンプルで定義された関数に対応する JSON メタデータを示します。 `id`プロパティと `name` プロパティの値は、大文字で記述します。これは、カスタム関数を記述するときのベストプラクティスです。 この JSON を追加する必要があるのは、自動生成を使用せずに、手動で独自の JSON ファイルを準備する場合だけです。 Autogeneration の詳細については、「[カスタム関数の JSON メタデータを作成](custom-functions-json-autogeneration.md)する」を参照してください。

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
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

[関数に名前を付けるためのベストプラクティス](custom-functions-naming.md)、または前述の手書き JSON メソッドを使用して[関数をローカライズ](custom-functions-localize.md)する方法について説明します。

## <a name="see-also"></a>関連項目

- [カスタム関数用の JSON メタデータの自動生成](custom-functions-json-autogeneration.md)
- [カスタム関数のパラメータオプション](custom-functions-parameter-options.md)
- [Excel でカスタム関数を作成する](custom-functions-overview.md)
