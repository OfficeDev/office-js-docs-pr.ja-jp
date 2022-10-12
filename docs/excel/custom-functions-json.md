---
title: Excel でカスタム関数の JSON メタデータを手動で作成する
description: Excel でカスタム関数の JSON メタデータを定義し、関数 ID と名前のプロパティを関連付けます。
ms.date: 10/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: b4bc9139b3e46bc64749a58537737db2f048ee82
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2022
ms.locfileid: "68540999"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a>カスタム関数の JSON メタデータを手動で作成する

[カスタム関数の概要](custom-functions-overview.md)に関する記事で説明されているように、カスタム関数プロジェクトには、JSON メタデータ ファイルとスクリプト (JavaScript または TypeScript) ファイルの両方を含めて関数を登録し、使用できるようにする必要があります。 カスタム関数は、ユーザーがアドインを初めて実行したときに登録され、その後、すべてのブックで同じユーザーが使用できるようになります。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

可能であれば、独自の JSON ファイルを作成する代わりに、JSON 自動生成を使用することをお勧めします。 自動生成はユーザー エラーの発生が少なく、 `yo office` スキャフォールディングされたファイルには既にこれが含まれています。 JSDoc タグと JSON 自動生成プロセスの詳細については、「 [カスタム関数の JSON メタデータを自動生成する」を](custom-functions-json-autogeneration.md)参照してください。

ただし、カスタム関数プロジェクトを最初から作成できます。 このプロセスでは、次のことを行う必要があります。

- JSON ファイルを記述します。
- マニフェスト ファイルが JSON ファイルに接続されていることを確認します。
- 関数を登録するために、スクリプト ファイル内の関数 `id` と `name` プロパティを関連付けます。

次の図では、スキャフォールディング ファイルの使用 `yo office` と JSON の最初からの書き込みの違いについて説明します。

![Office アドイン用の Yeoman ジェネレーターの使用と独自の JSON の作成の違いのイメージ。](../images/custom-functions-json.png)

> [!NOTE]
> [Office アドイン用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用しない場合は、XML マニフェスト ファイルのセクションを使用して **\<Resources\>**、作成した JSON ファイルにマニフェストを必ず接続してください。

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>メタデータの作成とマニフェストへの接続

プロジェクトに JSON ファイルを作成し、その中の関数に関するすべての詳細 (関数のパラメーターなど) を指定します。 関数プロパティの完全な一覧については、 [次のメタデータの例](#json-metadata-example) と [メタデータリファレンス](#metadata-reference) を参照してください。

次の例のように、XML マニフェスト ファイルがセクション内の **\<Resources\>** JSON ファイルを参照していることを確認します。

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
  "allowCustomDataForDataTypeAny": true,
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
> 完全なサンプル JSON ファイルは、 [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub リポジトリのコミット履歴で入手できます。 プロジェクトが JSON を自動的に生成するように調整されているため、手書きの JSON の完全なサンプルは、以前のバージョンのプロジェクトでのみ使用できます。

## <a name="metadata-reference"></a>メタデータリファレンス

### <a name="allowcustomdatafordatatypeany"></a>allowCustomDataForDataTypeAny

プロパティは `allowCustomDataForDataTypeAny` ブール型です。 この値を設定すると `true` 、カスタム関数がパラメーターとしてデータ型を受け入れ、値を返すことができます。 詳細については、「 [カスタム関数とデータ型](custom-functions-data-types-concepts.md)」を参照してください。

> [!NOTE]
> 他のほとんどの JSON メタデータ プロパティとは異なり、 `allowCustomDataForDataTypeAny` 最上位のプロパティであり、サブプロパティは含んでいません。 このプロパティを書式設定する方法の例については、上記の [JSON メタデータ コード サンプル](#json-metadata-example) を参照してください。

### <a name="allowerrorfordatatypeany"></a>allowErrorForDataTypeAny

プロパティは `allowErrorForDataTypeAny` ブール型です。 カスタム関数が入力値 `true` としてエラーを処理できるように値を設定します。 型`any`を持つすべてのパラメーター、または`any[][]`入力値としてエラーを受け入れる場合があります `true`。`allowErrorForDataTypeAny`. 既定値 `allowErrorForDataTypeAny` は `false`.

> [!NOTE]
> 他の JSON メタデータ プロパティとは異なり、 `allowErrorForDataTypeAny` 最上位のプロパティであり、サブプロパティは含んでいません。 このプロパティを書式設定する方法の例については、上記の [JSON メタデータ コード サンプル](#json-metadata-example) を参照してください。

### <a name="functions"></a>functions

`functions` プロパティは、カスタム関数オブジェクトの配列です。 次の表に、各オブジェクトのプロパティを示します。

| プロパティ      | データ型 | 必須 | 説明                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | 文字列    | いいえ       | Excel でエンド ユーザーに表示される関数の説明です。 たとえば、「**華氏の値を摂氏に変換する**」です。                                                            |
| `helpUrl`     | 文字列    | いいえ       | 関数に関する情報を提供する URL です  (作業ウィンドウに表示されます)。たとえば、`http://contoso.com/help/convertcelsiustofahrenheit.html` です。                      |
| `id`          | string    | はい      | 関数の一意の ID です。 この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。                                            |
| `name`        | 文字列    | はい      | Excel でエンド ユーザーに表示される関数の名前です。 Excel では、この関数名の前に、XML マニフェスト ファイルで指定されたカスタム関数名前空間が付けられます。 |
| `options`     | object    | いいえ       | Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 詳細については、[options](#options) に関する説明を参照してください。                                                          |
| `parameters`  | 配列     | はい      | 関数の入力パラメーターを定義する配列です。 詳細については [、パラメーター](#parameters) を参照してください。                                                                             |
| `result`      | object    | はい      | 関数が返す情報の種類を定義するオブジェクトです。 詳細については、[result](#result) に関する説明を参照してください。                                                                 |

### <a name="options"></a>options

`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 次の表に、`options` オブジェクトのプロパティを示します。

| プロパティ          | データ型 | 必須                               | 説明 |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | ブール   | いいえ<br/><br/>既定値は、`false` です。  | `true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `CancelableInvocation` ハンドラーが呼び出されます。 キャンセル可能な関数は通常、1 つの結果を返し、データの要求の取り消しを処理する必要がある非同期関数にのみ使用されます。 関数では、プロパティと`cancelable`プロパティの両方を`stream`使用できません。 |
| `requiresAddress` | ブール値   | いいえ <br/><br/>既定値は、`false` です。 | カスタム関数が呼び出したセルのアドレスにアクセスできる場合 `true`。 `address` [呼び出しパラメーター](custom-functions-parameter-options.md#invocation-parameter)のプロパティには、カスタム関数を呼び出したセルのアドレスが含まれています。 関数では、プロパティと`requiresAddress`プロパティの両方を`stream`使用できません。 |
| `requiresParameterAddresses` | ブール値   | いいえ <br/><br/>既定値は、`false` です。 | カスタム関数が関数の入力パラメーターのアドレスにアクセスできる場合 `true`。 このプロパティは[、結果](#result)オブジェクトのプロパティと`dimensionality`組み合わせて使用する必要`matrix`があり`dimensionality`、. 詳細については、「 [パラメーターのアドレスを検出](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) する」を参照してください。 |
| `stream`          | ブール値   | いいえ<br/><br/>既定値は、`false` です。  | `true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。 このオプションは、株価などの急速に変化するデータ ソースに便利です。 この関数には、`return` ステートメントは含めないようにする必要があります。 代わりに、結果の値がコールバック関数の `StreamingInvocation.setResult` 引数として渡されます。 詳細については、「 [ストリーミング関数を作成する](custom-functions-web-reqs.md#make-a-streaming-function)」を参照してください。 |
| `volatile`        | ブール値   | いいえ <br/><br/>既定値は、`false` です。 | 場合 `true`は、数式の依存値が変更されたときだけでなく、Excel が再計算するたびに関数が再計算されます。 関数では、プロパティと`volatile`プロパティの両方を`stream`使用できません。 プロパティと`volatile`プロパティの`stream`両方が設定`true`されている場合、volatile プロパティは無視されます。 |

### <a name="parameters"></a>parameters

`parameters` プロパティは、パラメーター オブジェクトの配列です。 次の表に、各オブジェクトのプロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  文字列  |  いいえ |  パラメーターの説明です。 これは Excel の IntelliSense に表示されます。  |
|  `dimensionality`  |  文字列  |  いいえ  |  `scalar` (配列以外の値) または `matrix` (2 次元配列) である必要があります。  |
|  `name`  |  string  |  はい  |  パラメーターの名前です。 この名前は、Excel の IntelliSense に表示されます。  |
|  `type`  |  文字列  |  いいえ  |  パラメーターのデータ型です。 `boolean`には、`number``string``any`前の 3 種類のいずれかを使用できます。 このプロパティが指定されていない場合、データ型の既定値 `any`は . |
|  `optional`  | ブール値 | いいえ | `true` の場合、パラメーターは省略可能です。 |
|`repeating`| ブール値 | いいえ | 場合 `true`は、指定した配列からパラメーターが設定されます。 関数のすべての繰り返しパラメーターは、定義によって省略可能なパラメーターと見なされることに注意してください。  |

### <a name="result"></a>result

`result` オブジェクトは、この関数が返す情報の種類を定義します。 次の表に、`result` オブジェクトのプロパティを示します。

| プロパティ         | データ型 | 必須 | 説明                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | 文字列    | いいえ       | `scalar` (配列以外の値) または `matrix` (2 次元配列) である必要があります。 |
| `type` | 文字列    | いいえ       | 結果のデータ型。 `boolean`、、`number``string`または `any` (前の 3 つの型のいずれかを使用できます) ことができます。 このプロパティが指定されていない場合、データ型の既定値 `any`は . |

## <a name="associating-function-names-with-json-metadata"></a>関数名を JSON メタデータに関連付ける

関数が正しく機能するには、関数の `id` プロパティを JavaScript 実装に関連付ける必要があります。 関連付けがあることを確認します。それ以外の場合、関数は登録されません。Excel では使用できません。 次のコード サンプルは、関数を使用して関連付けを行う方法を `CustomFunctions.associate()` 示しています。 このサンプルではカスタム関数 `add` を定義し、それを `id` プロパティ値が **ADD** の、JSON メタデータ ファイル内のオブジェクトに関連付けます。

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

- JavaScript ファイルで、各関数の後に使用する `CustomFunctions.associate` カスタム関数の関連付けを指定します。

次の例は、前の JavaScript コード サンプルで定義された関数に対応する JSON メタデータを示しています。 `id`プロパティ値と`name`プロパティ値は大文字です。これは、カスタム関数を記述する場合のベスト プラクティスです。 この JSON は、自動生成を使用せず、独自の JSON ファイルを手動で準備している場合にのみ追加する必要があります。 自動生成の詳細については、「 [カスタム関数の JSON メタデータを自動生成する」を](custom-functions-json-autogeneration.md)参照してください。

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

関数に [名前を付けるためのベスト プラクティス](custom-functions-naming.md) や、前述の手書き JSON メソッドを使用して [関数をローカライズ](custom-functions-localize.md) する方法について説明します。

## <a name="see-also"></a>関連項目

- [カスタム関数用の JSON メタデータの自動生成](custom-functions-json-autogeneration.md)
- [カスタム関数パラメーター オプション](custom-functions-parameter-options.md)
- [Excel でカスタム関数を作成する](custom-functions-overview.md)
