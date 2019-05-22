---
ms.date: 05/03/2019
description: Excel のカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ
localization_priority: Normal
ms.openlocfilehash: d6cfd61eabc5b27105414082675b35d3ff0ceb41
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337168"
---
# <a name="custom-functions-metadata"></a>カスタム関数のメタデータ

Excel アドイン内で[カスタム関数](custom-functions-overview.md)を定義する場合、アドインプロジェクトには、カスタム関数を登録してエンドユーザーが使用できるようにするために excel が必要とする情報を提供する JSON メタデータファイルが含まれています。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

このファイルは、次のいずれかの方法で生成されます。

- 手書きの JSON ファイル
- 関数の先頭に入力した JSDoc コメントから

ユーザーが初めてアドインを実行したときに、すべてのブックの同じユーザーがそのアドインを使用できるようになると、カスタム関数が登録されます。

この記事では、JSON メタデータファイルの形式について説明しています (手動で記述する場合を想定しています)。 JSDoc comment JSON ファイル生成の詳細については、「[カスタム関数の json メタデータの生成](custom-functions-json-autogeneration.md)」を参照してください。

カスタム関数を有効にするためにアドイン プロジェクトに含める必要のあるその他のファイルに関する情報については、「[Excel でのカスタム関数の作成](custom-functions-overview.md)」を参照してください。

JSON ファイルをホストするサーバーでは、カスタム関数を Excel Online で正しく作動させるために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効に設定する必要があります。

## <a name="example-metadata"></a>メタデータの例

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
      "description":  "Count up from zero",
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
      "description":  "Get the second highest number from a range",
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
> 完全な JSON ファイルのサンプルは、[OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub リポジトリにあります。

## <a name="functions"></a>functions 

`functions` プロパティは、カスタム関数オブジェクトの配列です。 次の表に、各オブジェクトのプロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  いいえ  |  Excel でエンド ユーザーに表示される関数の説明です。 たとえば、「**華氏の値を摂氏に変換する**」です。 |
|  `helpUrl`  |  string  |   いいえ  |  関数に関する情報を提供する URL です  (作業ウィンドウに表示されます)。たとえば、`http://contoso.com/help/convertcelsiustofahrenheit.html` です。 |
| `id`     | 文字列 | はい | 関数の一意の ID です。 この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。 |
|  `name`  |  文字列  |  はい  |  Excel でエンド ユーザーに表示される関数の名前です。 Excel では、この関数名は XML マニフェスト ファイルで指定されているカスタム関数の名前空間でプレフィックスされます。 |
|  `options`  |  オブジェクト  |  いいえ  |  Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 詳細については、[options](#options) に関する説明を参照してください。 |
|  `parameters`  |  配列  |  はい  |  関数の入力パラメーターを定義する配列です。 詳細については、[parameters](#parameters) に関する説明を参照してください。 |
|  `result`  |  object  |  はい  |  関数が返す情報の種類を定義するオブジェクトです。 詳細については、[result](#result) に関する説明を参照してください。 |

## <a name="options"></a>options

`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 次の表に、`options` オブジェクトのプロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  ブール  |  いいえ<br/><br/>既定値は、`false` です。  |  `true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `onCanceled` ハンドラーが呼び出されます。 このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します  (このパラメーターを `parameters` プロパティには登録し***ない***でください)。 この関数の本文では、ハンドラーを `caller.onCanceled` メンバーに割り当てる必要があります。 詳細については、「[関数をキャンセルする](custom-functions-web-reqs.md#stream-and-cancel-functions)」を参照してください。 |
|  `requiresAddress`  | ブール | いいえ <br/><br/>既定値は、`false` です。 | <br /><br /> True の場合、カスタム関数は、カスタム関数を呼び出したセルのアドレスにアクセスできます。 カスタム関数を呼び出したセルのアドレスを取得するには、カスタム関数で context を使用します。 詳しくは、「[カスタム関数が呼び出したセルを特定する](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function)」をご覧ください。 カスタム関数は、streaming と requiresAddress の両方として設定することはできません。 このオプションを使用する場合、' invocationContext ' パラメーターは、オプションで渡された最後のパラメーターである必要があります。 |
|  `stream`  |  ブール  |  いいえ<br/><br/>既定値は、`false` です。  |  `true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。 このオプションは、株価などの急速に変化するデータ ソースに便利です。 このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します  (このパラメーターを `parameters` プロパティには登録し***ない***でください)。 この関数には、`return` ステートメントは含めないようにする必要があります。 代わりに、結果の値は `caller.setResult` コールバック メソッドの引数として渡されます。 詳細については、「[ストリーミング関数](custom-functions-web-reqs.md#stream-and-cancel-functions)」を参照してください。 |
|  `volatile`  | ブール | いいえ <br/><br/>既定値は、`false` です。 | <br /><br /> `true` の場合は、数式の依存値が変更されたときのみではなく、Excel が再計算するたびに関数が再計算されます。 関数は、ストリーミングと揮発性の両方にすることはできません。 `stream` と `volatile` の両方のプロパティが `true` に設定されている場合は、揮発性のオプションが無視されます。 |

## <a name="parameters"></a>parameters

`parameters` プロパティは、パラメーター オブジェクトの配列です。 次の表に、各オブジェクトのプロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  いいえ |  パラメーターの説明です。 これは、Excel の intelliSense に表示されます。  |
|  `dimensionality`  |  string  |  いいえ  |  **スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。  |
|  `name`  |  文字列  |  はい  |  パラメーターの名前です。 この名前は、Excel の intelliSense に表示されます。  |
|  `type`  |  文字列  |  いいえ  |  パラメーターのデータ型です。 **boolean**、**number**、**string**、または **any** が可能です。ここでは、前の 3 種類のいずれかを使用できます。 このプロパティが指定されていない場合、データ型の既定は **any** です。 |
|  `optional`  | ブール | いいえ | `true` の場合、パラメーターは省略可能です。 |

## <a name="result"></a>result

`result` オブジェクトは、この関数が返す情報の種類を定義します。 次の表に、`result` オブジェクトのプロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  いいえ  |  **スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。 |

## <a name="next-steps"></a>次の手順
[関数に名前を付けるためのベストプラクティス](custom-functions-naming.md)、または前述の手書き JSON メソッドを使用して[関数をローカライズ](custom-functions-localize.md)する方法について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数用の JSON メタデータの自動生成](custom-functions-json-autogeneration.md)
* [カスタム関数のパラメータオプション](custom-functions-parameter-options.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)