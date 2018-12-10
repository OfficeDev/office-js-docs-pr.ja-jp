---
ms.date: 11/26/2018
description: Excel のカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ
ms.openlocfilehash: a3d4427af2c6ab46133cb4e2fd9ce384a6a8336c
ms.sourcegitcommit: e2ba9d7210c921d068f40d9f689314c73ad5ab4a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/05/2018
ms.locfileid: "27156594"
---
# <a name="custom-functions-metadata-preview"></a>カスタム関数のメタデータ (プレビュー)

Excel アドイン内に[カスタム関数](custom-functions-overview.md)を定義する場合、カスタム関数を登録し、エンド ユーザーが利用できるようにするために Excel が必要とする情報を提供する JSON メタデータ ファイルをアドイン プロジェクトに含める必要があります。 この記事では、その JSON メタデータ ファイルの形式について説明します。

カスタム関数を有効にするためにアドイン プロジェクトに含める必要のあるその他のファイルに関する情報については、「[Excel でのカスタム関数の作成](custom-functions-overview.md)」を参照してください。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

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
        "type": "string"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
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
        "type": "number",
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
> 完全な JSON ファイルのサンプルは、[OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub リポジトリにあります。

## <a name="functions"></a>functions 

`functions` プロパティは、カスタム関数オブジェクトの配列です。 次の表に、各オブジェクトのプロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  文字列  |  いいえ  |  Excel でエンド ユーザーに表示される関数の説明です。 たとえば、「**華氏の値を摂氏に変換する**」です。 |
|  `helpUrl`  |  文字列  |   いいえ  |  関数に関する情報を提供する URL です  (作業ウィンドウに表示されます)。たとえば、**http://contoso.com/help/convertcelsiustofahrenheit.html** です。 |
| `id`     | 文字列 | はい | 関数の一意の ID です。 この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。 |
|  `name`  |  文字列  |  はい  |  Excel でエンド ユーザーに表示される関数の名前です。 Excel では、この関数名は XML マニフェスト ファイルで指定されているカスタム関数の名前空間でプレフィックスされます。 |
|  `options`  |  オブジェクト  |  いいえ  |  Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 詳細については、[options オブジェクト](#options-object)に関する説明を参照してください。 |
|  `parameters`  |  配列  |  はい  |  関数の入力パラメーターを定義する配列です。 詳細については、[parameters 配列](#parameters-array)に関する説明を参照してください。 |
|  `result`  |  オブジェクト  |  はい  |  関数が返す情報の種類を定義するオブジェクトです。 詳細については、[result オブジェクト](#result-object)に関する説明を参照してください。 |

## <a name="options"></a>options

`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。 次の表に、`options` オブジェクトのプロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  ブール  |  いいえ<br/><br/>既定値は、`false` です。  |  `true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `onCanceled` ハンドラーが呼び出されます。 このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します  (このパラメーターを `parameters` プロパティには登録し***ない***でください)。 この関数の本文では、ハンドラーを `caller.onCanceled` メンバーに割り当てる必要があります。 詳細については、「[関数をキャンセルする](custom-functions-overview.md#canceling-a-function)」を参照してください。 |
|  `stream`  |  ブール  |  いいえ<br/><br/>既定値は、`false` です。  |  `true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。 このオプションは、株価などの急速に変化するデータ ソースに便利です。 このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します  (このパラメーターを `parameters` プロパティには登録し***ない***でください)。 この関数には、`return` ステートメントは含めないようにする必要があります。 代わりに、結果の値は `caller.setResult` コールバック メソッドの引数として渡されます。 詳細については、「[ストリーミング関数](custom-functions-overview.md#streaming-functions)」を参照してください。 |
|  `volatile`  | ブール | いいえ <br/><br/>既定値は、`false` です。 | <br /><br /> `true` の場合は、数式の依存値が変更されたときのみではなく、Excel が再計算するたびに関数が再計算されます。 関数は、ストリーミングと揮発性の両方にすることはできません。 `stream` と `volatile` の両方のプロパティが `true` に設定されている場合は、揮発性のオプションが無視されます。 |

## <a name="parameters"></a>parameters

`parameters` プロパティは、パラメーター オブジェクトの配列です。 次の表に、各オブジェクトのプロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  文字列  |  いいえ |  パラメーターの説明です。 これは、Excel の intelliSense に表示されます。  |
|  `dimensionality`  |  文字列  |  いいえ  |  **スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。  |
|  `name`  |  文字列  |  はい  |  パラメーターの名前です。 この名前は、Excel の intelliSense に表示されます。  |
|  `type`  |  文字列  |  いいえ  |  パラメーターのデータ型です。 **boolean**、**number**、**string**、または **any** が可能です。ここでは、前の 3 種類のいずれかを使用できます。 このプロパティが指定されていない場合、データ型の既定は **any** です。 |
|  `optional`  | ブール | いいえ | `true` の場合、パラメーターは省略可能です。 |

>[!NOTE]
> 省略可能なパラメーターの `type` プロパティが指定されていない場合や `any` に設定している場合は、Excel のセルに関数が入力されているときに、IDE の linting エラーや省略可能なパラメーターが表示されないなどの問題が発生することがあります。 これについては、2018 年 12 月に変更される予定です。

## <a name="result"></a>result

`result` オブジェクトは、この関数が返す情報の種類を定義します。 次の表に、`result` オブジェクトのプロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  文字列  |  いいえ  |  **スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。 |
|  `type`  |  文字列  |  はい  |  パラメーターのデータ型です。 **boolean**、**number**、**string**、または **any** である必要があります。ここでは、前の 3 種類のいずれかを使用できます。 |

## <a name="see-also"></a>関連項目

* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [チュートリアル: Excel でカスタム関数を作成します。](excel-tutorial-custom-functions.md)
