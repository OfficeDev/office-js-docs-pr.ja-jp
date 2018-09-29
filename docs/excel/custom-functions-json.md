---
ms.date: 09/27/2018
description: Excel でカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ
ms.openlocfilehash: a179a9c4bc071200cab1377c5e48913bfc8358cf
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348795"
---
# <a name="custom-functions-metadata-preview"></a>カスタム関数のメタデータ (プレビュー)

Excel アドインで [カスタム関数](custom-functions-overview.md) を定義するときに、アドイン プロジェクトは、Excel がカスタム関数を登録し、エンド ユーザーが利用できるようにする必要がある情報を提供する JSON メタデータ ファイルを含める必要があります。この記事では、JSON メタデータ ファイルの形式について説明します。

カスタム関数を有効にするためにアドイン プロジェクトに含める必要のある、その他のファイルに関する情報については、「[Excel でカスタム関数を作成する](custom-functions-overview.md)」を参照してください。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a>メタデータの例

次の例は、カスタム関数を定義するアドイン用の JSON メタデータ ファイルの内容を示しています。 この例に続くセクションでは、この JSON の例の中にある個々のプロパティについての詳細情報を提供します。

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
> JSON ファイルの完全なサンプルは、「[OfficeDev/Excel-Custom-Functions GitHub リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json)」で利用可能です。

## <a name="functions"></a>functions 

`functions` プロパティは、カスタム関数オブジェクトの配列です。 次の表で、各オブジェクトのプロパティを一覧表示します。

|  プロパティ  |  データ型  |  必須かどうか  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  文字列  |  いいえ  |  エンド ユーザーに Excel で表示される関数の説明です。たとえば、 **華氏温度値を摂氏に変換**します。 |
|  `helpUrl`  |  文字列  |   いいえ  |  関数に関する情報を提供する URL です。(これは、作業ウィンドウに表示されます。) たとえば、**http://contoso.com/help/convertcelsiustofahrenheit.html**です。 |
| `id`     | 文字列 | はい | 関数の一意の ID です。 設定後は、この ID は変更しないでください。 |
|  `name`  |  文字列  |  はい  |  エンド ユーザーに Excel で表示される関数の名前です。 Excel では、この関数名は、XML マニフェスト ファイルで指定されているカスタム関数の名前空間が接頭辞となります。 |
|  `options`  |  object  |  いいえ  |  Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。 詳細については、「[オプション オブジェクト](#options-object)」を参照してください。 |
|  `parameters`  |  array  |  はい  |  関数の入力パラメーターを定義する配列です。 詳細については、「[パラメーター配列](#parameters-array)」を参照してください。 |
|  `result`  |  object  |  はい  |  関数によって返される情報の種類を定義するオブジェクトです。 詳細については、「[結果オブジェクト](#result-object)」を参照してください。 |

## <a name="options"></a>options

`options` オブジェクトを使用すると、Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。 次の表で、`options` オブジェクトのプロパティを一覧表示します。

|  プロパティ  |  データ型  |  必須かどうか  |  説明  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  ブール値  |  いいえ<br/><br/>既定値は`false` です。  |  `true` の場合、Excel はユーザーが関数をキャンセルする操作をするたびに `onCanceled` ハンドラを呼び出します。たとえば、手動で再計算をトリガーするか、関数が参照するセルを編集する場合です。 このオプションを使用すると、Excelは `caller` パラメーターを追加して JavaScript 関数を呼び出します。 (このパラメータを `parameters` プロパティに登録***しない***でください)。 関数の本体では、ハンドラは `caller.onCanceled` メンバーに割り当てる必要があります。 詳細については、 「[関数をキャンセルする](custom-functions-overview.md#canceling-a-function)」を参照してください。 |
|  `stream`  |  ブール値  |  いいえ<br/><br/>既定値は`false` です。  |  `true` の場合、関数は一度だけの呼び出しでも繰り返しセルに出力できます。 このオプションは、株価など急激に変化するデータソースで役立ちます。 このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。 (このパラメーターを `parameters` プロパティに登録***しない***でください)。 関数では、`return` 文を使いません。 代わりに、結果値を `caller.setResult` コールバック メソッドの引数として渡します。 詳細については、「[ストリーム関数](custom-functions-overview.md#streamed-functions)」を参照してください。 |

## <a name="parameters"></a>parameters

`parameters` プロパティは、パラメーター オブジェクトの配列です。 次の表で、各オブジェクトのプロパティを一覧表示します。

|  プロパティ  |  データ型  |  必須かどうか  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  文字列  |  いいえ |  パラメータの説明。  |
|  `dimensionality`  |  文字列  |  いいえ  |  **scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。  |
|  `name`  |  文字列  |  はい  |  パラメーターの名前です。 この名前は Excel の IntelliSense で表示されます。  |
|  `type`  |  文字列  |  いいえ  |  パラメーターのデータ型です。 **ブール値**、 **数値**、または **文字列**である必要があります。  |

## <a name="result"></a>result

関数によって返される情報の種類を定義する`results` オブジェクトです。 次の表で、`result` オブジェクトのプロパティを一覧表示します。

|  プロパティ  |  データ型  |  必須かどうか  |  説明  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  文字列  |  いいえ  |  **scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。 |
|  `type`  |  文字列  |  はい  |  パラメーターのデータ型。 **ブール値**、 **数値**、または **文字列**である必要があります。  |

## <a name="see-also"></a>関連項目

* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [Excel カスタム関数のチュートリアル](excel-tutorial-custom-functions.md)