---
ms.date: 10/17/2018
description: Excel でカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ
ms.openlocfilehash: cff1cbc22f39c99597d4abe7005d7b8bbce6e185
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640009"
---
# <a name="custom-functions-metadata-preview"></a>カスタム関数のメタデータ (プレビュー)

Excel アドインで [カスタム関数](custom-functions-overview.md) を定義する場合、アドイン プロジェクトには、Excel がカスタム関数を登録してエンド ユーザーが利用できるようにするために必要な情報を提供する JSON メタデータ ファイルを含める必要があります。この記事では、JSON メタデータ ファイルの形式について説明します。

カスタム関数を有効にするためにアドイン プロジェクトに含める必要のあるその他のファイルに関する情報については、「[Excel でカスタム関数を作成する](custom-functions-overview.md)」を参照してください。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a>メタデータの例

次の使用例は、JSON のアドインでカスタム関数を定義するメタデータ ファイルの内容を示しています。次の使用例を次のセクションでは、この例を JSON 内の個別のプロパティに関する詳細情報を提供します。

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
> JSON ファイルの完全なサンプルは、「[OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) 」GitHub リポジトリで利用可能です。

## <a name="functions"></a>functions 

`functions` プロパティは、カスタム関数オブジェクトの配列です。次の表は、各オブジェクトのプロパティを一覧表示しています。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  文字列  |  いいえ  |  Excel でエンド ユーザーに表示される関数の説明です。例: "**華氏温度値を摂氏に変換する**" 。 |
|  `helpUrl`  |  文字列  |   いいえ  |  関数に関する情報を提供する URL です。(作業ウィンドウに表示されます。) 例: **http://contoso.com/help/convertcelsiustofahrenheit.html**。 |
| `id`     | 文字列 | はい | 関数の一意の ID です。 この ID は、英数字とピリオドのみを含めることができ、設定された後、変更してはいけません。 |
|  `name`  |  文字列  |  はい  |  Excel でエンド ユーザーに表示される関数の名前です。Excel では、この関数名が XML マニフェスト ファイルで指定されているカスタム関数の名前空間で接頭辞となります。 |
|  `options`  |  object  |  いいえ  |  Excel で関数を実行する方法とタイミングのいくつかの側面をカスタマイズできます。詳細については、 [オプションのオブジェクト](#options-object) を参照してください。 |
|  `parameters`  |  配列  |  はい  |  関数の入力パラメーターを定義する配列。詳細については、 [パラメーター配列](#parameters-array) を参照してください。 |
|  `result`  |  object  |  はい  |  関数によって返される情報の種類を定義するオブジェクト。詳細については、 [結果のオブジェクト](#result-object) を参照してください。 |

## <a name="options"></a>options

`options` オブジェクトは、Excel で関数を実行する方法とタイミングのいくつかの側面をカスタマイズすることができます。次の表に`options` オブジェクトのプロパティを一覧表示しています。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  いいえ<br/><br/>既定値は `false` です。  |  `true` を使用する場合、関数をキャンセルすることになる操作をユーザーが実行するたびに Excel は、 `onCanceled` ハンドラーを呼び出します。例えば、手動で再計算をトリガーしたり、関数が参照しているセルを編集したりなどの操作です。このオプションを使用する場合、Excel は、`caller` パラメータを追加して、JavaScript 関数を呼び出します 。(`parameters` プロパティにこのパラメータを登録***しない***でください )。関数の本文では、`caller.onCanceled` のメンバーにハンドラーを割り当てる必要があります。詳細については、 [関数をキャンセルする](custom-functions-overview.md#canceling-a-function)を参照してください。 |
|  `stream`  |  ブール値  |  いいえ<br/><br/>既定値は `false` です。  |  `true` の場合、関数を一度呼び出すだけでセルに繰り返し出力できます。 このオプションは、株価など急激に変化するデータソースで役立ちます。 このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。 (このパラメーターを `parameters` プロパティに登録***しない***でください)。 関数では、`return` 文を使わないでください。 代わりに、戻り値を `caller.setResult` コールバック メソッドの引数として渡します。 詳細については、「 [ストリーミング関数](custom-functions-overview.md#streaming-functions)」を参照してください。 |

## <a name="parameters"></a>パラメーター

`parameters` プロパティは、カスタム関数オブジェクトの配列です。次の表は、各オブジェクトのプロパティを一覧表示しています。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  文字列  |  いいえ |  パラメータの説明。  |
|  `dimensionality`  |  文字列  |  いいえ  |  **scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。  |
|  `name`  |  文字列  |  はい  |  パラメーターの名前です。この名前は、Excel の intelliSense に表示されます。  |
|  `type`  |  文字列  |  いいえ  |  パラメーターのデータ型です。 **ブール値**、 **数値**、または **文字列**である必要があります。  |

## <a name="result"></a>result

`results` オブジェクトは、関数によって返される情報の種類を定義します。次の表のプロパティの `result` オブジェクトです。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  文字列  |  いいえ  |  **scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。 |
|  `type`  |  文字列  |  はい  |  パラメーターのデータ型です。 **ブール値**、 **数値**、または **文字列**である必要があります。  |

## <a name="see-also"></a>関連項目

* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [Excel カスタム関数のチュートリアル](excel-tutorial-custom-functions.md)
