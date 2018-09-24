---
ms.date: 09/20/2018
description: Excel でカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ
ms.openlocfilehash: 815b0c6e65966867d9e5d953a40ffc705a63ee63
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062145"
---
# <a name="custom-functions-metadata"></a>カスタム関数のメタデータ

Excel アドインで[カスタム関数](custom-functions-overview.md) を定義する場合には、Excel でカスタム関数を登録してエンドユーザーが使用できるようにするための情報を提供する JSON メタデータ ファイルを、アドイン プロジェクトに含める必要があります。 この記事では、JSON メタデータ ファイルの形式について説明します。

> [!NOTE]
> カスタム関数を有効にするためにアドイン プロジェクトに含める必要のある、その他のファイルに関する情報については、「[Excel でカスタム関数を作成する](custom-functions-overview.md#learn-the-basics)」を参照してください。

## <a name="example-metadata"></a>メタデータの例

次の例は、カスタム関数を定義するアドイン用の JSON メタデータ ファイルの内容を示しています。 この例に続くセクションでは、この JSON の例の中にある個々のプロパティについての詳細情報を提供します。

```json
{
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ADD42ASYNC",
            "name": "ADD42ASYNC",
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ISEVEN",
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "GETDAY",
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
            "id": "INCREMENTVALUE",
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
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
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
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
|  `description`  |  string  |  いいえ  |  Excel UI に表示される関数の説明。 たとえば、「**摂氏の値を華氏に変換します**」など。 |
|  `helpUrl`  |  string  |   いいえ  |  ユーザーが関数についての情報を見ることができる URL です。 (作業ウィンドウに表示されます。) たとえば、 **http://contoso.com/help/convertcelsiustofahrenheit.html**。 |
| `id`     | string | はい | 関数の一意の ID です。 設定後は、この ID は変更しないでください。 |
|  `name`  |  string  |  はい  |  ユーザーが関数を選択しているときに Excel の UI の (名前空間の先頭に) 表示される関数の名前。 JavaScript で定義されているものと同じ関数の名前である必要はありません。 |
|  `options`  |  object  |  いいえ  |  Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。 詳細については、「[オプション オブジェクト](#options-object)」を参照してください。 |
|  `parameters`  |  array  |  はい  |  関数の入力パラメーターを定義する配列です。 詳細については、「[パラメーター配列](#parameters-array)」を参照してください。 |
|  `result`  |  object  |  はい  |  関数によって返される情報の種類を定義するオブジェクトです。 詳細については、「[結果オブジェクト](#result-object)」を参照してください。 |

## <a name="options"></a>options

`options` オブジェクトを使用すると、Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。 次の表で、`options` オブジェクトのプロパティを一覧表示します。

|  プロパティ  |  データ型  |  必須かどうか  |  説明  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  いいえ。既定値は `false` です。  |  の場合、Excel はユーザーが関数をキャンセルする操作をするたびに `onCanceled` ハンドラを呼び出します。たとえば、手動で再計算をトリガするか、関数が参照するセルを編集する場合です。`true` このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。 (このパラメータを `parameters` プロパティに登録***しない***でください)。 関数の本体では、ハンドラは `caller.onCanceled` メンバーに割り当てる必要があります。 詳細については、 「[関数をキャンセルする](custom-functions-overview.md#canceling-a-function)」を参照してください。 |
|  `stream`  |  boolean  |  いいえ。既定値は `false` です。  |  の場合、関数は一度の呼び出しで繰り返しセルに出力できます。`true` このオプションは、株価など急激に変化するデータソースで役立ちます。 このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。 (このパラメーターを `parameters` プロパティに登録***しない***でください)。 関数では、`return` 文を使いません。 代わりに、結果値を `caller.setResult` コールバック メソッドの引数として渡します。 詳細については、「[ストリーム関数](custom-functions-overview.md#streamed-functions)」を参照してください。 |

## <a name="parameters"></a>parameters

`parameters` プロパティは、パラメーター オブジェクトの配列です。 次の表で、各オブジェクトのプロパティを一覧表示します。

|  プロパティ  |  データ型  |  必須かどうか  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  いいえ |  パラメータの説明。  |
|  `dimensionality`  |  string  |  いいえ  |   **scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。  |
|  `name`  |  string  |  はい  |  パラメーターの名前です。 この名前は Excel の IntelliSense で表示されます。  |
|  `type`  |  string  |  いいえ  |  パラメーターのデータ型です。  **ブール値**、 **数値**、または **文字列**である必要があります。  |

## <a name="result"></a>result

関数によって返される情報の種類を定義する`results` オブジェクトです。 次の表で、`result` オブジェクトのプロパティを一覧表示します。

|  プロパティ  |  データ型  |  必須かどうか  |  説明  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  いいえ  |   **scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。 |
|  `type`  |  string  |  はい  |  パラメーターのデータ型です。  **ブール値**、 **数値**、または **文字列**である必要があります。  |

## <a name="see-also"></a>関連項目

* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)