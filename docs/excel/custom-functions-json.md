# <a name="custom-function-metadata"></a>カスタム関数のメタデータ

Excel アドインに[カスタム関数](custom-functions-overview.md)を組み込む場合は、関数に関するメタデータを含む JSON ファイルをホストする必要があります (関数の JavaScript ファイルと、JavaScript ファイルの親として機能する UI を持たない HTML ファイルに加えて必要となります)。 この記事では、JSON ファイルの書式をサンプルを用いて説明します。

JSON ファイルの詳細なサンプルは[こちら](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.json)でご覧いただけます。

## <a name="functions-array"></a>関数配列

メタデータは、オブジェクトの配列を値としてもつ単一の `functions` プロパティを含む JSON オブジェクトです。 各オブジェクトは、それぞれ 1 つのカスタム関数を表します。 次の表に、プロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  文字列  |  いいえ  |  Excel UI に表示される関数の説明。 例:「摂氏を華氏に変換する。」 |
|  `helpUrl`  |  文字列  |   いいえ  |  ユーザーが関数に関するヘルプを見ることができる URL。 (タスクペインに表示されます)。例: "http://contoso.com/help/convertcelsiustofahrenheit.html"  |
|  `name`  |  文字列  |  はい  |  ユーザーが関数を選択しているときに Excel の UI の (名前空間の先頭に) 表示される関数の名前。 関数の名前は JavaScript で定義されているものと同じでなければなりません。 |
|  `options`  |  オブジェクト  |  いいえ  |  Excel が関数を処理する方法を設定します。 詳細は、「[オプションオブジェクト](#options-object)」を参照してください。 |
|  `parameters`  |  配列  |  はい  |  関数に渡すパラメータに関するメタデータ。 詳細は、「[パラメータ配列](#parameters-array)」を参照してください。 |
|  `result`  |  オブジェクト  |  はい  |  関数が返す値に関するメタデータ。 詳細は、「[結果オブジェクト](#result-object)」をご覧ください。 |

## <a name="options-object"></a>オプション オブジェクト

`options` オブジェクトは、Excel が関数を処理する方法を設定します。 次の表に、プロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  ブール値  |  いいえ。既定値は `false` です。  |  `true` の場合、Excel はユーザーが関数をキャンセルする操作をするたびに `onCanceled` ハンドラを呼び出します。たとえば、手動で再計算をトリガするか、関数が参照するセルを編集する場合です。 このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。 (このパラメータを `parameters` プロパティに登録***しない***でください)。 関数本体では、`caller.onCanceled` メンバーにハンドラを割り当てる必要があります。 注意: `cancelable` と `sync` の両方を `true` にすることはできません。  |
|  `stream`  |  ブール値  |  いいえ。既定値は `false` です。  |  `true` の場合、関数は一度の呼び出しで繰り返しセルに出力できます。 このオプションは、株価など急激に変化するデータソースで役立ちます。 このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。 (このパラメータを `parameters` プロパティに登録***しない***でください)。 関数では、`return` 文を使いません。 代わりに、戻り値を `caller.setResult` コールバック メソッドの引数として渡します。 注意: `stream` と `sync` の両方が `true` ではない可能性があります。|
|  `sync`  |  ブール値  |  いいえ。既定値は `false`  |  `true` の場合、関数は同期して実行され、値を返す必要があります。 `false` の場合、関数は非同期に実行され、`OfficeExtension.Promise` オブジェクトを返す必要があります。 注意: `sync` は、`cancelable` か `stream` が `true` の場合、`true` ではない可能性があります。  |

## <a name="parameters-array"></a>パラメータ配列

`parameters` プロパティはオブジェクトの配列です。 各オブジェクトはそれぞれ 1 つのパラメータを表します。 次の表に、プロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  文字列  |  いいえ |  パラメータの説明。  |
|  `dimensionality`  |  文字列  |  はい  |  非配列値を意味する "scalar" か、行配列の配列を意味する "matrix" のどちらかです。  |
|  `name`  |  文字列  |  はい  |  パラメータの名前です。 この名前は Excel の IntelliSense で表示されます。  |
|  `type`  |  文字列  |  はい  |  パラメータのデータ型。 "boolean"、"number"、または "string" のいずれかです。  |

## <a name="result-object"></a>結果オブジェクト

`results` プロパティは、関数から返された値に関するメタデータです。 次の表に、プロパティを示します。

|  プロパティ  |  データ型  |  必須  |  説明  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  文字列  |  いいえ  |  非配列値を意味する "scalar" か、行配列の配列を意味する "matrix" のどちらかです。  |
|  `type`  |  文字列  |  はい  |  パラメータのデータ型。 "boolean"、"number"、または "string" のいずれかです。  |

## <a name="example"></a>例

次の JSON コードは、カスタム関数のメタデータ ファイルの例です。

```json
{
    "functions": [
        {
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
            ],
            "options": {
                "sync": true
            }
        },
        {
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
            ],
            "options": {
                "sync": false
            }
        },
        {
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
            ],
            "options": {
                "sync": true
            }
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": [],
            "options": {
                "sync": true
            }
        },
        {
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
                "sync": false,
                "stream": true,
                "cancelable": true
            }
        },
        {
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
            ],
            "options": {
                "sync": true
            }
        }
    ]
}

```

## <a name="see-also"></a>関連項目
[カスタム関数](custom-functions-overview.md)<br>
[配列数式のガイドラインと例](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
