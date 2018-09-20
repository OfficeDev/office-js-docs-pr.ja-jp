# <a name="custom-function-metadata"></a>カスタム関数のメタデータ

Excel アドインに[カスタム関数](custom-functions-overview.md)を組み込む場合は、関数に関するメタデータを含む JSON ファイルをホストする必要があります (関数の JavaScript ファイルと、JavaScript ファイルの親として機能する UI を持たない HTML ファイルに加えて必要となります)。 この記事では、JSON ファイルの書式をサンプルを用いて説明します。

JSON ファイルの詳細なサンプルは[こちら](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json)でご覧いただけます。

## <a name="functions-array"></a>関数配列

メタデータは、オブジェクトの配列を値としてもつ単一の `functions` プロパティを含む JSON オブジェクトです。 各オブジェクトは、それぞれ 1 つのカスタム関数を表します。 次の表に、プロパティを示します。

|  プロパティ  |  データ型  |  必須かどうか  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  文字列  |  いいえ  |  Excel UI に表示される関数の説明。 例:「摂氏を華氏に変換する。」 |
|  `helpUrl`  |  文字列  |   いいえ  |  ユーザーが関数に関するヘルプを見ることができる URL。 (タスクペインに表示されます)。例: "http://contoso.com/help/convertcelsiustofahrenheit.html"  |
|  `name`  |  文字列  |  はい  |  ユーザーが関数を選択しているときに Excel の UI の (名前空間の先頭に) 表示される関数の名前。 関数の名前は JavaScript で定義されているものと同じでなければなりません。 |
|  `options`  |  オブジェクト  |  いいえ  |  Excel が関数を処理する方法を設定します。 詳細は、「[オプション オブジェクト](#options-object)」を参照してください。 |
|  `parameters`  |  配列  |  はい  |  関数に渡すパラメータに関するメタデータ。 詳細は、「[パラメータ配列](#parameters-array)」を参照してください。 |
|  `result`  |  オブジェクト  |  はい  |  関数が返す値に関するメタデータ。 詳細は、「[結果オブジェクト](#result-object)」を参照してください。 |

## <a name="options-object"></a>Options オブジェクト

オブジェクトは、Excel が関数を処理する方法を設定します。`options` 次の表に、プロパティを示します。

|  プロパティ  |  データ型  |  必須かどうか  |  説明  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  ブール値  |  いいえ。既定値は `false` です。  |  `true` を使用する場合、関数をキャンセルすることになる操作をユーザーが実行するたびに Excel は、 `onCanceled` ハンドラーを呼び出します。例えば、手動で再計算をトリガーしたり、関数が参照しているセルを編集したりなどの操作です。このオプションを使用する場合、Excel は、`caller` パラメータを追加して、JavaScript 関数を呼び出します 。(`parameters` プロパティにこのパラメータを登録***しない***でください )。関数の本文では、`caller.onCanceled` のメンバーにハンドラーを割り当てる必要があります。|
|  `stream`  |  ブール値  |  いいえ。既定値は `false` です。  |  の場合、関数は一度の呼び出しで繰り返しセルに出力できます。`true` このオプションは、株価など急激に変化するデータソースで役立ちます。 このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。 (このパラメーターを `parameters` プロパティに登録***しない***でください)。 関数では、`return` 文を使いません。 代わりに、戻り値を `caller.setResult` コールバック メソッドの引数として渡します。|

## <a name="parameters-array"></a>パラメータ配列

プロパティはオブジェクトの配列です。`parameters` 各オブジェクトはそれぞれ 1 つのパラメータを表します。 次の表に、プロパティを示します。

|  プロパティ  |  データ型  |  必須かどうか  |  説明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  文字列  |  いいえ |  パラメータの説明。  |
|  `dimensionality`  |  文字列  |  はい  |  非配列値を意味する "scalar" か、行配列の配列を意味する "matrix" のどちらかでなければなりません。  |
|  `name`  |  文字列  |  はい  |  パラメーターの名前です。 この名前は Excel の IntelliSense で表示されます。  |
|  `type`  |  文字列  |  はい  |  パラメータのデータ型。 "boolean"、"number"、または "string" のいずれかです。  |

## <a name="result-object"></a>結果オブジェクト

プロパティは、関数から返された値に関するメタデータです。`results` 次の表に、プロパティを示します。

|  プロパティ  |  データ型  |  必須かどうか  |  説明  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  文字列  |  いいえ  |  非配列値を意味する "scalar" か、行配列の配列を意味する "matrix" のどちらかでなければなりません。  |
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
            ]
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
            ]
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
            ]
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
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
            ]
        }
    ]
}

```

## <a name="see-also"></a>関連項目
[カスタム関数](custom-functions-overview.md)<br>
[配列数式のガイドラインと例](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
