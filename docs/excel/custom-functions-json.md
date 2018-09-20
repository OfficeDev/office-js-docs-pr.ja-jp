# <a name="custom-function-metadata"></a><span data-ttu-id="e136d-101">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="e136d-101">Custom function metadata</span></span>

<span data-ttu-id="e136d-102">Excel アドインに[カスタム関数](custom-functions-overview.md)を組み込む場合は、関数に関するメタデータを含む JSON ファイルをホストする必要があります (関数の JavaScript ファイルと、JavaScript ファイルの親として機能する UI を持たない HTML ファイルに加えて必要となります)。</span><span class="sxs-lookup"><span data-stu-id="e136d-102">When you include [custom functions](custom-functions-overview.md) in an Excel add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file).</span></span> <span data-ttu-id="e136d-103">この記事では、JSON ファイルの書式をサンプルを用いて説明します。</span><span class="sxs-lookup"><span data-stu-id="e136d-103">This article describes the format of the JSON file with examples.</span></span>

<span data-ttu-id="e136d-104">JSON ファイルの詳細なサンプルは[こちら](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json)でご覧いただけます。</span><span class="sxs-lookup"><span data-stu-id="e136d-104">A complete sample JSON file is available [here](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions-array"></a><span data-ttu-id="e136d-105">関数配列</span><span class="sxs-lookup"><span data-stu-id="e136d-105">Functions array</span></span>

<span data-ttu-id="e136d-106">メタデータは、オブジェクトの配列を値としてもつ単一の `functions` プロパティを含む JSON オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="e136d-106">The metadata is a JSON object that contains a single `functions` property whose value is an array of objects.</span></span> <span data-ttu-id="e136d-107">各オブジェクトは、それぞれ 1 つのカスタム関数を表します。</span><span class="sxs-lookup"><span data-stu-id="e136d-107">Each of these objects represents one custom function.</span></span> <span data-ttu-id="e136d-108">次の表に、プロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="e136d-108">The following table contains its properties:</span></span>

|  <span data-ttu-id="e136d-109">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e136d-109">Property</span></span>  |  <span data-ttu-id="e136d-110">データ型</span><span class="sxs-lookup"><span data-stu-id="e136d-110">Data Type</span></span>  |  <span data-ttu-id="e136d-111">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="e136d-111">Required?</span></span>  |  <span data-ttu-id="e136d-112">説明</span><span class="sxs-lookup"><span data-stu-id="e136d-112">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="e136d-113">文字列</span><span class="sxs-lookup"><span data-stu-id="e136d-113">string</span></span>  |  <span data-ttu-id="e136d-114">いいえ</span><span class="sxs-lookup"><span data-stu-id="e136d-114">No</span></span>  |  <span data-ttu-id="e136d-115">Excel UI に表示される関数の説明。</span><span class="sxs-lookup"><span data-stu-id="e136d-115">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="e136d-116">例:「摂氏を華氏に変換する。」</span><span class="sxs-lookup"><span data-stu-id="e136d-116">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="e136d-117">文字列</span><span class="sxs-lookup"><span data-stu-id="e136d-117">string</span></span>  |   <span data-ttu-id="e136d-118">いいえ</span><span class="sxs-lookup"><span data-stu-id="e136d-118">No</span></span>  |  <span data-ttu-id="e136d-119">ユーザーが関数に関するヘルプを見ることができる URL。</span><span class="sxs-lookup"><span data-stu-id="e136d-119">URL where your users can get help about the function.</span></span> <span data-ttu-id="e136d-120">(タスクペインに表示されます)。例: "http://contoso.com/help/convertcelsiustofahrenheit.html"</span><span class="sxs-lookup"><span data-stu-id="e136d-120">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span>  |
|  `name`  |  <span data-ttu-id="e136d-121">文字列</span><span class="sxs-lookup"><span data-stu-id="e136d-121">string</span></span>  |  <span data-ttu-id="e136d-122">はい</span><span class="sxs-lookup"><span data-stu-id="e136d-122">Yes</span></span>  |  <span data-ttu-id="e136d-123">ユーザーが関数を選択しているときに Excel の UI の (名前空間の先頭に) 表示される関数の名前。</span><span class="sxs-lookup"><span data-stu-id="e136d-123">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="e136d-124">関数の名前は JavaScript で定義されているものと同じでなければなりません。</span><span class="sxs-lookup"><span data-stu-id="e136d-124">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="e136d-125">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e136d-125">object</span></span>  |  <span data-ttu-id="e136d-126">いいえ</span><span class="sxs-lookup"><span data-stu-id="e136d-126">No</span></span>  |  <span data-ttu-id="e136d-127">Excel が関数を処理する方法を設定します。</span><span class="sxs-lookup"><span data-stu-id="e136d-127">Configure how Excel processes the function.</span></span> <span data-ttu-id="e136d-128">詳細は、「[オプション オブジェクト](#options-object)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e136d-128">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="e136d-129">配列</span><span class="sxs-lookup"><span data-stu-id="e136d-129">array</span></span>  |  <span data-ttu-id="e136d-130">はい</span><span class="sxs-lookup"><span data-stu-id="e136d-130">Yes</span></span>  |  <span data-ttu-id="e136d-131">関数に渡すパラメータに関するメタデータ。</span><span class="sxs-lookup"><span data-stu-id="e136d-131">Metadata about the parameters to the function.</span></span> <span data-ttu-id="e136d-132">詳細は、「[パラメータ配列](#parameters-array)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e136d-132">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="e136d-133">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e136d-133">object</span></span>  |  <span data-ttu-id="e136d-134">はい</span><span class="sxs-lookup"><span data-stu-id="e136d-134">Yes</span></span>  |  <span data-ttu-id="e136d-135">関数が返す値に関するメタデータ。</span><span class="sxs-lookup"><span data-stu-id="e136d-135">Metadata about the value returned by the function.</span></span> <span data-ttu-id="e136d-136">詳細は、「[結果オブジェクト](#result-object)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e136d-136">See [result object](#result-object) for details.</span></span> |

## <a name="options-object"></a><span data-ttu-id="e136d-137">Options オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e136d-137">Options object</span></span>

<span data-ttu-id="e136d-138">オブジェクトは、Excel が関数を処理する方法を設定します。`options`</span><span class="sxs-lookup"><span data-stu-id="e136d-138">The `options` object configures how Excel processes the function.</span></span> <span data-ttu-id="e136d-139">次の表に、プロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="e136d-139">The following table contains its properties:</span></span>

|  <span data-ttu-id="e136d-140">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e136d-140">Property</span></span>  |  <span data-ttu-id="e136d-141">データ型</span><span class="sxs-lookup"><span data-stu-id="e136d-141">Data Type</span></span>  |  <span data-ttu-id="e136d-142">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="e136d-142">Required?</span></span>  |  <span data-ttu-id="e136d-143">説明</span><span class="sxs-lookup"><span data-stu-id="e136d-143">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="e136d-144">ブール値</span><span class="sxs-lookup"><span data-stu-id="e136d-144">boolean</span></span>  |  <span data-ttu-id="e136d-145">いいえ。既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="e136d-145">No, default is `false`.</span></span>  |  <span data-ttu-id="e136d-p110">`true` を使用する場合、関数をキャンセルすることになる操作をユーザーが実行するたびに Excel は、 `onCanceled` ハンドラーを呼び出します。例えば、手動で再計算をトリガーしたり、関数が参照しているセルを編集したりなどの操作です。このオプションを使用する場合、Excel は、`caller` パラメータを追加して、JavaScript 関数を呼び出します 。(`parameters` プロパティにこのパラメータを登録***しない***でください )。関数の本文では、`caller.onCanceled` のメンバーにハンドラーを割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="e136d-p110">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. Note,  and  cannot both be .</span></span>|
|  `stream`  |  <span data-ttu-id="e136d-150">ブール値</span><span class="sxs-lookup"><span data-stu-id="e136d-150">boolean</span></span>  |  <span data-ttu-id="e136d-151">いいえ。既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="e136d-151">No, default is `false`.</span></span>  |  <span data-ttu-id="e136d-152">の場合、関数は一度の呼び出しで繰り返しセルに出力できます。`true`</span><span class="sxs-lookup"><span data-stu-id="e136d-152">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="e136d-153">このオプションは、株価など急激に変化するデータソースで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="e136d-153">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="e136d-154">このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="e136d-154">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="e136d-155">(このパラメーターを `parameters` プロパティに登録***しない***でください)。</span><span class="sxs-lookup"><span data-stu-id="e136d-155">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="e136d-156">関数では、`return` 文を使いません。</span><span class="sxs-lookup"><span data-stu-id="e136d-156">The function should have no `return` statement.</span></span> <span data-ttu-id="e136d-157">代わりに、戻り値を `caller.setResult` コールバック メソッドの引数として渡します。</span><span class="sxs-lookup"><span data-stu-id="e136d-157">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span>|

## <a name="parameters-array"></a><span data-ttu-id="e136d-158">パラメータ配列</span><span class="sxs-lookup"><span data-stu-id="e136d-158">Parameters array</span></span>

<span data-ttu-id="e136d-159">プロパティはオブジェクトの配列です。`parameters`</span><span class="sxs-lookup"><span data-stu-id="e136d-159">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="e136d-160">各オブジェクトはそれぞれ 1 つのパラメータを表します。</span><span class="sxs-lookup"><span data-stu-id="e136d-160">Each of these objects represents a parameter.</span></span> <span data-ttu-id="e136d-161">次の表に、プロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="e136d-161">The following table contains its properties:</span></span>

|  <span data-ttu-id="e136d-162">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e136d-162">Property</span></span>  |  <span data-ttu-id="e136d-163">データ型</span><span class="sxs-lookup"><span data-stu-id="e136d-163">Data Type</span></span>  |  <span data-ttu-id="e136d-164">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="e136d-164">Required?</span></span>  |  <span data-ttu-id="e136d-165">説明</span><span class="sxs-lookup"><span data-stu-id="e136d-165">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="e136d-166">文字列</span><span class="sxs-lookup"><span data-stu-id="e136d-166">string</span></span>  |  <span data-ttu-id="e136d-167">いいえ</span><span class="sxs-lookup"><span data-stu-id="e136d-167">No</span></span> |  <span data-ttu-id="e136d-168">パラメータの説明。</span><span class="sxs-lookup"><span data-stu-id="e136d-168">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="e136d-169">文字列</span><span class="sxs-lookup"><span data-stu-id="e136d-169">string</span></span>  |  <span data-ttu-id="e136d-170">はい</span><span class="sxs-lookup"><span data-stu-id="e136d-170">Yes</span></span>  |  <span data-ttu-id="e136d-171">非配列値を意味する "scalar" か、行配列の配列を意味する "matrix" のどちらかでなければなりません。</span><span class="sxs-lookup"><span data-stu-id="e136d-171">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `name`  |  <span data-ttu-id="e136d-172">文字列</span><span class="sxs-lookup"><span data-stu-id="e136d-172">string</span></span>  |  <span data-ttu-id="e136d-173">はい</span><span class="sxs-lookup"><span data-stu-id="e136d-173">Yes</span></span>  |  <span data-ttu-id="e136d-174">パラメーターの名前です。</span><span class="sxs-lookup"><span data-stu-id="e136d-174">The name of the parameter.</span></span> <span data-ttu-id="e136d-175">この名前は Excel の IntelliSense で表示されます。</span><span class="sxs-lookup"><span data-stu-id="e136d-175">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="e136d-176">文字列</span><span class="sxs-lookup"><span data-stu-id="e136d-176">string</span></span>  |  <span data-ttu-id="e136d-177">はい</span><span class="sxs-lookup"><span data-stu-id="e136d-177">Yes</span></span>  |  <span data-ttu-id="e136d-178">パラメータのデータ型。</span><span class="sxs-lookup"><span data-stu-id="e136d-178">The data type of the parameter.</span></span> <span data-ttu-id="e136d-179">"boolean"、"number"、または "string" のいずれかです。</span><span class="sxs-lookup"><span data-stu-id="e136d-179">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result-object"></a><span data-ttu-id="e136d-180">結果オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e136d-180">Result object</span></span>

<span data-ttu-id="e136d-181">プロパティは、関数から返された値に関するメタデータです。`results`</span><span class="sxs-lookup"><span data-stu-id="e136d-181">The `results` property provides metadata about the value returned from the function.</span></span> <span data-ttu-id="e136d-182">次の表に、プロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="e136d-182">The following table contains its properties:</span></span>

|  <span data-ttu-id="e136d-183">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e136d-183">Property</span></span>  |  <span data-ttu-id="e136d-184">データ型</span><span class="sxs-lookup"><span data-stu-id="e136d-184">Data Type</span></span>  |  <span data-ttu-id="e136d-185">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="e136d-185">Required?</span></span>  |  <span data-ttu-id="e136d-186">説明</span><span class="sxs-lookup"><span data-stu-id="e136d-186">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="e136d-187">文字列</span><span class="sxs-lookup"><span data-stu-id="e136d-187">string</span></span>  |  <span data-ttu-id="e136d-188">いいえ</span><span class="sxs-lookup"><span data-stu-id="e136d-188">No</span></span>  |  <span data-ttu-id="e136d-189">非配列値を意味する "scalar" か、行配列の配列を意味する "matrix" のどちらかでなければなりません。</span><span class="sxs-lookup"><span data-stu-id="e136d-189">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `type`  |  <span data-ttu-id="e136d-190">文字列</span><span class="sxs-lookup"><span data-stu-id="e136d-190">string</span></span>  |  <span data-ttu-id="e136d-191">はい</span><span class="sxs-lookup"><span data-stu-id="e136d-191">Yes</span></span>  |  <span data-ttu-id="e136d-192">パラメータのデータ型。</span><span class="sxs-lookup"><span data-stu-id="e136d-192">The data type of the parameter.</span></span> <span data-ttu-id="e136d-193">"boolean"、"number"、または "string" のいずれかです。</span><span class="sxs-lookup"><span data-stu-id="e136d-193">Must be "boolean", "number", or "string".</span></span>  |

## <a name="example"></a><span data-ttu-id="e136d-194">例</span><span class="sxs-lookup"><span data-stu-id="e136d-194">Example</span></span>

<span data-ttu-id="e136d-195">次の JSON コードは、カスタム関数のメタデータ ファイルの例です。</span><span class="sxs-lookup"><span data-stu-id="e136d-195">The following JSON code is an example of a metadata file for custom functions.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="e136d-196">関連項目</span><span class="sxs-lookup"><span data-stu-id="e136d-196">See also</span></span>
[<span data-ttu-id="e136d-197">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="e136d-197">Custom functions</span></span>](custom-functions-overview.md)<br>
[<span data-ttu-id="e136d-198">配列数式のガイドラインと例</span><span class="sxs-lookup"><span data-stu-id="e136d-198">Guidelines and examples of array formulas</span></span>](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
