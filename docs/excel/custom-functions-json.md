# <a name="custom-function-metadata"></a><span data-ttu-id="e9552-101">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="e9552-101">Custom function metadata</span></span>

<span data-ttu-id="e9552-102">Excel アドインに[カスタム関数](custom-functions-overview.md)を組み込む場合は、関数に関するメタデータを含む JSON ファイルをホストする必要があります (関数の JavaScript ファイルと、JavaScript ファイルの親として機能する UI を持たない HTML ファイルに加えて必要となります)。</span><span class="sxs-lookup"><span data-stu-id="e9552-102">When you include [custom functions](custom-functions-overview.md) in an Excel add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file).</span></span> <span data-ttu-id="e9552-103">この記事では、JSON ファイルの書式をサンプルを用いて説明します。</span><span class="sxs-lookup"><span data-stu-id="e9552-103">This article describes the format of the JSON file with examples.</span></span>

<span data-ttu-id="e9552-104">JSON ファイルの詳細なサンプルは[こちら](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.json)でご覧いただけます。</span><span class="sxs-lookup"><span data-stu-id="e9552-104">A complete sample JSON file is available [here](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.json).</span></span>

## <a name="functions-array"></a><span data-ttu-id="e9552-105">関数配列</span><span class="sxs-lookup"><span data-stu-id="e9552-105">Functions array</span></span>

<span data-ttu-id="e9552-106">メタデータは、オブジェクトの配列を値としてもつ単一の `functions` プロパティを含む JSON オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="e9552-106">The metadata is a JSON object that contains a single `functions` property whose value is an array of objects.</span></span> <span data-ttu-id="e9552-107">各オブジェクトは、それぞれ 1 つのカスタム関数を表します。</span><span class="sxs-lookup"><span data-stu-id="e9552-107">Each of these objects represents one custom function.</span></span> <span data-ttu-id="e9552-108">次の表に、プロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="e9552-108">The following table contains its properties:</span></span>

|  <span data-ttu-id="e9552-109">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e9552-109">Property</span></span>  |  <span data-ttu-id="e9552-110">データ型</span><span class="sxs-lookup"><span data-stu-id="e9552-110">Data Type</span></span>  |  <span data-ttu-id="e9552-111">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="e9552-111">Required?</span></span>  |  <span data-ttu-id="e9552-112">説明</span><span class="sxs-lookup"><span data-stu-id="e9552-112">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="e9552-113">文字列</span><span class="sxs-lookup"><span data-stu-id="e9552-113">string</span></span>  |  <span data-ttu-id="e9552-114">いいえ</span><span class="sxs-lookup"><span data-stu-id="e9552-114">No</span></span>  |  <span data-ttu-id="e9552-115">Excel UI に表示される関数の説明。</span><span class="sxs-lookup"><span data-stu-id="e9552-115">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="e9552-116">例:「摂氏を華氏に変換する。」</span><span class="sxs-lookup"><span data-stu-id="e9552-116">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="e9552-117">文字列</span><span class="sxs-lookup"><span data-stu-id="e9552-117">string</span></span>  |   <span data-ttu-id="e9552-118">いいえ</span><span class="sxs-lookup"><span data-stu-id="e9552-118">No</span></span>  |  <span data-ttu-id="e9552-119">ユーザーが関数に関するヘルプを見ることができる URL。</span><span class="sxs-lookup"><span data-stu-id="e9552-119">URL where your users can get help about the function.</span></span> <span data-ttu-id="e9552-120">(タスクペインに表示されます)。例: "http://contoso.com/help/convertcelsiustofahrenheit.html"</span><span class="sxs-lookup"><span data-stu-id="e9552-120">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span>  |
|  `name`  |  <span data-ttu-id="e9552-121">文字列</span><span class="sxs-lookup"><span data-stu-id="e9552-121">string</span></span>  |  <span data-ttu-id="e9552-122">はい</span><span class="sxs-lookup"><span data-stu-id="e9552-122">Yes</span></span>  |  <span data-ttu-id="e9552-123">ユーザーが関数を選択しているときに Excel の UI の (名前空間の先頭に) 表示される関数の名前。</span><span class="sxs-lookup"><span data-stu-id="e9552-123">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="e9552-124">関数の名前は JavaScript で定義されているものと同じでなければなりません。</span><span class="sxs-lookup"><span data-stu-id="e9552-124">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="e9552-125">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e9552-125">object</span></span>  |  <span data-ttu-id="e9552-126">いいえ</span><span class="sxs-lookup"><span data-stu-id="e9552-126">No</span></span>  |  <span data-ttu-id="e9552-127">Excel が関数を処理する方法を設定します。</span><span class="sxs-lookup"><span data-stu-id="e9552-127">Configure how Excel processes the function.</span></span> <span data-ttu-id="e9552-128">詳細は、「[オプション オブジェクト](#options-object)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e9552-128">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="e9552-129">配列</span><span class="sxs-lookup"><span data-stu-id="e9552-129">array</span></span>  |  <span data-ttu-id="e9552-130">はい</span><span class="sxs-lookup"><span data-stu-id="e9552-130">Yes</span></span>  |  <span data-ttu-id="e9552-131">関数に渡すパラメータに関するメタデータ。</span><span class="sxs-lookup"><span data-stu-id="e9552-131">Metadata about the parameters to the function.</span></span> <span data-ttu-id="e9552-132">詳細は、「[パラメータ配列](#parameters-array)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e9552-132">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="e9552-133">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e9552-133">object</span></span>  |  <span data-ttu-id="e9552-134">はい</span><span class="sxs-lookup"><span data-stu-id="e9552-134">Yes</span></span>  |  <span data-ttu-id="e9552-135">関数が返す値に関するメタデータ。</span><span class="sxs-lookup"><span data-stu-id="e9552-135">Metadata about the value returned by the function.</span></span> <span data-ttu-id="e9552-136">詳細は、「[結果オブジェクト](#result-object)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e9552-136">See [result object](#result-object) for details.</span></span> |

## <a name="options-object"></a><span data-ttu-id="e9552-137">オプション オブジェクト
</span><span class="sxs-lookup"><span data-stu-id="e9552-137">Options object</span></span>

<span data-ttu-id="e9552-138">`options` オブジェクトは、Excel が関数を処理する方法を設定します。</span><span class="sxs-lookup"><span data-stu-id="e9552-138">The `options` object configures how Excel processes the function.</span></span> <span data-ttu-id="e9552-139">次の表に、プロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="e9552-139">The following table contains its properties:</span></span>

|  <span data-ttu-id="e9552-140">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e9552-140">Property</span></span>  |  <span data-ttu-id="e9552-141">データ型</span><span class="sxs-lookup"><span data-stu-id="e9552-141">Data Type</span></span>  |  <span data-ttu-id="e9552-142">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="e9552-142">Required?</span></span>  |  <span data-ttu-id="e9552-143">説明</span><span class="sxs-lookup"><span data-stu-id="e9552-143">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="e9552-144">ブール型</span><span class="sxs-lookup"><span data-stu-id="e9552-144">boolean</span></span>  |  <span data-ttu-id="e9552-145">いいえ。既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="e9552-145">No, default is `false`.</span></span>  |  <span data-ttu-id="e9552-146">`true` の場合、Excel はユーザーが関数をキャンセルする操作をするたびに `onCanceled` ハンドラを呼び出します。たとえば、手動で再計算をトリガするか、関数が参照するセルを編集する場合です。</span><span class="sxs-lookup"><span data-stu-id="e9552-146">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="e9552-147">このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="e9552-147">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="e9552-148">(このパラメータを `parameters` プロパティに登録***しない***でください)。</span><span class="sxs-lookup"><span data-stu-id="e9552-148">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="e9552-149">関数本体では、`caller.onCanceled` メンバーにハンドラを割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="e9552-149">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="e9552-150">注意: `cancelable` と `sync` の両方を `true` にすることはできません。</span><span class="sxs-lookup"><span data-stu-id="e9552-150">Note, `cancelable` and `sync` cannot both be `true`.</span></span>  |
|  `stream`  |  <span data-ttu-id="e9552-151">ブール型</span><span class="sxs-lookup"><span data-stu-id="e9552-151">boolean</span></span>  |  <span data-ttu-id="e9552-152">いいえ。既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="e9552-152">No, default is `false`.</span></span>  |  <span data-ttu-id="e9552-153">`true` の場合、関数は一度の呼び出しで繰り返しセルに出力できます。</span><span class="sxs-lookup"><span data-stu-id="e9552-153">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="e9552-154">このオプションは、株価など急激に変化するデータソースで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="e9552-154">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="e9552-155">このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="e9552-155">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="e9552-156">(このパラメーターを `parameters` プロパティに登録***しない***でください)。</span><span class="sxs-lookup"><span data-stu-id="e9552-156">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="e9552-157">関数では、`return` 文を使いません。</span><span class="sxs-lookup"><span data-stu-id="e9552-157">The function should have no `return` statement.</span></span> <span data-ttu-id="e9552-158">代わりに、戻り値を `caller.setResult` コールバック メソッドの引数として渡します。</span><span class="sxs-lookup"><span data-stu-id="e9552-158">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="e9552-159">注意: `stream` と `sync` の両方が `true` ではない可能性があります。</span><span class="sxs-lookup"><span data-stu-id="e9552-159">Note, `stream` and `sync` may not both be `true`.</span></span>|
|  `sync`  |  <span data-ttu-id="e9552-160">ブール型</span><span class="sxs-lookup"><span data-stu-id="e9552-160">boolean</span></span>  |  <span data-ttu-id="e9552-161">いいえ。既定値は `false`</span><span class="sxs-lookup"><span data-stu-id="e9552-161">No, default is `false`</span></span>  |  <span data-ttu-id="e9552-162">`true` の場合、関数は同期して実行され、値を返す必要があります。</span><span class="sxs-lookup"><span data-stu-id="e9552-162">If `true`, the function runs synchronously and it must return a value.</span></span> <span data-ttu-id="e9552-163">`false` の場合、関数は非同期に実行され、`OfficeExtension.Promise` オブジェクトを返す必要があります。</span><span class="sxs-lookup"><span data-stu-id="e9552-163">If `false`, the function runs asynchronously and it must return a `OfficeExtension.Promise` object.</span></span> <span data-ttu-id="e9552-164">注意: `sync` は、`cancelable` か `stream` が `true` の場合、`true` ではない可能性があります。</span><span class="sxs-lookup"><span data-stu-id="e9552-164">Note, `sync`  may not be `true` if either `cancelable` or `stream` are `true`.</span></span>  |

## <a name="parameters-array"></a><span data-ttu-id="e9552-165">パラメータ配列</span><span class="sxs-lookup"><span data-stu-id="e9552-165">Parameters array</span></span>

<span data-ttu-id="e9552-166">`parameters` プロパティはオブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="e9552-166">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="e9552-167">各オブジェクトはそれぞれ 1 つのパラメータを表します。</span><span class="sxs-lookup"><span data-stu-id="e9552-167">Each of these objects represents a parameter.</span></span> <span data-ttu-id="e9552-168">次の表に、プロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="e9552-168">The following table contains its properties:</span></span>

|  <span data-ttu-id="e9552-169">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e9552-169">Property</span></span>  |  <span data-ttu-id="e9552-170">データ型</span><span class="sxs-lookup"><span data-stu-id="e9552-170">Data Type</span></span>  |  <span data-ttu-id="e9552-171">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="e9552-171">Required?</span></span>  |  <span data-ttu-id="e9552-172">説明</span><span class="sxs-lookup"><span data-stu-id="e9552-172">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="e9552-173">文字列</span><span class="sxs-lookup"><span data-stu-id="e9552-173">string</span></span>  |  <span data-ttu-id="e9552-174">いいえ</span><span class="sxs-lookup"><span data-stu-id="e9552-174">No</span></span> |  <span data-ttu-id="e9552-175">パラメータの説明。</span><span class="sxs-lookup"><span data-stu-id="e9552-175">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="e9552-176">文字列</span><span class="sxs-lookup"><span data-stu-id="e9552-176">string</span></span>  |  <span data-ttu-id="e9552-177">はい</span><span class="sxs-lookup"><span data-stu-id="e9552-177">Yes</span></span>  |  <span data-ttu-id="e9552-178">非配列値を意味する "scalar" か、行配列の配列を意味する "matrix" のどちらかでなければなりません。</span><span class="sxs-lookup"><span data-stu-id="e9552-178">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `name`  |  <span data-ttu-id="e9552-179">文字列</span><span class="sxs-lookup"><span data-stu-id="e9552-179">string</span></span>  |  <span data-ttu-id="e9552-180">はい</span><span class="sxs-lookup"><span data-stu-id="e9552-180">Yes</span></span>  |  <span data-ttu-id="e9552-181">パラメーターの名前です。</span><span class="sxs-lookup"><span data-stu-id="e9552-181">The name of the parameter.</span></span> <span data-ttu-id="e9552-182">この名前は Excel の IntelliSense で表示されます。</span><span class="sxs-lookup"><span data-stu-id="e9552-182">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="e9552-183">文字列</span><span class="sxs-lookup"><span data-stu-id="e9552-183">string</span></span>  |  <span data-ttu-id="e9552-184">はい</span><span class="sxs-lookup"><span data-stu-id="e9552-184">Yes</span></span>  |  <span data-ttu-id="e9552-185">パラメータのデータ型。</span><span class="sxs-lookup"><span data-stu-id="e9552-185">The data type of the parameter.</span></span> <span data-ttu-id="e9552-186">"boolean"、"number"、または "string" のいずれかです。</span><span class="sxs-lookup"><span data-stu-id="e9552-186">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result-object"></a><span data-ttu-id="e9552-187">結果オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e9552-187">Result object</span></span>

<span data-ttu-id="e9552-188">`results` プロパティは、関数から返された値に関するメタデータです。</span><span class="sxs-lookup"><span data-stu-id="e9552-188">The `results` property provides metadata about the value returned from the function.</span></span> <span data-ttu-id="e9552-189">次の表に、プロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="e9552-189">The following table contains its properties:</span></span>

|  <span data-ttu-id="e9552-190">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e9552-190">Property</span></span>  |  <span data-ttu-id="e9552-191">データ型</span><span class="sxs-lookup"><span data-stu-id="e9552-191">Data Type</span></span>  |  <span data-ttu-id="e9552-192">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="e9552-192">Required?</span></span>  |  <span data-ttu-id="e9552-193">説明</span><span class="sxs-lookup"><span data-stu-id="e9552-193">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="e9552-194">文字列</span><span class="sxs-lookup"><span data-stu-id="e9552-194">string</span></span>  |  <span data-ttu-id="e9552-195">いいえ</span><span class="sxs-lookup"><span data-stu-id="e9552-195">No</span></span>  |  <span data-ttu-id="e9552-196">非配列値を意味する "scalar" か、行配列の配列を意味する "matrix" のどちらかでなければなりません。</span><span class="sxs-lookup"><span data-stu-id="e9552-196">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `type`  |  <span data-ttu-id="e9552-197">文字列</span><span class="sxs-lookup"><span data-stu-id="e9552-197">string</span></span>  |  <span data-ttu-id="e9552-198">はい</span><span class="sxs-lookup"><span data-stu-id="e9552-198">Yes</span></span>  |  <span data-ttu-id="e9552-199">パラメータのデータ型。</span><span class="sxs-lookup"><span data-stu-id="e9552-199">The data type of the parameter.</span></span> <span data-ttu-id="e9552-200">"boolean"、"number"、または "string" のいずれかです。</span><span class="sxs-lookup"><span data-stu-id="e9552-200">Must be "boolean", "number", or "string".</span></span>  |

## <a name="example"></a><span data-ttu-id="e9552-201">例</span><span class="sxs-lookup"><span data-stu-id="e9552-201">Example</span></span>

<span data-ttu-id="e9552-202">次の JSON コードは、カスタム関数のメタデータ ファイルの例です。</span><span class="sxs-lookup"><span data-stu-id="e9552-202">The following JSON code is an example of a metadata file for custom functions.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="e9552-203">関連項目</span><span class="sxs-lookup"><span data-stu-id="e9552-203">See also</span></span>
[<span data-ttu-id="e9552-204">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="e9552-204">Custom functions</span></span>](custom-functions-overview.md)<br>
[<span data-ttu-id="e9552-205">配列数式のガイドラインと例</span><span class="sxs-lookup"><span data-stu-id="e9552-205">Guidelines and examples of array formulas</span></span>](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
