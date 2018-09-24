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
# <a name="custom-functions-metadata"></a><span data-ttu-id="3bc73-103">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="3bc73-103">Custom functions metadata</span></span>

<span data-ttu-id="3bc73-104">Excel アドインで[カスタム関数](custom-functions-overview.md) を定義する場合には、Excel でカスタム関数を登録してエンドユーザーが使用できるようにするための情報を提供する JSON メタデータ ファイルを、アドイン プロジェクトに含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="3bc73-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end-users.</span></span> <span data-ttu-id="3bc73-105">この記事では、JSON メタデータ ファイルの形式について説明します。</span><span class="sxs-lookup"><span data-stu-id="3bc73-105">This article describes the format of the JSON file with examples.</span></span>

> [!NOTE]
> <span data-ttu-id="3bc73-106">カスタム関数を有効にするためにアドイン プロジェクトに含める必要のある、その他のファイルに関する情報については、「[Excel でカスタム関数を作成する](custom-functions-overview.md#learn-the-basics)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3bc73-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md#learn-the-basics).</span></span>

## <a name="example-metadata"></a><span data-ttu-id="3bc73-107">メタデータの例</span><span class="sxs-lookup"><span data-stu-id="3bc73-107">Example metadata</span></span>

<span data-ttu-id="3bc73-108">次の例は、カスタム関数を定義するアドイン用の JSON メタデータ ファイルの内容を示しています。</span><span class="sxs-lookup"><span data-stu-id="3bc73-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="3bc73-109">この例に続くセクションでは、この JSON の例の中にある個々のプロパティについての詳細情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="3bc73-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="3bc73-110">JSON ファイルの完全なサンプルは、「[OfficeDev/Excel-Custom-Functions GitHub リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json)」で利用可能です。</span><span class="sxs-lookup"><span data-stu-id="3bc73-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="3bc73-111">functions</span><span class="sxs-lookup"><span data-stu-id="3bc73-111">functions</span></span> 

<span data-ttu-id="3bc73-112"> `functions` プロパティは、カスタム関数オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="3bc73-112">The `functions` property is an array of objects.</span></span> <span data-ttu-id="3bc73-113">次の表で、各オブジェクトのプロパティを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="3bc73-113">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="3bc73-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3bc73-114">Property</span></span>  |  <span data-ttu-id="3bc73-115">データ型</span><span class="sxs-lookup"><span data-stu-id="3bc73-115">Data type</span></span>  |  <span data-ttu-id="3bc73-116">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="3bc73-116">Required</span></span>  |  <span data-ttu-id="3bc73-117">説明</span><span class="sxs-lookup"><span data-stu-id="3bc73-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="3bc73-118">string</span><span class="sxs-lookup"><span data-stu-id="3bc73-118">string</span></span>  |  <span data-ttu-id="3bc73-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="3bc73-119">No</span></span>  |  <span data-ttu-id="3bc73-120">Excel UI に表示される関数の説明。</span><span class="sxs-lookup"><span data-stu-id="3bc73-120">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="3bc73-121">たとえば、「**摂氏の値を華氏に変換します**」など。</span><span class="sxs-lookup"><span data-stu-id="3bc73-121">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="3bc73-122">string</span><span class="sxs-lookup"><span data-stu-id="3bc73-122">string</span></span>  |   <span data-ttu-id="3bc73-123">いいえ</span><span class="sxs-lookup"><span data-stu-id="3bc73-123">No</span></span>  |  <span data-ttu-id="3bc73-124">ユーザーが関数についての情報を見ることができる URL です。</span><span class="sxs-lookup"><span data-stu-id="3bc73-124">URL where your users can get help about the function.</span></span> <span data-ttu-id="3bc73-125">(作業ウィンドウに表示されます。) たとえば、 **http://contoso.com/help/convertcelsiustofahrenheit.html**。</span><span class="sxs-lookup"><span data-stu-id="3bc73-125">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span> |
| `id`     | <span data-ttu-id="3bc73-126">string</span><span class="sxs-lookup"><span data-stu-id="3bc73-126">string</span></span> | <span data-ttu-id="3bc73-127">はい</span><span class="sxs-lookup"><span data-stu-id="3bc73-127">Yes</span></span> | <span data-ttu-id="3bc73-128">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="3bc73-128">A unique ID for the group.</span></span> <span data-ttu-id="3bc73-129">設定後は、この ID は変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="3bc73-129">This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="3bc73-130">string</span><span class="sxs-lookup"><span data-stu-id="3bc73-130">string</span></span>  |  <span data-ttu-id="3bc73-131">はい</span><span class="sxs-lookup"><span data-stu-id="3bc73-131">Yes</span></span>  |  <span data-ttu-id="3bc73-132">ユーザーが関数を選択しているときに Excel の UI の (名前空間の先頭に) 表示される関数の名前。</span><span class="sxs-lookup"><span data-stu-id="3bc73-132">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="3bc73-133">JavaScript で定義されているものと同じ関数の名前である必要はありません。</span><span class="sxs-lookup"><span data-stu-id="3bc73-133">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="3bc73-134">object</span><span class="sxs-lookup"><span data-stu-id="3bc73-134">object</span></span>  |  <span data-ttu-id="3bc73-135">いいえ</span><span class="sxs-lookup"><span data-stu-id="3bc73-135">No</span></span>  |  <span data-ttu-id="3bc73-136">Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。</span><span class="sxs-lookup"><span data-stu-id="3bc73-136">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="3bc73-137">詳細については、「[オプション オブジェクト](#options-object)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3bc73-137">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="3bc73-138">array</span><span class="sxs-lookup"><span data-stu-id="3bc73-138">array</span></span>  |  <span data-ttu-id="3bc73-139">はい</span><span class="sxs-lookup"><span data-stu-id="3bc73-139">Yes</span></span>  |  <span data-ttu-id="3bc73-140">関数の入力パラメーターを定義する配列です。</span><span class="sxs-lookup"><span data-stu-id="3bc73-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="3bc73-141">詳細については、「[パラメーター配列](#parameters-array)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3bc73-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="3bc73-142">object</span><span class="sxs-lookup"><span data-stu-id="3bc73-142">object</span></span>  |  <span data-ttu-id="3bc73-143">はい</span><span class="sxs-lookup"><span data-stu-id="3bc73-143">Yes</span></span>  |  <span data-ttu-id="3bc73-144">関数によって返される情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="3bc73-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="3bc73-145">詳細については、「[結果オブジェクト](#result-object)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3bc73-145">See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="3bc73-146">options</span><span class="sxs-lookup"><span data-stu-id="3bc73-146">options</span></span>

<span data-ttu-id="3bc73-147">`options` オブジェクトを使用すると、Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。</span><span class="sxs-lookup"><span data-stu-id="3bc73-147">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="3bc73-148">次の表で、`options` オブジェクトのプロパティを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="3bc73-148">The following table lists the parts of the `options` claim.</span></span>

|  <span data-ttu-id="3bc73-149">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3bc73-149">Property</span></span>  |  <span data-ttu-id="3bc73-150">データ型</span><span class="sxs-lookup"><span data-stu-id="3bc73-150">Data type</span></span>  |  <span data-ttu-id="3bc73-151">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="3bc73-151">Required</span></span>  |  <span data-ttu-id="3bc73-152">説明</span><span class="sxs-lookup"><span data-stu-id="3bc73-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="3bc73-153">boolean</span><span class="sxs-lookup"><span data-stu-id="3bc73-153">boolean</span></span>  |  <span data-ttu-id="3bc73-154">いいえ。既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="3bc73-154">No, default is `false`.</span></span>  |  <span data-ttu-id="3bc73-155">の場合、Excel はユーザーが関数をキャンセルする操作をするたびに `onCanceled` ハンドラを呼び出します。たとえば、手動で再計算をトリガするか、関数が参照するセルを編集する場合です。`true`</span><span class="sxs-lookup"><span data-stu-id="3bc73-155">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="3bc73-156">このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="3bc73-156">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="3bc73-157">(このパラメータを `parameters` プロパティに登録***しない***でください)。</span><span class="sxs-lookup"><span data-stu-id="3bc73-157">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="3bc73-158">関数の本体では、ハンドラは `caller.onCanceled` メンバーに割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="3bc73-158">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="3bc73-159">詳細については、 「[関数をキャンセルする](custom-functions-overview.md#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3bc73-159">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="3bc73-160">boolean</span><span class="sxs-lookup"><span data-stu-id="3bc73-160">boolean</span></span>  |  <span data-ttu-id="3bc73-161">いいえ。既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="3bc73-161">No, default is `false`.</span></span>  |  <span data-ttu-id="3bc73-162">の場合、関数は一度の呼び出しで繰り返しセルに出力できます。`true`</span><span class="sxs-lookup"><span data-stu-id="3bc73-162">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="3bc73-163">このオプションは、株価など急激に変化するデータソースで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="3bc73-163">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="3bc73-164">このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="3bc73-164">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="3bc73-165">(このパラメーターを `parameters` プロパティに登録***しない***でください)。</span><span class="sxs-lookup"><span data-stu-id="3bc73-165">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="3bc73-166">関数では、`return` 文を使いません。</span><span class="sxs-lookup"><span data-stu-id="3bc73-166">The function should have no `return` statement.</span></span> <span data-ttu-id="3bc73-167">代わりに、結果値を `caller.setResult` コールバック メソッドの引数として渡します。</span><span class="sxs-lookup"><span data-stu-id="3bc73-167">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="3bc73-168">詳細については、「[ストリーム関数](custom-functions-overview.md#streamed-functions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3bc73-168">For more information, see [Excel functions by category](custom-functions-overview.md#streamed-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="3bc73-169">parameters</span><span class="sxs-lookup"><span data-stu-id="3bc73-169">parameters</span></span>

<span data-ttu-id="3bc73-170">`parameters` プロパティは、パラメーター オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="3bc73-170">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="3bc73-171">次の表で、各オブジェクトのプロパティを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="3bc73-171">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="3bc73-172">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3bc73-172">Property</span></span>  |  <span data-ttu-id="3bc73-173">データ型</span><span class="sxs-lookup"><span data-stu-id="3bc73-173">Data type</span></span>  |  <span data-ttu-id="3bc73-174">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="3bc73-174">Required</span></span>  |  <span data-ttu-id="3bc73-175">説明</span><span class="sxs-lookup"><span data-stu-id="3bc73-175">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="3bc73-176">string</span><span class="sxs-lookup"><span data-stu-id="3bc73-176">string</span></span>  |  <span data-ttu-id="3bc73-177">いいえ</span><span class="sxs-lookup"><span data-stu-id="3bc73-177">No</span></span> |  <span data-ttu-id="3bc73-178">パラメータの説明。</span><span class="sxs-lookup"><span data-stu-id="3bc73-178">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="3bc73-179">string</span><span class="sxs-lookup"><span data-stu-id="3bc73-179">string</span></span>  |  <span data-ttu-id="3bc73-180">いいえ</span><span class="sxs-lookup"><span data-stu-id="3bc73-180">No</span></span>  |  <span data-ttu-id="3bc73-181"> *\*scholar** (非配列値) または *\*matrix** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="3bc73-181">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="3bc73-182">string</span><span class="sxs-lookup"><span data-stu-id="3bc73-182">string</span></span>  |  <span data-ttu-id="3bc73-183">はい</span><span class="sxs-lookup"><span data-stu-id="3bc73-183">Yes</span></span>  |  <span data-ttu-id="3bc73-184">パラメーターの名前です。</span><span class="sxs-lookup"><span data-stu-id="3bc73-184">The name of the parameter.</span></span> <span data-ttu-id="3bc73-185">この名前は Excel の IntelliSense で表示されます。</span><span class="sxs-lookup"><span data-stu-id="3bc73-185">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="3bc73-186">string</span><span class="sxs-lookup"><span data-stu-id="3bc73-186">string</span></span>  |  <span data-ttu-id="3bc73-187">いいえ</span><span class="sxs-lookup"><span data-stu-id="3bc73-187">No</span></span>  |  <span data-ttu-id="3bc73-188">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="3bc73-188">The data type of the parameter.</span></span> <span data-ttu-id="3bc73-189"> *\*ブール値*\*、 *\*数値*\*、または *\*文字列*\*である必要があります。</span><span class="sxs-lookup"><span data-stu-id="3bc73-189">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result"></a><span data-ttu-id="3bc73-190">result</span><span class="sxs-lookup"><span data-stu-id="3bc73-190">result</span></span>

<span data-ttu-id="3bc73-191">関数によって返される情報の種類を定義する`results` オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="3bc73-191">The `results` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="3bc73-192">次の表で、`result` オブジェクトのプロパティを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="3bc73-192">The following table lists the parts of the `result` claim.</span></span>

|  <span data-ttu-id="3bc73-193">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3bc73-193">Property</span></span>  |  <span data-ttu-id="3bc73-194">データ型</span><span class="sxs-lookup"><span data-stu-id="3bc73-194">Data type</span></span>  |  <span data-ttu-id="3bc73-195">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="3bc73-195">Required</span></span>  |  <span data-ttu-id="3bc73-196">説明</span><span class="sxs-lookup"><span data-stu-id="3bc73-196">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="3bc73-197">string</span><span class="sxs-lookup"><span data-stu-id="3bc73-197">string</span></span>  |  <span data-ttu-id="3bc73-198">いいえ</span><span class="sxs-lookup"><span data-stu-id="3bc73-198">No</span></span>  |  <span data-ttu-id="3bc73-199"> *\*scholar** (非配列値) または *\*matrix** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="3bc73-199">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="3bc73-200">string</span><span class="sxs-lookup"><span data-stu-id="3bc73-200">string</span></span>  |  <span data-ttu-id="3bc73-201">はい</span><span class="sxs-lookup"><span data-stu-id="3bc73-201">Yes</span></span>  |  <span data-ttu-id="3bc73-202">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="3bc73-202">The data type of the parameter.</span></span> <span data-ttu-id="3bc73-203"> *\*ブール値*\*、 *\*数値*\*、または *\*文字列*\*である必要があります。</span><span class="sxs-lookup"><span data-stu-id="3bc73-203">Must be "boolean", "number", or "string".</span></span>  |

## <a name="see-also"></a><span data-ttu-id="3bc73-204">関連項目</span><span class="sxs-lookup"><span data-stu-id="3bc73-204">See also</span></span>

* [<span data-ttu-id="3bc73-205">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="3bc73-205">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="3bc73-206">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="3bc73-206">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="3bc73-207">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="3bc73-207">Custom functions best practices</span></span>](custom-functions-best-practices.md)