---
ms.date: 03/29/2019
description: Excel のカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 3703699348e99fd076fe0e3affac88038e3aaf59
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914257"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="0e7ee-103">カスタム関数のメタデータ (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="0e7ee-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="0e7ee-104">excel アドイン内で[カスタム関数](custom-functions-overview.md)を定義する場合、アドインプロジェクトには、カスタム関数を登録してエンドユーザーが使用できるようにするために excel が必要とする情報を提供する JSON メタデータファイルが含まれています。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="0e7ee-105">このファイルは、次のいずれかの方法で生成されます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-105">This file is generated either:</span></span>

- <span data-ttu-id="0e7ee-106">手書きの JSON ファイル</span><span class="sxs-lookup"><span data-stu-id="0e7ee-106">by you, in a handwritten JSON file</span></span>
- <span data-ttu-id="0e7ee-107">関数の先頭に入力した JSDoc コメントから</span><span class="sxs-lookup"><span data-stu-id="0e7ee-107">from the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="0e7ee-108">ユーザーが初めてアドインを実行したときに、すべてのブックの同じユーザーがそのアドインを使用できるようになると、カスタム関数が登録されます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="0e7ee-109">この記事では、JSON メタデータファイルの形式について説明しています (手動で記述する場合を想定しています)。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="0e7ee-110">JSDoc comment json ファイル生成の詳細については、「[カスタム関数の json メタデータの生成](custom-functions-json-autogeneration.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="0e7ee-111">カスタム関数を有効にするためにアドイン プロジェクトに含める必要のあるその他のファイルに関する情報については、「[Excel でのカスタム関数の作成](custom-functions-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> <span data-ttu-id="0e7ee-112">JSON ファイルをホストするサーバーでは、カスタム関数を Excel Online で正しく作動させるために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="0e7ee-113">メタデータの例</span><span class="sxs-lookup"><span data-stu-id="0e7ee-113">Example metadata</span></span>

<span data-ttu-id="0e7ee-114">次の例では、カスタム関数を定義するアドインの JSON メタデータ ファイルの内容を示します。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="0e7ee-115">この例の後に続くセクションでは、JSON の例に含まれる個々のプロパティの詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="0e7ee-116">完全な JSON ファイルのサンプルは、[OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub リポジトリにあります。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="0e7ee-117">functions</span><span class="sxs-lookup"><span data-stu-id="0e7ee-117">functions</span></span> 

<span data-ttu-id="0e7ee-118">`functions` プロパティは、カスタム関数オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-118">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="0e7ee-119">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-119">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="0e7ee-120">プロパティ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-120">Property</span></span>  |  <span data-ttu-id="0e7ee-121">データ型</span><span class="sxs-lookup"><span data-stu-id="0e7ee-121">Data type</span></span>  |  <span data-ttu-id="0e7ee-122">必須</span><span class="sxs-lookup"><span data-stu-id="0e7ee-122">Required</span></span>  |  <span data-ttu-id="0e7ee-123">説明</span><span class="sxs-lookup"><span data-stu-id="0e7ee-123">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="0e7ee-124">string</span><span class="sxs-lookup"><span data-stu-id="0e7ee-124">string</span></span>  |  <span data-ttu-id="0e7ee-125">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-125">No</span></span>  |  <span data-ttu-id="0e7ee-126">Excel でエンド ユーザーに表示される関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-126">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="0e7ee-127">たとえば、「**華氏の値を摂氏に変換する**」です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-127">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="0e7ee-128">string</span><span class="sxs-lookup"><span data-stu-id="0e7ee-128">string</span></span>  |   <span data-ttu-id="0e7ee-129">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-129">No</span></span>  |  <span data-ttu-id="0e7ee-130">関数に関する情報を提供する URL です </span><span class="sxs-lookup"><span data-stu-id="0e7ee-130">URL that provides information about the function.</span></span> <span data-ttu-id="0e7ee-131">(作業ウィンドウに表示されます)。たとえば、**http://contoso.com/help/convertcelsiustofahrenheit.html** です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-131">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="0e7ee-132">文字列</span><span class="sxs-lookup"><span data-stu-id="0e7ee-132">string</span></span> | <span data-ttu-id="0e7ee-133">はい</span><span class="sxs-lookup"><span data-stu-id="0e7ee-133">Yes</span></span> | <span data-ttu-id="0e7ee-134">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-134">A unique ID for the function.</span></span> <span data-ttu-id="0e7ee-135">この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-135">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="0e7ee-136">文字列</span><span class="sxs-lookup"><span data-stu-id="0e7ee-136">string</span></span>  |  <span data-ttu-id="0e7ee-137">はい</span><span class="sxs-lookup"><span data-stu-id="0e7ee-137">Yes</span></span>  |  <span data-ttu-id="0e7ee-138">Excel でエンド ユーザーに表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-138">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="0e7ee-139">Excel では、この関数名は XML マニフェスト ファイルで指定されているカスタム関数の名前空間でプレフィックスされます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-139">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="0e7ee-140">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0e7ee-140">object</span></span>  |  <span data-ttu-id="0e7ee-141">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-141">No</span></span>  |  <span data-ttu-id="0e7ee-142">Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-142">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="0e7ee-143">詳細については、[options](#options) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-143">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="0e7ee-144">配列</span><span class="sxs-lookup"><span data-stu-id="0e7ee-144">array</span></span>  |  <span data-ttu-id="0e7ee-145">はい</span><span class="sxs-lookup"><span data-stu-id="0e7ee-145">Yes</span></span>  |  <span data-ttu-id="0e7ee-146">関数の入力パラメーターを定義する配列です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-146">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="0e7ee-147">詳細については、[parameters](#parameters) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-147">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="0e7ee-148">object</span><span class="sxs-lookup"><span data-stu-id="0e7ee-148">object</span></span>  |  <span data-ttu-id="0e7ee-149">はい</span><span class="sxs-lookup"><span data-stu-id="0e7ee-149">Yes</span></span>  |  <span data-ttu-id="0e7ee-150">関数が返す情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-150">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="0e7ee-151">詳細については、[result](#result) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-151">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="0e7ee-152">options</span><span class="sxs-lookup"><span data-stu-id="0e7ee-152">options</span></span>

<span data-ttu-id="0e7ee-153">`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-153">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="0e7ee-154">次の表に、`options` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-154">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="0e7ee-155">プロパティ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-155">Property</span></span>  |  <span data-ttu-id="0e7ee-156">データ型</span><span class="sxs-lookup"><span data-stu-id="0e7ee-156">Data type</span></span>  |  <span data-ttu-id="0e7ee-157">必須</span><span class="sxs-lookup"><span data-stu-id="0e7ee-157">Required</span></span>  |  <span data-ttu-id="0e7ee-158">説明</span><span class="sxs-lookup"><span data-stu-id="0e7ee-158">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="0e7ee-159">ブール</span><span class="sxs-lookup"><span data-stu-id="0e7ee-159">boolean</span></span>  |  <span data-ttu-id="0e7ee-160">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-160">No</span></span><br/><br/><span data-ttu-id="0e7ee-161">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-161">Default value is `false`.</span></span>  |  <span data-ttu-id="0e7ee-162">`true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `onCanceled` ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-162">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="0e7ee-163">このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します </span><span class="sxs-lookup"><span data-stu-id="0e7ee-163">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="0e7ee-164">(このパラメーターを `parameters` プロパティには登録し***ない***でください)。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-164">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="0e7ee-165">この関数の本文では、ハンドラーを `caller.onCanceled` メンバーに割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-165">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="0e7ee-166">詳細については、「[関数をキャンセルする](custom-functions-web-reqs.md#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-166">For more information, see [Canceling a function](custom-functions-web-reqs.md#canceling-a-function).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="0e7ee-167">ブール</span><span class="sxs-lookup"><span data-stu-id="0e7ee-167">boolean</span></span> | <span data-ttu-id="0e7ee-168">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-168">No</span></span> <br/><br/><span data-ttu-id="0e7ee-169">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-169">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="0e7ee-170">true の場合、カスタム関数は、カスタム関数を呼び出したセルのアドレスにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-170">If true, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="0e7ee-171">カスタム関数を呼び出したセルのアドレスを取得するには、カスタム関数で context を使用します。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-171">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="0e7ee-172">詳しくは、「[カスタム関数が呼び出したセルを特定する](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-172">For more information, see [Determine which cell invoked your custom function](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span></span> <span data-ttu-id="0e7ee-173">カスタム関数は、streaming と requiresAddress の両方として設定することはできません。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-173">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="0e7ee-174">このオプションを使用する場合、' invocationContext ' パラメーターは、オプションで渡された最後のパラメーターである必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-174">When using this option, the 'invocationContext' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="0e7ee-175">ブール</span><span class="sxs-lookup"><span data-stu-id="0e7ee-175">boolean</span></span>  |  <span data-ttu-id="0e7ee-176">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-176">No</span></span><br/><br/><span data-ttu-id="0e7ee-177">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-177">Default value is `false`.</span></span>  |  <span data-ttu-id="0e7ee-178">`true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-178">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="0e7ee-179">このオプションは、株価などの急速に変化するデータ ソースに便利です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-179">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="0e7ee-180">このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します </span><span class="sxs-lookup"><span data-stu-id="0e7ee-180">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="0e7ee-181">(このパラメーターを `parameters` プロパティには登録し***ない***でください)。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-181">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="0e7ee-182">この関数には、`return` ステートメントは含めないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-182">The function should have no `return` statement.</span></span> <span data-ttu-id="0e7ee-183">代わりに、結果の値は `caller.setResult` コールバック メソッドの引数として渡されます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-183">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="0e7ee-184">詳細については、「[ストリーミング関数](custom-functions-web-reqs.md#streaming-functions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-184">For more information, see [Streaming functions](custom-functions-web-reqs.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="0e7ee-185">ブール</span><span class="sxs-lookup"><span data-stu-id="0e7ee-185">boolean</span></span> | <span data-ttu-id="0e7ee-186">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-186">No</span></span> <br/><br/><span data-ttu-id="0e7ee-187">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-187">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="0e7ee-188">`true` の場合は、数式の依存値が変更されたときのみではなく、Excel が再計算するたびに関数が再計算されます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-188">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="0e7ee-189">関数は、ストリーミングと揮発性の両方にすることはできません。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-189">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="0e7ee-190">`stream` と `volatile` の両方のプロパティが `true` に設定されている場合は、揮発性のオプションが無視されます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-190">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="0e7ee-191">parameters</span><span class="sxs-lookup"><span data-stu-id="0e7ee-191">parameters</span></span>

<span data-ttu-id="0e7ee-192">`parameters` プロパティは、パラメーター オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-192">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="0e7ee-193">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-193">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="0e7ee-194">プロパティ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-194">Property</span></span>  |  <span data-ttu-id="0e7ee-195">データ型</span><span class="sxs-lookup"><span data-stu-id="0e7ee-195">Data type</span></span>  |  <span data-ttu-id="0e7ee-196">必須</span><span class="sxs-lookup"><span data-stu-id="0e7ee-196">Required</span></span>  |  <span data-ttu-id="0e7ee-197">説明</span><span class="sxs-lookup"><span data-stu-id="0e7ee-197">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="0e7ee-198">string</span><span class="sxs-lookup"><span data-stu-id="0e7ee-198">string</span></span>  |  <span data-ttu-id="0e7ee-199">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-199">No</span></span> |  <span data-ttu-id="0e7ee-200">パラメーターの説明です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-200">A description of the parameter.</span></span> <span data-ttu-id="0e7ee-201">これは、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-201">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="0e7ee-202">string</span><span class="sxs-lookup"><span data-stu-id="0e7ee-202">string</span></span>  |  <span data-ttu-id="0e7ee-203">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-203">No</span></span>  |  <span data-ttu-id="0e7ee-204">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-204">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="0e7ee-205">文字列</span><span class="sxs-lookup"><span data-stu-id="0e7ee-205">string</span></span>  |  <span data-ttu-id="0e7ee-206">はい</span><span class="sxs-lookup"><span data-stu-id="0e7ee-206">Yes</span></span>  |  <span data-ttu-id="0e7ee-207">パラメーターの名前です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-207">The name of the parameter.</span></span> <span data-ttu-id="0e7ee-208">この名前は、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-208">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="0e7ee-209">文字列</span><span class="sxs-lookup"><span data-stu-id="0e7ee-209">string</span></span>  |  <span data-ttu-id="0e7ee-210">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-210">No</span></span>  |  <span data-ttu-id="0e7ee-211">パラメーターのデータ型を指定します。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-211">The data type of the parameter.</span></span> <span data-ttu-id="0e7ee-212">**boolean**、**number**、**string**、または **any** が可能です。ここでは、前の 3 種類のいずれかを使用できます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-212">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="0e7ee-213">このプロパティが指定されていない場合、データ型の既定は **any** です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-213">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="0e7ee-214">ブール</span><span class="sxs-lookup"><span data-stu-id="0e7ee-214">boolean</span></span> | <span data-ttu-id="0e7ee-215">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-215">No</span></span> | <span data-ttu-id="0e7ee-216">`true` の場合、パラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-216">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="0e7ee-217">省略可能なパラメーターの `type` プロパティが指定されていない場合や `any` に設定している場合は、Excel のセルに関数が入力されているときに、IDE の linting エラーや省略可能なパラメーターが表示されないなどの問題が発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-217">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="0e7ee-218">これについては、2018 年 12 月に変更される予定です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-218">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="0e7ee-219">result</span><span class="sxs-lookup"><span data-stu-id="0e7ee-219">result</span></span>

<span data-ttu-id="0e7ee-220">`result` オブジェクトは、この関数が返す情報の種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-220">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="0e7ee-221">次の表に、`result` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-221">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="0e7ee-222">プロパティ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-222">Property</span></span>  |  <span data-ttu-id="0e7ee-223">データ型</span><span class="sxs-lookup"><span data-stu-id="0e7ee-223">Data type</span></span>  |  <span data-ttu-id="0e7ee-224">必須</span><span class="sxs-lookup"><span data-stu-id="0e7ee-224">Required</span></span>  |  <span data-ttu-id="0e7ee-225">説明</span><span class="sxs-lookup"><span data-stu-id="0e7ee-225">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="0e7ee-226">string</span><span class="sxs-lookup"><span data-stu-id="0e7ee-226">string</span></span>  |  <span data-ttu-id="0e7ee-227">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-227">No</span></span>  |  <span data-ttu-id="0e7ee-228">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-228">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="0e7ee-229">文字列</span><span class="sxs-lookup"><span data-stu-id="0e7ee-229">string</span></span>  |  <span data-ttu-id="0e7ee-230">はい</span><span class="sxs-lookup"><span data-stu-id="0e7ee-230">Yes</span></span>  |  <span data-ttu-id="0e7ee-231">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-231">The data type of the parameter.</span></span> <span data-ttu-id="0e7ee-232">**boolean**、**number**、**string**、または **any** である必要があります。ここでは、前の 3 種類のいずれかを使用できます。</span><span class="sxs-lookup"><span data-stu-id="0e7ee-232">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="0e7ee-233">関連項目</span><span class="sxs-lookup"><span data-stu-id="0e7ee-233">See also</span></span>

* [<span data-ttu-id="0e7ee-234">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0e7ee-234">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="0e7ee-235">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="0e7ee-235">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="0e7ee-236">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="0e7ee-236">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="0e7ee-237">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="0e7ee-237">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="0e7ee-238">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="0e7ee-238">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
