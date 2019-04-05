---
ms.date: 03/29/2019
description: Excel のカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 28a9a0207f7439af164eb9ca7c4b9ed9e966b3ed
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477552"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="45050-103">カスタム関数のメタデータ (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="45050-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="45050-104">excel アドイン内で[カスタム関数](custom-functions-overview.md)を定義する場合、アドインプロジェクトには、カスタム関数を登録してエンドユーザーが使用できるようにするために excel が必要とする情報を提供する JSON メタデータファイルが含まれています。</span><span class="sxs-lookup"><span data-stu-id="45050-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="45050-105">このファイルは、次のいずれかの方法で生成されます。</span><span class="sxs-lookup"><span data-stu-id="45050-105">This file is generated either:</span></span>

- <span data-ttu-id="45050-106">手書きの JSON ファイル</span><span class="sxs-lookup"><span data-stu-id="45050-106">by you, in a handwritten JSON file</span></span>
- <span data-ttu-id="45050-107">関数の先頭に入力した JSDoc コメントから</span><span class="sxs-lookup"><span data-stu-id="45050-107">from the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="45050-108">ユーザーが初めてアドインを実行したときに、すべてのブックの同じユーザーがそのアドインを使用できるようになると、カスタム関数が登録されます。</span><span class="sxs-lookup"><span data-stu-id="45050-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="45050-109">この記事では、JSON メタデータファイルの形式について説明しています (手動で記述する場合を想定しています)。</span><span class="sxs-lookup"><span data-stu-id="45050-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="45050-110">JSDoc comment json ファイル生成の詳細については、「[カスタム関数の json メタデータの生成](custom-functions-json-autogeneration.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="45050-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="45050-111">カスタム関数を有効にするためにアドイン プロジェクトに含める必要のあるその他のファイルに関する情報については、「[Excel でのカスタム関数の作成](custom-functions-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="45050-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> <span data-ttu-id="45050-112">JSON ファイルをホストするサーバーでは、カスタム関数を Excel Online で正しく作動させるために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="45050-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="45050-113">メタデータの例</span><span class="sxs-lookup"><span data-stu-id="45050-113">Example metadata</span></span>

<span data-ttu-id="45050-114">次の例では、カスタム関数を定義するアドインの JSON メタデータ ファイルの内容を示します。</span><span class="sxs-lookup"><span data-stu-id="45050-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="45050-115">この例の後に続くセクションでは、JSON の例に含まれる個々のプロパティの詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="45050-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="45050-116">完全な JSON ファイルのサンプルは、[OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub リポジトリにあります。</span><span class="sxs-lookup"><span data-stu-id="45050-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="45050-117">functions</span><span class="sxs-lookup"><span data-stu-id="45050-117">functions</span></span> 

<span data-ttu-id="45050-118">`functions` プロパティは、カスタム関数オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="45050-118">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="45050-119">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="45050-119">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="45050-120">プロパティ</span><span class="sxs-lookup"><span data-stu-id="45050-120">Property</span></span>  |  <span data-ttu-id="45050-121">データ型</span><span class="sxs-lookup"><span data-stu-id="45050-121">Data type</span></span>  |  <span data-ttu-id="45050-122">必須</span><span class="sxs-lookup"><span data-stu-id="45050-122">Required</span></span>  |  <span data-ttu-id="45050-123">説明</span><span class="sxs-lookup"><span data-stu-id="45050-123">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="45050-124">string</span><span class="sxs-lookup"><span data-stu-id="45050-124">string</span></span>  |  <span data-ttu-id="45050-125">いいえ</span><span class="sxs-lookup"><span data-stu-id="45050-125">No</span></span>  |  <span data-ttu-id="45050-126">Excel でエンド ユーザーに表示される関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="45050-126">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="45050-127">たとえば、「**華氏の値を摂氏に変換する**」です。</span><span class="sxs-lookup"><span data-stu-id="45050-127">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="45050-128">string</span><span class="sxs-lookup"><span data-stu-id="45050-128">string</span></span>  |   <span data-ttu-id="45050-129">いいえ</span><span class="sxs-lookup"><span data-stu-id="45050-129">No</span></span>  |  <span data-ttu-id="45050-130">関数に関する情報を提供する URL です </span><span class="sxs-lookup"><span data-stu-id="45050-130">URL that provides information about the function.</span></span> <span data-ttu-id="45050-131">(作業ウィンドウに表示されます)。たとえば、**http://contoso.com/help/convertcelsiustofahrenheit.html** です。</span><span class="sxs-lookup"><span data-stu-id="45050-131">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="45050-132">文字列</span><span class="sxs-lookup"><span data-stu-id="45050-132">string</span></span> | <span data-ttu-id="45050-133">はい</span><span class="sxs-lookup"><span data-stu-id="45050-133">Yes</span></span> | <span data-ttu-id="45050-134">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="45050-134">A unique ID for the function.</span></span> <span data-ttu-id="45050-135">この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。</span><span class="sxs-lookup"><span data-stu-id="45050-135">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="45050-136">文字列</span><span class="sxs-lookup"><span data-stu-id="45050-136">string</span></span>  |  <span data-ttu-id="45050-137">はい</span><span class="sxs-lookup"><span data-stu-id="45050-137">Yes</span></span>  |  <span data-ttu-id="45050-138">Excel でエンド ユーザーに表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="45050-138">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="45050-139">Excel では、この関数名は XML マニフェスト ファイルで指定されているカスタム関数の名前空間でプレフィックスされます。</span><span class="sxs-lookup"><span data-stu-id="45050-139">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="45050-140">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="45050-140">object</span></span>  |  <span data-ttu-id="45050-141">いいえ</span><span class="sxs-lookup"><span data-stu-id="45050-141">No</span></span>  |  <span data-ttu-id="45050-142">Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="45050-142">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="45050-143">詳細については、[options](#options) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="45050-143">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="45050-144">配列</span><span class="sxs-lookup"><span data-stu-id="45050-144">array</span></span>  |  <span data-ttu-id="45050-145">はい</span><span class="sxs-lookup"><span data-stu-id="45050-145">Yes</span></span>  |  <span data-ttu-id="45050-146">関数の入力パラメーターを定義する配列です。</span><span class="sxs-lookup"><span data-stu-id="45050-146">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="45050-147">詳細については、[parameters](#parameters) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="45050-147">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="45050-148">object</span><span class="sxs-lookup"><span data-stu-id="45050-148">object</span></span>  |  <span data-ttu-id="45050-149">はい</span><span class="sxs-lookup"><span data-stu-id="45050-149">Yes</span></span>  |  <span data-ttu-id="45050-150">関数が返す情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="45050-150">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="45050-151">詳細については、[result](#result) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="45050-151">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="45050-152">options</span><span class="sxs-lookup"><span data-stu-id="45050-152">options</span></span>

<span data-ttu-id="45050-153">`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="45050-153">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="45050-154">次の表に、`options` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="45050-154">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="45050-155">プロパティ</span><span class="sxs-lookup"><span data-stu-id="45050-155">Property</span></span>  |  <span data-ttu-id="45050-156">データ型</span><span class="sxs-lookup"><span data-stu-id="45050-156">Data type</span></span>  |  <span data-ttu-id="45050-157">必須</span><span class="sxs-lookup"><span data-stu-id="45050-157">Required</span></span>  |  <span data-ttu-id="45050-158">説明</span><span class="sxs-lookup"><span data-stu-id="45050-158">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="45050-159">ブール</span><span class="sxs-lookup"><span data-stu-id="45050-159">boolean</span></span>  |  <span data-ttu-id="45050-160">いいえ</span><span class="sxs-lookup"><span data-stu-id="45050-160">No</span></span><br/><br/><span data-ttu-id="45050-161">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="45050-161">Default value is `false`.</span></span>  |  <span data-ttu-id="45050-162">`true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `onCanceled` ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="45050-162">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="45050-163">このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します </span><span class="sxs-lookup"><span data-stu-id="45050-163">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="45050-164">(このパラメーターを `parameters` プロパティには登録し***ない***でください)。</span><span class="sxs-lookup"><span data-stu-id="45050-164">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="45050-165">この関数の本文では、ハンドラーを `caller.onCanceled` メンバーに割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="45050-165">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="45050-166">詳細については、「[関数をキャンセルする](custom-functions-web-reqs.md#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="45050-166">For more information, see [Canceling a function](custom-functions-web-reqs.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="45050-167">ブール</span><span class="sxs-lookup"><span data-stu-id="45050-167">boolean</span></span>  |  <span data-ttu-id="45050-168">いいえ</span><span class="sxs-lookup"><span data-stu-id="45050-168">No</span></span><br/><br/><span data-ttu-id="45050-169">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="45050-169">Default value is `false`.</span></span>  |  <span data-ttu-id="45050-170">`true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。</span><span class="sxs-lookup"><span data-stu-id="45050-170">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="45050-171">このオプションは、株価などの急速に変化するデータ ソースに便利です。</span><span class="sxs-lookup"><span data-stu-id="45050-171">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="45050-172">このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します </span><span class="sxs-lookup"><span data-stu-id="45050-172">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="45050-173">(このパラメーターを `parameters` プロパティには登録し***ない***でください)。</span><span class="sxs-lookup"><span data-stu-id="45050-173">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="45050-174">この関数には、`return` ステートメントは含めないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="45050-174">The function should have no `return` statement.</span></span> <span data-ttu-id="45050-175">代わりに、結果の値は `caller.setResult` コールバック メソッドの引数として渡されます。</span><span class="sxs-lookup"><span data-stu-id="45050-175">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="45050-176">詳細については、「[ストリーミング関数](custom-functions-web-reqs.md#streaming-functions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="45050-176">For more information, see [Streaming functions](custom-functions-web-reqs.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="45050-177">ブール</span><span class="sxs-lookup"><span data-stu-id="45050-177">boolean</span></span> | <span data-ttu-id="45050-178">いいえ</span><span class="sxs-lookup"><span data-stu-id="45050-178">No</span></span> <br/><br/><span data-ttu-id="45050-179">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="45050-179">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="45050-180">`true` の場合は、数式の依存値が変更されたときのみではなく、Excel が再計算するたびに関数が再計算されます。</span><span class="sxs-lookup"><span data-stu-id="45050-180">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="45050-181">関数は、ストリーミングと揮発性の両方にすることはできません。</span><span class="sxs-lookup"><span data-stu-id="45050-181">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="45050-182">`stream` と `volatile` の両方のプロパティが `true` に設定されている場合は、揮発性のオプションが無視されます。</span><span class="sxs-lookup"><span data-stu-id="45050-182">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="45050-183">parameters</span><span class="sxs-lookup"><span data-stu-id="45050-183">parameters</span></span>

<span data-ttu-id="45050-184">`parameters` プロパティは、パラメーター オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="45050-184">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="45050-185">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="45050-185">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="45050-186">プロパティ</span><span class="sxs-lookup"><span data-stu-id="45050-186">Property</span></span>  |  <span data-ttu-id="45050-187">データ型</span><span class="sxs-lookup"><span data-stu-id="45050-187">Data type</span></span>  |  <span data-ttu-id="45050-188">必須</span><span class="sxs-lookup"><span data-stu-id="45050-188">Required</span></span>  |  <span data-ttu-id="45050-189">説明</span><span class="sxs-lookup"><span data-stu-id="45050-189">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="45050-190">string</span><span class="sxs-lookup"><span data-stu-id="45050-190">string</span></span>  |  <span data-ttu-id="45050-191">いいえ</span><span class="sxs-lookup"><span data-stu-id="45050-191">No</span></span> |  <span data-ttu-id="45050-192">パラメーターの説明です。</span><span class="sxs-lookup"><span data-stu-id="45050-192">A description of the parameter.</span></span> <span data-ttu-id="45050-193">これは、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="45050-193">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="45050-194">string</span><span class="sxs-lookup"><span data-stu-id="45050-194">string</span></span>  |  <span data-ttu-id="45050-195">いいえ</span><span class="sxs-lookup"><span data-stu-id="45050-195">No</span></span>  |  <span data-ttu-id="45050-196">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="45050-196">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="45050-197">文字列</span><span class="sxs-lookup"><span data-stu-id="45050-197">string</span></span>  |  <span data-ttu-id="45050-198">はい</span><span class="sxs-lookup"><span data-stu-id="45050-198">Yes</span></span>  |  <span data-ttu-id="45050-199">パラメーターの名前です。</span><span class="sxs-lookup"><span data-stu-id="45050-199">The name of the parameter.</span></span> <span data-ttu-id="45050-200">この名前は、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="45050-200">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="45050-201">文字列</span><span class="sxs-lookup"><span data-stu-id="45050-201">string</span></span>  |  <span data-ttu-id="45050-202">いいえ</span><span class="sxs-lookup"><span data-stu-id="45050-202">No</span></span>  |  <span data-ttu-id="45050-203">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="45050-203">The data type of the parameter.</span></span> <span data-ttu-id="45050-204">**boolean**、**number**、**string**、または **any** が可能です。ここでは、前の 3 種類のいずれかを使用できます。</span><span class="sxs-lookup"><span data-stu-id="45050-204">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="45050-205">このプロパティが指定されていない場合、データ型の既定は **any** です。</span><span class="sxs-lookup"><span data-stu-id="45050-205">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="45050-206">ブール</span><span class="sxs-lookup"><span data-stu-id="45050-206">boolean</span></span> | <span data-ttu-id="45050-207">いいえ</span><span class="sxs-lookup"><span data-stu-id="45050-207">No</span></span> | <span data-ttu-id="45050-208">`true` の場合、パラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="45050-208">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="45050-209">省略可能なパラメーターの `type` プロパティが指定されていない場合や `any` に設定している場合は、Excel のセルに関数が入力されているときに、IDE の linting エラーや省略可能なパラメーターが表示されないなどの問題が発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="45050-209">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="45050-210">これについては、2018 年 12 月に変更される予定です。</span><span class="sxs-lookup"><span data-stu-id="45050-210">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="45050-211">result</span><span class="sxs-lookup"><span data-stu-id="45050-211">result</span></span>

<span data-ttu-id="45050-212">`result` オブジェクトは、この関数が返す情報の種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="45050-212">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="45050-213">次の表に、`result` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="45050-213">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="45050-214">プロパティ</span><span class="sxs-lookup"><span data-stu-id="45050-214">Property</span></span>  |  <span data-ttu-id="45050-215">データ型</span><span class="sxs-lookup"><span data-stu-id="45050-215">Data type</span></span>  |  <span data-ttu-id="45050-216">必須</span><span class="sxs-lookup"><span data-stu-id="45050-216">Required</span></span>  |  <span data-ttu-id="45050-217">説明</span><span class="sxs-lookup"><span data-stu-id="45050-217">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="45050-218">string</span><span class="sxs-lookup"><span data-stu-id="45050-218">string</span></span>  |  <span data-ttu-id="45050-219">いいえ</span><span class="sxs-lookup"><span data-stu-id="45050-219">No</span></span>  |  <span data-ttu-id="45050-220">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="45050-220">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="45050-221">文字列</span><span class="sxs-lookup"><span data-stu-id="45050-221">string</span></span>  |  <span data-ttu-id="45050-222">はい</span><span class="sxs-lookup"><span data-stu-id="45050-222">Yes</span></span>  |  <span data-ttu-id="45050-223">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="45050-223">The data type of the parameter.</span></span> <span data-ttu-id="45050-224">**boolean**、**number**、**string**、または **any** である必要があります。ここでは、前の 3 種類のいずれかを使用できます。</span><span class="sxs-lookup"><span data-stu-id="45050-224">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="45050-225">関連項目</span><span class="sxs-lookup"><span data-stu-id="45050-225">See also</span></span>

* [<span data-ttu-id="45050-226">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="45050-226">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="45050-227">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="45050-227">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="45050-228">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="45050-228">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="45050-229">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="45050-229">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="45050-230">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="45050-230">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
