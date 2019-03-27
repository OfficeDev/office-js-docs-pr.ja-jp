---
ms.date: 01/08/2019
description: Excel のカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 43ec436d15d118346bb04dcd4d16f5eb180ecbd3
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872089"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="1aeed-103">カスタム関数のメタデータ (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="1aeed-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="1aeed-104">Excel アドイン内に[カスタム関数](custom-functions-overview.md)を定義する場合、カスタム関数を登録し、エンド ユーザーが利用できるようにするために Excel が必要とする情報を提供する JSON メタデータ ファイルをアドイン プロジェクトに含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="1aeed-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="1aeed-105">この記事では、その JSON メタデータ ファイルの形式について説明します。</span><span class="sxs-lookup"><span data-stu-id="1aeed-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="1aeed-106">カスタム関数を有効にするためにアドイン プロジェクトに含める必要のあるその他のファイルに関する情報については、「[Excel でのカスタム関数の作成](custom-functions-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1aeed-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="1aeed-107">メタデータの例</span><span class="sxs-lookup"><span data-stu-id="1aeed-107">Example metadata</span></span>

<span data-ttu-id="1aeed-108">次の例では、カスタム関数を定義するアドインの JSON メタデータ ファイルの内容を示します。</span><span class="sxs-lookup"><span data-stu-id="1aeed-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="1aeed-109">この例の後に続くセクションでは、JSON の例に含まれる個々のプロパティの詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="1aeed-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="1aeed-110">完全な JSON ファイルのサンプルは、[OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub リポジトリにあります。</span><span class="sxs-lookup"><span data-stu-id="1aeed-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="1aeed-111">functions</span><span class="sxs-lookup"><span data-stu-id="1aeed-111">functions</span></span> 

<span data-ttu-id="1aeed-112">`functions` プロパティは、カスタム関数オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-112">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="1aeed-113">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="1aeed-113">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="1aeed-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="1aeed-114">Property</span></span>  |  <span data-ttu-id="1aeed-115">データ型</span><span class="sxs-lookup"><span data-stu-id="1aeed-115">Data type</span></span>  |  <span data-ttu-id="1aeed-116">必須</span><span class="sxs-lookup"><span data-stu-id="1aeed-116">Required</span></span>  |  <span data-ttu-id="1aeed-117">説明</span><span class="sxs-lookup"><span data-stu-id="1aeed-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="1aeed-118">string</span><span class="sxs-lookup"><span data-stu-id="1aeed-118">string</span></span>  |  <span data-ttu-id="1aeed-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="1aeed-119">No</span></span>  |  <span data-ttu-id="1aeed-120">Excel でエンド ユーザーに表示される関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="1aeed-121">たとえば、「**華氏の値を摂氏に変換する**」です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="1aeed-122">string</span><span class="sxs-lookup"><span data-stu-id="1aeed-122">string</span></span>  |   <span data-ttu-id="1aeed-123">いいえ</span><span class="sxs-lookup"><span data-stu-id="1aeed-123">No</span></span>  |  <span data-ttu-id="1aeed-124">関数に関する情報を提供する URL です </span><span class="sxs-lookup"><span data-stu-id="1aeed-124">URL that provides information about the function.</span></span> <span data-ttu-id="1aeed-125">(作業ウィンドウに表示されます)。たとえば、**http://contoso.com/help/convertcelsiustofahrenheit.html** です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="1aeed-126">文字列</span><span class="sxs-lookup"><span data-stu-id="1aeed-126">string</span></span> | <span data-ttu-id="1aeed-127">はい</span><span class="sxs-lookup"><span data-stu-id="1aeed-127">Yes</span></span> | <span data-ttu-id="1aeed-128">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-128">A unique ID for the function.</span></span> <span data-ttu-id="1aeed-129">この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。</span><span class="sxs-lookup"><span data-stu-id="1aeed-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="1aeed-130">文字列</span><span class="sxs-lookup"><span data-stu-id="1aeed-130">string</span></span>  |  <span data-ttu-id="1aeed-131">はい</span><span class="sxs-lookup"><span data-stu-id="1aeed-131">Yes</span></span>  |  <span data-ttu-id="1aeed-132">Excel でエンド ユーザーに表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="1aeed-133">Excel では、この関数名は XML マニフェスト ファイルで指定されているカスタム関数の名前空間でプレフィックスされます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-133">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="1aeed-134">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1aeed-134">object</span></span>  |  <span data-ttu-id="1aeed-135">いいえ</span><span class="sxs-lookup"><span data-stu-id="1aeed-135">No</span></span>  |  <span data-ttu-id="1aeed-136">Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-136">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="1aeed-137">詳細については、[options](#options) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1aeed-137">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="1aeed-138">配列</span><span class="sxs-lookup"><span data-stu-id="1aeed-138">array</span></span>  |  <span data-ttu-id="1aeed-139">はい</span><span class="sxs-lookup"><span data-stu-id="1aeed-139">Yes</span></span>  |  <span data-ttu-id="1aeed-140">関数の入力パラメーターを定義する配列です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="1aeed-141">詳細については、[parameters](#parameters) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1aeed-141">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="1aeed-142">object</span><span class="sxs-lookup"><span data-stu-id="1aeed-142">object</span></span>  |  <span data-ttu-id="1aeed-143">はい</span><span class="sxs-lookup"><span data-stu-id="1aeed-143">Yes</span></span>  |  <span data-ttu-id="1aeed-144">関数が返す情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="1aeed-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="1aeed-145">詳細については、[result](#result) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1aeed-145">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="1aeed-146">options</span><span class="sxs-lookup"><span data-stu-id="1aeed-146">options</span></span>

<span data-ttu-id="1aeed-147">`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-147">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="1aeed-148">次の表に、`options` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="1aeed-148">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="1aeed-149">プロパティ</span><span class="sxs-lookup"><span data-stu-id="1aeed-149">Property</span></span>  |  <span data-ttu-id="1aeed-150">データ型</span><span class="sxs-lookup"><span data-stu-id="1aeed-150">Data type</span></span>  |  <span data-ttu-id="1aeed-151">必須</span><span class="sxs-lookup"><span data-stu-id="1aeed-151">Required</span></span>  |  <span data-ttu-id="1aeed-152">説明</span><span class="sxs-lookup"><span data-stu-id="1aeed-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="1aeed-153">ブール</span><span class="sxs-lookup"><span data-stu-id="1aeed-153">boolean</span></span>  |  <span data-ttu-id="1aeed-154">いいえ</span><span class="sxs-lookup"><span data-stu-id="1aeed-154">No</span></span><br/><br/><span data-ttu-id="1aeed-155">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-155">Default value is `false`.</span></span>  |  <span data-ttu-id="1aeed-156">`true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `onCanceled` ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="1aeed-157">このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します </span><span class="sxs-lookup"><span data-stu-id="1aeed-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="1aeed-158">(このパラメーターを `parameters` プロパティには登録し***ない***でください)。</span><span class="sxs-lookup"><span data-stu-id="1aeed-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="1aeed-159">この関数の本文では、ハンドラーを `caller.onCanceled` メンバーに割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="1aeed-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="1aeed-160">詳細については、「[関数をキャンセルする](custom-functions-web-reqs.md#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1aeed-160">For more information, see [Canceling a function](custom-functions-web-reqs.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="1aeed-161">ブール</span><span class="sxs-lookup"><span data-stu-id="1aeed-161">boolean</span></span>  |  <span data-ttu-id="1aeed-162">いいえ</span><span class="sxs-lookup"><span data-stu-id="1aeed-162">No</span></span><br/><br/><span data-ttu-id="1aeed-163">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-163">Default value is `false`.</span></span>  |  <span data-ttu-id="1aeed-164">`true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="1aeed-165">このオプションは、株価などの急速に変化するデータ ソースに便利です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="1aeed-166">このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します </span><span class="sxs-lookup"><span data-stu-id="1aeed-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="1aeed-167">(このパラメーターを `parameters` プロパティには登録し***ない***でください)。</span><span class="sxs-lookup"><span data-stu-id="1aeed-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="1aeed-168">この関数には、`return` ステートメントは含めないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="1aeed-168">The function should have no `return` statement.</span></span> <span data-ttu-id="1aeed-169">代わりに、結果の値は `caller.setResult` コールバック メソッドの引数として渡されます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="1aeed-170">詳細については、「[ストリーミング関数](custom-functions-web-reqs.md#streaming-functions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1aeed-170">For more information, see [Streaming functions](custom-functions-web-reqs.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="1aeed-171">ブール</span><span class="sxs-lookup"><span data-stu-id="1aeed-171">boolean</span></span> | <span data-ttu-id="1aeed-172">いいえ</span><span class="sxs-lookup"><span data-stu-id="1aeed-172">No</span></span> <br/><br/><span data-ttu-id="1aeed-173">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-173">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="1aeed-174">`true` の場合は、数式の依存値が変更されたときのみではなく、Excel が再計算するたびに関数が再計算されます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-174">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="1aeed-175">関数は、ストリーミングと揮発性の両方にすることはできません。</span><span class="sxs-lookup"><span data-stu-id="1aeed-175">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="1aeed-176">`stream` と `volatile` の両方のプロパティが `true` に設定されている場合は、揮発性のオプションが無視されます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-176">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="1aeed-177">parameters</span><span class="sxs-lookup"><span data-stu-id="1aeed-177">parameters</span></span>

<span data-ttu-id="1aeed-178">`parameters` プロパティは、パラメーター オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-178">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="1aeed-179">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="1aeed-179">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="1aeed-180">プロパティ</span><span class="sxs-lookup"><span data-stu-id="1aeed-180">Property</span></span>  |  <span data-ttu-id="1aeed-181">データ型</span><span class="sxs-lookup"><span data-stu-id="1aeed-181">Data type</span></span>  |  <span data-ttu-id="1aeed-182">必須</span><span class="sxs-lookup"><span data-stu-id="1aeed-182">Required</span></span>  |  <span data-ttu-id="1aeed-183">説明</span><span class="sxs-lookup"><span data-stu-id="1aeed-183">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="1aeed-184">string</span><span class="sxs-lookup"><span data-stu-id="1aeed-184">string</span></span>  |  <span data-ttu-id="1aeed-185">いいえ</span><span class="sxs-lookup"><span data-stu-id="1aeed-185">No</span></span> |  <span data-ttu-id="1aeed-186">パラメーターの説明です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-186">A description of the parameter.</span></span> <span data-ttu-id="1aeed-187">これは、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-187">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="1aeed-188">string</span><span class="sxs-lookup"><span data-stu-id="1aeed-188">string</span></span>  |  <span data-ttu-id="1aeed-189">いいえ</span><span class="sxs-lookup"><span data-stu-id="1aeed-189">No</span></span>  |  <span data-ttu-id="1aeed-190">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="1aeed-190">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="1aeed-191">文字列</span><span class="sxs-lookup"><span data-stu-id="1aeed-191">string</span></span>  |  <span data-ttu-id="1aeed-192">はい</span><span class="sxs-lookup"><span data-stu-id="1aeed-192">Yes</span></span>  |  <span data-ttu-id="1aeed-193">パラメーターの名前です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-193">The name of the parameter.</span></span> <span data-ttu-id="1aeed-194">この名前は、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-194">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="1aeed-195">文字列</span><span class="sxs-lookup"><span data-stu-id="1aeed-195">string</span></span>  |  <span data-ttu-id="1aeed-196">いいえ</span><span class="sxs-lookup"><span data-stu-id="1aeed-196">No</span></span>  |  <span data-ttu-id="1aeed-197">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-197">The data type of the parameter.</span></span> <span data-ttu-id="1aeed-198">**boolean**、**number**、**string**、または **any** が可能です。ここでは、前の 3 種類のいずれかを使用できます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-198">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="1aeed-199">このプロパティが指定されていない場合、データ型の既定は **any** です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-199">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="1aeed-200">ブール</span><span class="sxs-lookup"><span data-stu-id="1aeed-200">boolean</span></span> | <span data-ttu-id="1aeed-201">いいえ</span><span class="sxs-lookup"><span data-stu-id="1aeed-201">No</span></span> | <span data-ttu-id="1aeed-202">`true` の場合、パラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-202">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="1aeed-203">省略可能なパラメーターの `type` プロパティが指定されていない場合や `any` に設定している場合は、Excel のセルに関数が入力されているときに、IDE の linting エラーや省略可能なパラメーターが表示されないなどの問題が発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="1aeed-203">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="1aeed-204">これについては、2018 年 12 月に変更される予定です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-204">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="1aeed-205">result</span><span class="sxs-lookup"><span data-stu-id="1aeed-205">result</span></span>

<span data-ttu-id="1aeed-206">`result` オブジェクトは、この関数が返す情報の種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="1aeed-206">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="1aeed-207">次の表に、`result` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="1aeed-207">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="1aeed-208">プロパティ</span><span class="sxs-lookup"><span data-stu-id="1aeed-208">Property</span></span>  |  <span data-ttu-id="1aeed-209">データ型</span><span class="sxs-lookup"><span data-stu-id="1aeed-209">Data type</span></span>  |  <span data-ttu-id="1aeed-210">必須</span><span class="sxs-lookup"><span data-stu-id="1aeed-210">Required</span></span>  |  <span data-ttu-id="1aeed-211">説明</span><span class="sxs-lookup"><span data-stu-id="1aeed-211">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="1aeed-212">string</span><span class="sxs-lookup"><span data-stu-id="1aeed-212">string</span></span>  |  <span data-ttu-id="1aeed-213">いいえ</span><span class="sxs-lookup"><span data-stu-id="1aeed-213">No</span></span>  |  <span data-ttu-id="1aeed-214">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="1aeed-214">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="1aeed-215">文字列</span><span class="sxs-lookup"><span data-stu-id="1aeed-215">string</span></span>  |  <span data-ttu-id="1aeed-216">はい</span><span class="sxs-lookup"><span data-stu-id="1aeed-216">Yes</span></span>  |  <span data-ttu-id="1aeed-217">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="1aeed-217">The data type of the parameter.</span></span> <span data-ttu-id="1aeed-218">**boolean**、**number**、**string**、または **any** である必要があります。ここでは、前の 3 種類のいずれかを使用できます。</span><span class="sxs-lookup"><span data-stu-id="1aeed-218">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="1aeed-219">関連項目</span><span class="sxs-lookup"><span data-stu-id="1aeed-219">See also</span></span>

* [<span data-ttu-id="1aeed-220">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="1aeed-220">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="1aeed-221">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="1aeed-221">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="1aeed-222">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="1aeed-222">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="1aeed-223">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="1aeed-223">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="1aeed-224">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="1aeed-224">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
