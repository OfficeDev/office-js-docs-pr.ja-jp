---
ms.date: 10/17/2018
description: Excel のカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ
ms.openlocfilehash: 0c77474188a2deefd23a73bb64e87569bb1fa52a
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298545"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="f45f0-103">カスタム関数のメタデータ (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="f45f0-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="f45f0-104">Excel アドイン内に[カスタム関数](custom-functions-overview.md)を定義する場合、カスタム関数を登録し、エンド ユーザーが利用できるようにするために Excel が必要とする情報を提供する JSON メタデータ ファイルをアドイン プロジェクトに含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="f45f0-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="f45f0-105">この記事では、その JSON メタデータ ファイルの形式について説明します。</span><span class="sxs-lookup"><span data-stu-id="f45f0-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="f45f0-106">カスタム関数を有効にするためにアドイン プロジェクトに含める必要のあるその他のファイルに関する情報については、「[Excel でのカスタム関数の作成](custom-functions-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f45f0-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="f45f0-107">メタデータの例</span><span class="sxs-lookup"><span data-stu-id="f45f0-107">Example metadata</span></span>

<span data-ttu-id="f45f0-108">次の例では、カスタム関数を定義するアドインの JSON メタデータ ファイルの内容を示します。</span><span class="sxs-lookup"><span data-stu-id="f45f0-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="f45f0-109">この例の後に続くセクションでは、JSON の例に含まれる個々のプロパティの詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="f45f0-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="f45f0-110">完全な JSON ファイルのサンプルは、[OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub リポジトリにあります。</span><span class="sxs-lookup"><span data-stu-id="f45f0-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="f45f0-111">functions</span><span class="sxs-lookup"><span data-stu-id="f45f0-111">functions</span></span> 

<span data-ttu-id="f45f0-112">`functions` プロパティは、カスタム関数オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-112">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="f45f0-113">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="f45f0-113">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="f45f0-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f45f0-114">Property</span></span>  |  <span data-ttu-id="f45f0-115">データ型</span><span class="sxs-lookup"><span data-stu-id="f45f0-115">Data type</span></span>  |  <span data-ttu-id="f45f0-116">必須</span><span class="sxs-lookup"><span data-stu-id="f45f0-116">Required</span></span>  |  <span data-ttu-id="f45f0-117">説明</span><span class="sxs-lookup"><span data-stu-id="f45f0-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="f45f0-118">文字列</span><span class="sxs-lookup"><span data-stu-id="f45f0-118">string</span></span>  |  <span data-ttu-id="f45f0-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="f45f0-119">No</span></span>  |  <span data-ttu-id="f45f0-120">Excel でエンド ユーザーに表示される関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="f45f0-121">たとえば、「**華氏の値を摂氏に変換する**」です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="f45f0-122">文字列</span><span class="sxs-lookup"><span data-stu-id="f45f0-122">string</span></span>  |   <span data-ttu-id="f45f0-123">いいえ</span><span class="sxs-lookup"><span data-stu-id="f45f0-123">No</span></span>  |  <span data-ttu-id="f45f0-124">関数に関する情報を提供する URL です </span><span class="sxs-lookup"><span data-stu-id="f45f0-124">URL that provides information about the function.</span></span> <span data-ttu-id="f45f0-125">(作業ウィンドウに表示されます)。たとえば、**http://contoso.com/help/convertcelsiustofahrenheit.html** です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="f45f0-126">文字列</span><span class="sxs-lookup"><span data-stu-id="f45f0-126">string</span></span> | <span data-ttu-id="f45f0-127">はい</span><span class="sxs-lookup"><span data-stu-id="f45f0-127">Yes</span></span> | <span data-ttu-id="f45f0-128">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-128">A unique ID for the group.</span></span> <span data-ttu-id="f45f0-129">この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。</span><span class="sxs-lookup"><span data-stu-id="f45f0-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="f45f0-130">文字列</span><span class="sxs-lookup"><span data-stu-id="f45f0-130">string</span></span>  |  <span data-ttu-id="f45f0-131">はい</span><span class="sxs-lookup"><span data-stu-id="f45f0-131">Yes</span></span>  |  <span data-ttu-id="f45f0-132">Excel でエンド ユーザーに表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="f45f0-133">Excel では、この関数名は XML マニフェスト ファイルで指定されているカスタム関数の名前空間でプレフィックスされます。</span><span class="sxs-lookup"><span data-stu-id="f45f0-133">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="f45f0-134">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f45f0-134">object</span></span>  |  <span data-ttu-id="f45f0-135">いいえ</span><span class="sxs-lookup"><span data-stu-id="f45f0-135">No</span></span>  |  <span data-ttu-id="f45f0-136">Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="f45f0-136">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="f45f0-137">詳細については、[options オブジェクト](#options-object)に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f45f0-137">See object load [options](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="f45f0-138">配列</span><span class="sxs-lookup"><span data-stu-id="f45f0-138">array</span></span>  |  <span data-ttu-id="f45f0-139">はい</span><span class="sxs-lookup"><span data-stu-id="f45f0-139">Yes</span></span>  |  <span data-ttu-id="f45f0-140">関数の入力パラメーターを定義する配列です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="f45f0-141">詳細については、[parameters 配列](#parameters-array)に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f45f0-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="f45f0-142">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f45f0-142">object</span></span>  |  <span data-ttu-id="f45f0-143">はい</span><span class="sxs-lookup"><span data-stu-id="f45f0-143">Yes</span></span>  |  <span data-ttu-id="f45f0-144">関数が返す情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f45f0-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="f45f0-145">詳細については、[result オブジェクト](#result-object)に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f45f0-145">See object load [options](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="f45f0-146">options</span><span class="sxs-lookup"><span data-stu-id="f45f0-146">options</span></span>

<span data-ttu-id="f45f0-147">`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="f45f0-147">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="f45f0-148">次の表に、`options` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="f45f0-148">The following table lists the parts of the `options` claim.</span></span>

|  <span data-ttu-id="f45f0-149">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f45f0-149">Property</span></span>  |  <span data-ttu-id="f45f0-150">データ型</span><span class="sxs-lookup"><span data-stu-id="f45f0-150">Data type</span></span>  |  <span data-ttu-id="f45f0-151">必須</span><span class="sxs-lookup"><span data-stu-id="f45f0-151">Required</span></span>  |  <span data-ttu-id="f45f0-152">説明</span><span class="sxs-lookup"><span data-stu-id="f45f0-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="f45f0-153">ブール</span><span class="sxs-lookup"><span data-stu-id="f45f0-153">boolean</span></span>  |  <span data-ttu-id="f45f0-154">いいえ</span><span class="sxs-lookup"><span data-stu-id="f45f0-154">No</span></span><br/><br/><span data-ttu-id="f45f0-155">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-155">Default value is  `false`.</span></span>  |  <span data-ttu-id="f45f0-156">`true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `onCanceled` ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f45f0-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="f45f0-157">このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します </span><span class="sxs-lookup"><span data-stu-id="f45f0-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="f45f0-158">(このパラメーターを `parameters` プロパティには登録し***ない***でください)。</span><span class="sxs-lookup"><span data-stu-id="f45f0-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="f45f0-159">この関数の本文では、ハンドラーを `caller.onCanceled` メンバーに割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="f45f0-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="f45f0-160">詳細については、「[関数をキャンセルする](custom-functions-overview.md#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f45f0-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="f45f0-161">ブール</span><span class="sxs-lookup"><span data-stu-id="f45f0-161">boolean</span></span>  |  <span data-ttu-id="f45f0-162">いいえ</span><span class="sxs-lookup"><span data-stu-id="f45f0-162">No</span></span><br/><br/><span data-ttu-id="f45f0-163">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-163">Default value is  `false`.</span></span>  |  <span data-ttu-id="f45f0-164">`true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。</span><span class="sxs-lookup"><span data-stu-id="f45f0-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="f45f0-165">このオプションは、株価などの急速に変化するデータ ソースに便利です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="f45f0-166">このオプションを使用する場合、Excel は追加の `caller` パラメーターを使用して JavaScript 関数を呼び出します </span><span class="sxs-lookup"><span data-stu-id="f45f0-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="f45f0-167">(このパラメーターを `parameters` プロパティには登録し***ない***でください)。</span><span class="sxs-lookup"><span data-stu-id="f45f0-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="f45f0-168">この関数には、`return` ステートメントは含めないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f45f0-168">The function should have no `return` statement.</span></span> <span data-ttu-id="f45f0-169">代わりに、結果の値は `caller.setResult` コールバック メソッドの引数として渡されます。</span><span class="sxs-lookup"><span data-stu-id="f45f0-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="f45f0-170">詳細については、「[ストリーミング関数](custom-functions-overview.md#streaming-functions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f45f0-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="f45f0-171">parameters</span><span class="sxs-lookup"><span data-stu-id="f45f0-171">parameters</span></span>

<span data-ttu-id="f45f0-172">`parameters` プロパティは、パラメーター オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-172">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="f45f0-173">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="f45f0-173">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="f45f0-174">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f45f0-174">Property</span></span>  |  <span data-ttu-id="f45f0-175">データ型</span><span class="sxs-lookup"><span data-stu-id="f45f0-175">Data type</span></span>  |  <span data-ttu-id="f45f0-176">必須</span><span class="sxs-lookup"><span data-stu-id="f45f0-176">Required</span></span>  |  <span data-ttu-id="f45f0-177">説明</span><span class="sxs-lookup"><span data-stu-id="f45f0-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="f45f0-178">文字列</span><span class="sxs-lookup"><span data-stu-id="f45f0-178">string</span></span>  |  <span data-ttu-id="f45f0-179">いいえ</span><span class="sxs-lookup"><span data-stu-id="f45f0-179">No</span></span> |  <span data-ttu-id="f45f0-180">パラメーターの説明です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-180">A description of the parameter.</span></span> <span data-ttu-id="f45f0-181">これは、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="f45f0-181">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="f45f0-182">文字列</span><span class="sxs-lookup"><span data-stu-id="f45f0-182">string</span></span>  |  <span data-ttu-id="f45f0-183">いいえ</span><span class="sxs-lookup"><span data-stu-id="f45f0-183">No</span></span>  |  <span data-ttu-id="f45f0-184">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="f45f0-184">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="f45f0-185">文字列</span><span class="sxs-lookup"><span data-stu-id="f45f0-185">string</span></span>  |  <span data-ttu-id="f45f0-186">はい</span><span class="sxs-lookup"><span data-stu-id="f45f0-186">Yes</span></span>  |  <span data-ttu-id="f45f0-187">パラメーターの名前です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-187">The name of the parameter.</span></span> <span data-ttu-id="f45f0-188">この名前は、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="f45f0-188">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="f45f0-189">文字列</span><span class="sxs-lookup"><span data-stu-id="f45f0-189">string</span></span>  |  <span data-ttu-id="f45f0-190">いいえ</span><span class="sxs-lookup"><span data-stu-id="f45f0-190">No</span></span>  |  <span data-ttu-id="f45f0-191">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-191">The System data type of the parameter.</span></span> <span data-ttu-id="f45f0-192">**boolean**、**number**、**string**、または **any** が可能です。ここでは、前の 3 種類のいずれかを使用できます。</span><span class="sxs-lookup"><span data-stu-id="f45f0-192">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="f45f0-193">このプロパティが指定されていない場合、データ型の既定は **any** です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-193">If this property is not specified, the data type defaults to **any**.</span></span> |

## <a name="result"></a><span data-ttu-id="f45f0-194">result</span><span class="sxs-lookup"><span data-stu-id="f45f0-194">result</span></span>

<span data-ttu-id="f45f0-195">`result` オブジェクトは、この関数が返す情報の種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="f45f0-195">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="f45f0-196">次の表に、`result` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="f45f0-196">The following table lists the parts of the `result` claim.</span></span>

|  <span data-ttu-id="f45f0-197">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f45f0-197">Property</span></span>  |  <span data-ttu-id="f45f0-198">データ型</span><span class="sxs-lookup"><span data-stu-id="f45f0-198">Data type</span></span>  |  <span data-ttu-id="f45f0-199">必須</span><span class="sxs-lookup"><span data-stu-id="f45f0-199">Required</span></span>  |  <span data-ttu-id="f45f0-200">説明</span><span class="sxs-lookup"><span data-stu-id="f45f0-200">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="f45f0-201">文字列</span><span class="sxs-lookup"><span data-stu-id="f45f0-201">string</span></span>  |  <span data-ttu-id="f45f0-202">いいえ</span><span class="sxs-lookup"><span data-stu-id="f45f0-202">No</span></span>  |  <span data-ttu-id="f45f0-203">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="f45f0-203">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="f45f0-204">文字列</span><span class="sxs-lookup"><span data-stu-id="f45f0-204">string</span></span>  |  <span data-ttu-id="f45f0-205">はい</span><span class="sxs-lookup"><span data-stu-id="f45f0-205">Yes</span></span>  |  <span data-ttu-id="f45f0-206">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="f45f0-206">The System data type of the parameter.</span></span> <span data-ttu-id="f45f0-207">**boolean**、**number**、**string**、または **any** である必要があります。ここでは、前の 3 種類のいずれかを使用できます。</span><span class="sxs-lookup"><span data-stu-id="f45f0-207">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="f45f0-208">関連項目</span><span class="sxs-lookup"><span data-stu-id="f45f0-208">See also</span></span>

* [<span data-ttu-id="f45f0-209">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="f45f0-209">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="f45f0-210">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="f45f0-210">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="f45f0-211">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="f45f0-211">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="f45f0-212">チュートリアル: Excel でカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="f45f0-212">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
