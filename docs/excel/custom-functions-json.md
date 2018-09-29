---
ms.date: 09/27/2018
description: Excel でカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ
ms.openlocfilehash: a179a9c4bc071200cab1377c5e48913bfc8358cf
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348795"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="99d01-103">カスタム関数のメタデータ (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="99d01-103">Custom functions metadata</span></span>

<span data-ttu-id="99d01-p101">Excel アドインで [カスタム関数](custom-functions-overview.md) を定義するときに、アドイン プロジェクトは、Excel がカスタム関数を登録し、エンド ユーザーが利用できるようにする必要がある情報を提供する JSON メタデータ ファイルを含める必要があります。この記事では、JSON メタデータ ファイルの形式について説明します。</span><span class="sxs-lookup"><span data-stu-id="99d01-p101">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end-users. This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="99d01-106">カスタム関数を有効にするためにアドイン プロジェクトに含める必要のある、その他のファイルに関する情報については、「[Excel でカスタム関数を作成する](custom-functions-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="99d01-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="99d01-107">メタデータの例</span><span class="sxs-lookup"><span data-stu-id="99d01-107">Example metadata</span></span>

<span data-ttu-id="99d01-108">次の例は、カスタム関数を定義するアドイン用の JSON メタデータ ファイルの内容を示しています。</span><span class="sxs-lookup"><span data-stu-id="99d01-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="99d01-109">この例に続くセクションでは、この JSON の例の中にある個々のプロパティについての詳細情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="99d01-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="99d01-110">JSON ファイルの完全なサンプルは、「[OfficeDev/Excel-Custom-Functions GitHub リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json)」で利用可能です。</span><span class="sxs-lookup"><span data-stu-id="99d01-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="99d01-111">functions</span><span class="sxs-lookup"><span data-stu-id="99d01-111">functions</span></span> 

<span data-ttu-id="99d01-112">`functions` プロパティは、カスタム関数オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="99d01-112">The `functions` property is an array of objects.</span></span> <span data-ttu-id="99d01-113">次の表で、各オブジェクトのプロパティを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="99d01-113">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="99d01-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="99d01-114">Property</span></span>  |  <span data-ttu-id="99d01-115">データ型</span><span class="sxs-lookup"><span data-stu-id="99d01-115">Data type</span></span>  |  <span data-ttu-id="99d01-116">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="99d01-116">Required</span></span>  |  <span data-ttu-id="99d01-117">説明</span><span class="sxs-lookup"><span data-stu-id="99d01-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="99d01-118">文字列</span><span class="sxs-lookup"><span data-stu-id="99d01-118">string</span></span>  |  <span data-ttu-id="99d01-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="99d01-119">No</span></span>  |  <span data-ttu-id="99d01-p104">エンド ユーザーに Excel で表示される関数の説明です。たとえば、 **華氏温度値を摂氏に変換**します。</span><span class="sxs-lookup"><span data-stu-id="99d01-p104">A description of the function that appears in the Excel UI. For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="99d01-122">文字列</span><span class="sxs-lookup"><span data-stu-id="99d01-122">string</span></span>  |   <span data-ttu-id="99d01-123">いいえ</span><span class="sxs-lookup"><span data-stu-id="99d01-123">No</span></span>  |  <span data-ttu-id="99d01-p105">関数に関する情報を提供する URL です。(これは、作業ウィンドウに表示されます。) たとえば、**http://contoso.com/help/convertcelsiustofahrenheit.html**です。</span><span class="sxs-lookup"><span data-stu-id="99d01-p105">URL where users can get information about the function. (It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="99d01-126">文字列</span><span class="sxs-lookup"><span data-stu-id="99d01-126">string</span></span> | <span data-ttu-id="99d01-127">はい</span><span class="sxs-lookup"><span data-stu-id="99d01-127">Yes</span></span> | <span data-ttu-id="99d01-128">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="99d01-128">A unique ID for the group.</span></span> <span data-ttu-id="99d01-129">設定後は、この ID は変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="99d01-129">This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="99d01-130">文字列</span><span class="sxs-lookup"><span data-stu-id="99d01-130">string</span></span>  |  <span data-ttu-id="99d01-131">はい</span><span class="sxs-lookup"><span data-stu-id="99d01-131">Yes</span></span>  |  <span data-ttu-id="99d01-132">エンド ユーザーに Excel で表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="99d01-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="99d01-133">Excel では、この関数名は、XML マニフェスト ファイルで指定されているカスタム関数の名前空間が接頭辞となります。</span><span class="sxs-lookup"><span data-stu-id="99d01-133">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="99d01-134">object</span><span class="sxs-lookup"><span data-stu-id="99d01-134">object</span></span>  |  <span data-ttu-id="99d01-135">いいえ</span><span class="sxs-lookup"><span data-stu-id="99d01-135">No</span></span>  |  <span data-ttu-id="99d01-136">Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。</span><span class="sxs-lookup"><span data-stu-id="99d01-136">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="99d01-137">詳細については、「[オプション オブジェクト](#options-object)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="99d01-137">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="99d01-138">array</span><span class="sxs-lookup"><span data-stu-id="99d01-138">array</span></span>  |  <span data-ttu-id="99d01-139">はい</span><span class="sxs-lookup"><span data-stu-id="99d01-139">Yes</span></span>  |  <span data-ttu-id="99d01-140">関数の入力パラメーターを定義する配列です。</span><span class="sxs-lookup"><span data-stu-id="99d01-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="99d01-141">詳細については、「[パラメーター配列](#parameters-array)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="99d01-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="99d01-142">object</span><span class="sxs-lookup"><span data-stu-id="99d01-142">object</span></span>  |  <span data-ttu-id="99d01-143">はい</span><span class="sxs-lookup"><span data-stu-id="99d01-143">Yes</span></span>  |  <span data-ttu-id="99d01-144">関数によって返される情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="99d01-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="99d01-145">詳細については、「[結果オブジェクト](#result-object)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="99d01-145">See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="99d01-146">options</span><span class="sxs-lookup"><span data-stu-id="99d01-146">options</span></span>

<span data-ttu-id="99d01-147">`options` オブジェクトを使用すると、Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。</span><span class="sxs-lookup"><span data-stu-id="99d01-147">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="99d01-148">次の表で、`options` オブジェクトのプロパティを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="99d01-148">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="99d01-149">プロパティ</span><span class="sxs-lookup"><span data-stu-id="99d01-149">Property</span></span>  |  <span data-ttu-id="99d01-150">データ型</span><span class="sxs-lookup"><span data-stu-id="99d01-150">Data type</span></span>  |  <span data-ttu-id="99d01-151">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="99d01-151">Required</span></span>  |  <span data-ttu-id="99d01-152">説明</span><span class="sxs-lookup"><span data-stu-id="99d01-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="99d01-153">ブール値</span><span class="sxs-lookup"><span data-stu-id="99d01-153">boolean</span></span>  |  <span data-ttu-id="99d01-154">いいえ</span><span class="sxs-lookup"><span data-stu-id="99d01-154">No</span></span><br/><br/><span data-ttu-id="99d01-155">既定値は`false` です。</span><span class="sxs-lookup"><span data-stu-id="99d01-155">Default value is 4.</span></span>  |  <span data-ttu-id="99d01-156">`true` の場合、Excel はユーザーが関数をキャンセルする操作をするたびに `onCanceled` ハンドラを呼び出します。たとえば、手動で再計算をトリガーするか、関数が参照するセルを編集する場合です。</span><span class="sxs-lookup"><span data-stu-id="99d01-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="99d01-157">このオプションを使用すると、Excelは `caller` パラメーターを追加して JavaScript 関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="99d01-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="99d01-158">(このパラメータを `parameters` プロパティに登録***しない***でください)。</span><span class="sxs-lookup"><span data-stu-id="99d01-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="99d01-159">関数の本体では、ハンドラは `caller.onCanceled` メンバーに割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="99d01-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="99d01-160">詳細については、 「[関数をキャンセルする](custom-functions-overview.md#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="99d01-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="99d01-161">ブール値</span><span class="sxs-lookup"><span data-stu-id="99d01-161">boolean</span></span>  |  <span data-ttu-id="99d01-162">いいえ</span><span class="sxs-lookup"><span data-stu-id="99d01-162">No</span></span><br/><br/><span data-ttu-id="99d01-163">既定値は`false` です。</span><span class="sxs-lookup"><span data-stu-id="99d01-163">Default value is 4.</span></span>  |  <span data-ttu-id="99d01-164">`true` の場合、関数は一度だけの呼び出しでも繰り返しセルに出力できます。</span><span class="sxs-lookup"><span data-stu-id="99d01-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="99d01-165">このオプションは、株価など急激に変化するデータソースで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="99d01-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="99d01-166">このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="99d01-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="99d01-167">(このパラメーターを `parameters` プロパティに登録***しない***でください)。</span><span class="sxs-lookup"><span data-stu-id="99d01-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="99d01-168">関数では、`return` 文を使いません。</span><span class="sxs-lookup"><span data-stu-id="99d01-168">The function should have no `return` statement.</span></span> <span data-ttu-id="99d01-169">代わりに、結果値を `caller.setResult` コールバック メソッドの引数として渡します。</span><span class="sxs-lookup"><span data-stu-id="99d01-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="99d01-170">詳細については、「[ストリーム関数](custom-functions-overview.md#streamed-functions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="99d01-170">For more information, see [Excel functions by category](custom-functions-overview.md#streamed-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="99d01-171">parameters</span><span class="sxs-lookup"><span data-stu-id="99d01-171">parameters</span></span>

<span data-ttu-id="99d01-172">`parameters` プロパティは、パラメーター オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="99d01-172">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="99d01-173">次の表で、各オブジェクトのプロパティを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="99d01-173">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="99d01-174">プロパティ</span><span class="sxs-lookup"><span data-stu-id="99d01-174">Property</span></span>  |  <span data-ttu-id="99d01-175">データ型</span><span class="sxs-lookup"><span data-stu-id="99d01-175">Data type</span></span>  |  <span data-ttu-id="99d01-176">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="99d01-176">Required</span></span>  |  <span data-ttu-id="99d01-177">説明</span><span class="sxs-lookup"><span data-stu-id="99d01-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="99d01-178">文字列</span><span class="sxs-lookup"><span data-stu-id="99d01-178">string</span></span>  |  <span data-ttu-id="99d01-179">いいえ</span><span class="sxs-lookup"><span data-stu-id="99d01-179">No</span></span> |  <span data-ttu-id="99d01-180">パラメータの説明。</span><span class="sxs-lookup"><span data-stu-id="99d01-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="99d01-181">文字列</span><span class="sxs-lookup"><span data-stu-id="99d01-181">string</span></span>  |  <span data-ttu-id="99d01-182">いいえ</span><span class="sxs-lookup"><span data-stu-id="99d01-182">No</span></span>  |  <span data-ttu-id="99d01-183">**scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="99d01-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="99d01-184">文字列</span><span class="sxs-lookup"><span data-stu-id="99d01-184">string</span></span>  |  <span data-ttu-id="99d01-185">はい</span><span class="sxs-lookup"><span data-stu-id="99d01-185">Yes</span></span>  |  <span data-ttu-id="99d01-186">パラメーターの名前です。</span><span class="sxs-lookup"><span data-stu-id="99d01-186">The name of the parameter.</span></span> <span data-ttu-id="99d01-187">この名前は Excel の IntelliSense で表示されます。</span><span class="sxs-lookup"><span data-stu-id="99d01-187">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="99d01-188">文字列</span><span class="sxs-lookup"><span data-stu-id="99d01-188">string</span></span>  |  <span data-ttu-id="99d01-189">いいえ</span><span class="sxs-lookup"><span data-stu-id="99d01-189">No</span></span>  |  <span data-ttu-id="99d01-190">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="99d01-190">The data type of the parameter.</span></span> <span data-ttu-id="99d01-191">**ブール値**、 **数値**、または **文字列**である必要があります。</span><span class="sxs-lookup"><span data-stu-id="99d01-191">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result"></a><span data-ttu-id="99d01-192">result</span><span class="sxs-lookup"><span data-stu-id="99d01-192">result</span></span>

<span data-ttu-id="99d01-193">関数によって返される情報の種類を定義する`results` オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="99d01-193">The `results` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="99d01-194">次の表で、`result` オブジェクトのプロパティを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="99d01-194">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="99d01-195">プロパティ</span><span class="sxs-lookup"><span data-stu-id="99d01-195">Property</span></span>  |  <span data-ttu-id="99d01-196">データ型</span><span class="sxs-lookup"><span data-stu-id="99d01-196">Data type</span></span>  |  <span data-ttu-id="99d01-197">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="99d01-197">Required</span></span>  |  <span data-ttu-id="99d01-198">説明</span><span class="sxs-lookup"><span data-stu-id="99d01-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="99d01-199">文字列</span><span class="sxs-lookup"><span data-stu-id="99d01-199">string</span></span>  |  <span data-ttu-id="99d01-200">いいえ</span><span class="sxs-lookup"><span data-stu-id="99d01-200">No</span></span>  |  <span data-ttu-id="99d01-201">**scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="99d01-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="99d01-202">文字列</span><span class="sxs-lookup"><span data-stu-id="99d01-202">string</span></span>  |  <span data-ttu-id="99d01-203">はい</span><span class="sxs-lookup"><span data-stu-id="99d01-203">Yes</span></span>  |  <span data-ttu-id="99d01-204">パラメーターのデータ型。</span><span class="sxs-lookup"><span data-stu-id="99d01-204">The data type of the parameter.</span></span> <span data-ttu-id="99d01-205">**ブール値**、 **数値**、または **文字列**である必要があります。</span><span class="sxs-lookup"><span data-stu-id="99d01-205">Must be "boolean", "number", or "string".</span></span>  |

## <a name="see-also"></a><span data-ttu-id="99d01-206">関連項目</span><span class="sxs-lookup"><span data-stu-id="99d01-206">See also</span></span>

* [<span data-ttu-id="99d01-207">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="99d01-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="99d01-208">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="99d01-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="99d01-209">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="99d01-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="99d01-210">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="99d01-210">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)