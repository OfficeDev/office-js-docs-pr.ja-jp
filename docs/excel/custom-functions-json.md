---
ms.date: 10/17/2018
description: Excel でカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ
ms.openlocfilehash: cff1cbc22f39c99597d4abe7005d7b8bbce6e185
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640009"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="9cd88-103">カスタム関数のメタデータ (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="9cd88-103">Custom functions metadata</span></span>

<span data-ttu-id="9cd88-p101">Excel アドインで [カスタム関数](custom-functions-overview.md) を定義する場合、アドイン プロジェクトには、Excel がカスタム関数を登録してエンド ユーザーが利用できるようにするために必要な情報を提供する JSON メタデータ ファイルを含める必要があります。この記事では、JSON メタデータ ファイルの形式について説明します。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p101">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="9cd88-106">カスタム関数を有効にするためにアドイン プロジェクトに含める必要のあるその他のファイルに関する情報については、「[Excel でカスタム関数を作成する](custom-functions-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9cd88-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="9cd88-107">メタデータの例</span><span class="sxs-lookup"><span data-stu-id="9cd88-107">Example metadata</span></span>

<span data-ttu-id="9cd88-p102">次の使用例は、JSON のアドインでカスタム関数を定義するメタデータ ファイルの内容を示しています。次の使用例を次のセクションでは、この例を JSON 内の個別のプロパティに関する詳細情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p102">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions. The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="9cd88-110">JSON ファイルの完全なサンプルは、「[OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) 」GitHub リポジトリで利用可能です。</span><span class="sxs-lookup"><span data-stu-id="9cd88-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="9cd88-111">functions</span><span class="sxs-lookup"><span data-stu-id="9cd88-111">functions</span></span> 

<span data-ttu-id="9cd88-p103">`functions` プロパティは、カスタム関数オブジェクトの配列です。次の表は、各オブジェクトのプロパティを一覧表示しています。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p103">The `functions` property is an array of custom function objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="9cd88-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="9cd88-114">Property</span></span>  |  <span data-ttu-id="9cd88-115">データ型</span><span class="sxs-lookup"><span data-stu-id="9cd88-115">Data type</span></span>  |  <span data-ttu-id="9cd88-116">必須</span><span class="sxs-lookup"><span data-stu-id="9cd88-116">Required</span></span>  |  <span data-ttu-id="9cd88-117">説明</span><span class="sxs-lookup"><span data-stu-id="9cd88-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="9cd88-118">文字列</span><span class="sxs-lookup"><span data-stu-id="9cd88-118">string</span></span>  |  <span data-ttu-id="9cd88-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="9cd88-119">No</span></span>  |  <span data-ttu-id="9cd88-p104">Excel でエンド ユーザーに表示される関数の説明です。例: "**華氏温度値を摂氏に変換する**" 。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p104">The description of the function that end users see in Excel. For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="9cd88-122">文字列</span><span class="sxs-lookup"><span data-stu-id="9cd88-122">string</span></span>  |   <span data-ttu-id="9cd88-123">いいえ</span><span class="sxs-lookup"><span data-stu-id="9cd88-123">No</span></span>  |  <span data-ttu-id="9cd88-p105">関数に関する情報を提供する URL です。(作業ウィンドウに表示されます。) 例: **http://contoso.com/help/convertcelsiustofahrenheit.html**。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p105">URL that provides information about the function. (It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="9cd88-126">文字列</span><span class="sxs-lookup"><span data-stu-id="9cd88-126">string</span></span> | <span data-ttu-id="9cd88-127">はい</span><span class="sxs-lookup"><span data-stu-id="9cd88-127">Yes</span></span> | <span data-ttu-id="9cd88-128">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="9cd88-128">A unique ID for the function.</span></span> <span data-ttu-id="9cd88-129">この ID は、英数字とピリオドのみを含めることができ、設定された後、変更してはいけません。</span><span class="sxs-lookup"><span data-stu-id="9cd88-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="9cd88-130">文字列</span><span class="sxs-lookup"><span data-stu-id="9cd88-130">string</span></span>  |  <span data-ttu-id="9cd88-131">はい</span><span class="sxs-lookup"><span data-stu-id="9cd88-131">Yes</span></span>  |  <span data-ttu-id="9cd88-p107">Excel でエンド ユーザーに表示される関数の名前です。Excel では、この関数名が XML マニフェスト ファイルで指定されているカスタム関数の名前空間で接頭辞となります。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p107">The name of the function that end users see in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="9cd88-134">object</span><span class="sxs-lookup"><span data-stu-id="9cd88-134">object</span></span>  |  <span data-ttu-id="9cd88-135">いいえ</span><span class="sxs-lookup"><span data-stu-id="9cd88-135">No</span></span>  |  <span data-ttu-id="9cd88-p108">Excel で関数を実行する方法とタイミングのいくつかの側面をカスタマイズできます。詳細については、 [オプションのオブジェクト](#options-object) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p108">Enables you to customize some aspects of how and when Excel executes the function. See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="9cd88-138">配列</span><span class="sxs-lookup"><span data-stu-id="9cd88-138">array</span></span>  |  <span data-ttu-id="9cd88-139">はい</span><span class="sxs-lookup"><span data-stu-id="9cd88-139">Yes</span></span>  |  <span data-ttu-id="9cd88-p109">関数の入力パラメーターを定義する配列。詳細については、 [パラメーター配列](#parameters-array) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p109">Array that defines the input parameters for the function. See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="9cd88-142">object</span><span class="sxs-lookup"><span data-stu-id="9cd88-142">object</span></span>  |  <span data-ttu-id="9cd88-143">はい</span><span class="sxs-lookup"><span data-stu-id="9cd88-143">Yes</span></span>  |  <span data-ttu-id="9cd88-p110">関数によって返される情報の種類を定義するオブジェクト。詳細については、 [結果のオブジェクト](#result-object) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p110">Object that defines the type of information that is returned by the function. See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="9cd88-146">options</span><span class="sxs-lookup"><span data-stu-id="9cd88-146">options</span></span>

<span data-ttu-id="9cd88-p111">`options` オブジェクトは、Excel で関数を実行する方法とタイミングのいくつかの側面をカスタマイズすることができます。次の表に`options` オブジェクトのプロパティを一覧表示しています。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p111">The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="9cd88-149">プロパティ</span><span class="sxs-lookup"><span data-stu-id="9cd88-149">Property</span></span>  |  <span data-ttu-id="9cd88-150">データ型</span><span class="sxs-lookup"><span data-stu-id="9cd88-150">Data type</span></span>  |  <span data-ttu-id="9cd88-151">必須</span><span class="sxs-lookup"><span data-stu-id="9cd88-151">Required</span></span>  |  <span data-ttu-id="9cd88-152">説明</span><span class="sxs-lookup"><span data-stu-id="9cd88-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="9cd88-153">boolean</span><span class="sxs-lookup"><span data-stu-id="9cd88-153">boolean</span></span>  |  <span data-ttu-id="9cd88-154">いいえ</span><span class="sxs-lookup"><span data-stu-id="9cd88-154">No</span></span><br/><br/><span data-ttu-id="9cd88-155">既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="9cd88-155">Default value is 4.</span></span>  |  <span data-ttu-id="9cd88-p112">`true` を使用する場合、関数をキャンセルすることになる操作をユーザーが実行するたびに Excel は、 `onCanceled` ハンドラーを呼び出します。例えば、手動で再計算をトリガーしたり、関数が参照しているセルを編集したりなどの操作です。このオプションを使用する場合、Excel は、`caller` パラメータを追加して、JavaScript 関数を呼び出します 。(`parameters` プロパティにこのパラメータを登録***しない***でください )。関数の本文では、`caller.onCanceled` のメンバーにハンドラーを割り当てる必要があります。詳細については、 [関数をキャンセルする](custom-functions-overview.md#canceling-a-function)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p112">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="9cd88-161">ブール値</span><span class="sxs-lookup"><span data-stu-id="9cd88-161">boolean</span></span>  |  <span data-ttu-id="9cd88-162">いいえ</span><span class="sxs-lookup"><span data-stu-id="9cd88-162">No</span></span><br/><br/><span data-ttu-id="9cd88-163">既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="9cd88-163">Default value is 4.</span></span>  |  <span data-ttu-id="9cd88-164">`true` の場合、関数を一度呼び出すだけでセルに繰り返し出力できます。</span><span class="sxs-lookup"><span data-stu-id="9cd88-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="9cd88-165">このオプションは、株価など急激に変化するデータソースで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="9cd88-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="9cd88-166">このオプションを使用すると、Excelは `caller` パラメータを追加して JavaScript 関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="9cd88-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="9cd88-167">(このパラメーターを `parameters` プロパティに登録***しない***でください)。</span><span class="sxs-lookup"><span data-stu-id="9cd88-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="9cd88-168">関数では、`return` 文を使わないでください。</span><span class="sxs-lookup"><span data-stu-id="9cd88-168">The function should have no `return` statement.</span></span> <span data-ttu-id="9cd88-169">代わりに、戻り値を `caller.setResult` コールバック メソッドの引数として渡します。</span><span class="sxs-lookup"><span data-stu-id="9cd88-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="9cd88-170">詳細については、「 [ストリーミング関数](custom-functions-overview.md#streaming-functions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9cd88-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="9cd88-171">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9cd88-171">parameters</span></span>

<span data-ttu-id="9cd88-p114">`parameters` プロパティは、カスタム関数オブジェクトの配列です。次の表は、各オブジェクトのプロパティを一覧表示しています。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p114">The `parameters` property is an array of parameter objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="9cd88-174">プロパティ</span><span class="sxs-lookup"><span data-stu-id="9cd88-174">Property</span></span>  |  <span data-ttu-id="9cd88-175">データ型</span><span class="sxs-lookup"><span data-stu-id="9cd88-175">Data type</span></span>  |  <span data-ttu-id="9cd88-176">必須</span><span class="sxs-lookup"><span data-stu-id="9cd88-176">Required</span></span>  |  <span data-ttu-id="9cd88-177">説明</span><span class="sxs-lookup"><span data-stu-id="9cd88-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="9cd88-178">文字列</span><span class="sxs-lookup"><span data-stu-id="9cd88-178">string</span></span>  |  <span data-ttu-id="9cd88-179">いいえ</span><span class="sxs-lookup"><span data-stu-id="9cd88-179">No</span></span> |  <span data-ttu-id="9cd88-180">パラメータの説明。</span><span class="sxs-lookup"><span data-stu-id="9cd88-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="9cd88-181">文字列</span><span class="sxs-lookup"><span data-stu-id="9cd88-181">string</span></span>  |  <span data-ttu-id="9cd88-182">いいえ</span><span class="sxs-lookup"><span data-stu-id="9cd88-182">No</span></span>  |  <span data-ttu-id="9cd88-183">**scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="9cd88-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="9cd88-184">文字列</span><span class="sxs-lookup"><span data-stu-id="9cd88-184">string</span></span>  |  <span data-ttu-id="9cd88-185">はい</span><span class="sxs-lookup"><span data-stu-id="9cd88-185">Yes</span></span>  |  <span data-ttu-id="9cd88-p115">パラメーターの名前です。この名前は、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p115">The name of the parameter. This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="9cd88-188">文字列</span><span class="sxs-lookup"><span data-stu-id="9cd88-188">string</span></span>  |  <span data-ttu-id="9cd88-189">いいえ</span><span class="sxs-lookup"><span data-stu-id="9cd88-189">No</span></span>  |  <span data-ttu-id="9cd88-p116">パラメーターのデータ型です。 **ブール値**、 **数値**、または **文字列**である必要があります。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p116">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="result"></a><span data-ttu-id="9cd88-192">result</span><span class="sxs-lookup"><span data-stu-id="9cd88-192">result</span></span>

<span data-ttu-id="9cd88-p117">`results` オブジェクトは、関数によって返される情報の種類を定義します。次の表のプロパティの `result` オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p117">The `results` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="9cd88-195">プロパティ</span><span class="sxs-lookup"><span data-stu-id="9cd88-195">Property</span></span>  |  <span data-ttu-id="9cd88-196">データ型</span><span class="sxs-lookup"><span data-stu-id="9cd88-196">Data type</span></span>  |  <span data-ttu-id="9cd88-197">必須</span><span class="sxs-lookup"><span data-stu-id="9cd88-197">Required</span></span>  |  <span data-ttu-id="9cd88-198">説明</span><span class="sxs-lookup"><span data-stu-id="9cd88-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="9cd88-199">文字列</span><span class="sxs-lookup"><span data-stu-id="9cd88-199">string</span></span>  |  <span data-ttu-id="9cd88-200">いいえ</span><span class="sxs-lookup"><span data-stu-id="9cd88-200">No</span></span>  |  <span data-ttu-id="9cd88-201">**scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="9cd88-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="9cd88-202">文字列</span><span class="sxs-lookup"><span data-stu-id="9cd88-202">string</span></span>  |  <span data-ttu-id="9cd88-203">はい</span><span class="sxs-lookup"><span data-stu-id="9cd88-203">Yes</span></span>  |  <span data-ttu-id="9cd88-p118">パラメーターのデータ型です。 **ブール値**、 **数値**、または **文字列**である必要があります。</span><span class="sxs-lookup"><span data-stu-id="9cd88-p118">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="see-also"></a><span data-ttu-id="9cd88-206">関連項目</span><span class="sxs-lookup"><span data-stu-id="9cd88-206">See also</span></span>

* [<span data-ttu-id="9cd88-207">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="9cd88-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="9cd88-208">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="9cd88-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="9cd88-209">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="9cd88-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="9cd88-210">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="9cd88-210">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
