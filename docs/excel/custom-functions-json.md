---
ms.date: 09/27/2018
description: Excel でカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ
ms.openlocfilehash: e8af13b8855d6c5e1a3b1ce99edb24445e066756
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459239"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="43e3a-103">カスタム関数のメタデータ (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="43e3a-103">Custom functions metadata</span></span>

<span data-ttu-id="43e3a-104">Excel アドインで[カスタム関数](custom-functions-overview.md) を定義する場合には、Excel でカスタム関数を登録してエンドユーザーが使用できるようにするための情報を提供する JSON メタデータ ファイルを、アドイン プロジェクトに含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="43e3a-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="43e3a-105">この記事では、JSON メタデータ ファイルの形式について説明します。</span><span class="sxs-lookup"><span data-stu-id="43e3a-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="43e3a-106">カスタム関数を有効にするためにアドイン プロジェクトに含める必要のある、その他のファイルに関する情報については、「[Excel でカスタム関数を作成する](custom-functions-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="43e3a-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="43e3a-107">メタデータの例</span><span class="sxs-lookup"><span data-stu-id="43e3a-107">Example metadata</span></span>

<span data-ttu-id="43e3a-p102">次の使用例は、JSON のアドインでカスタム関数を定義するメタデータ ファイルの内容を示しています。次の使用例を次のセクションでは、この例を JSON 内の個別のプロパティに関する詳細情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p102">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions. The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="43e3a-110">JSON ファイルの完全なサンプルは、「[OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) 」GitHub リポジトリで利用可能です。</span><span class="sxs-lookup"><span data-stu-id="43e3a-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="43e3a-111">functions</span><span class="sxs-lookup"><span data-stu-id="43e3a-111">functions</span></span> 

<span data-ttu-id="43e3a-p103"> `functions` プロパティは、カスタム関数オブジェクトの配列です。次の表は、各オブジェクトのプロパティを一覧表示しています。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p103">The `functions` property is an array of custom function objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="43e3a-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="43e3a-114">Property</span></span>  |  <span data-ttu-id="43e3a-115">データ型</span><span class="sxs-lookup"><span data-stu-id="43e3a-115">Data type</span></span>  |  <span data-ttu-id="43e3a-116">必須</span><span class="sxs-lookup"><span data-stu-id="43e3a-116">Required</span></span>  |  <span data-ttu-id="43e3a-117">説明</span><span class="sxs-lookup"><span data-stu-id="43e3a-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="43e3a-118">文字列</span><span class="sxs-lookup"><span data-stu-id="43e3a-118">string</span></span>  |  <span data-ttu-id="43e3a-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="43e3a-119">No</span></span>  |  <span data-ttu-id="43e3a-120">Excel でエンド ユーザーに表示される関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="43e3a-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="43e3a-121">たとえば、「**摂氏の値を華氏に変換する**」などです。</span><span class="sxs-lookup"><span data-stu-id="43e3a-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="43e3a-122">文字列</span><span class="sxs-lookup"><span data-stu-id="43e3a-122">string</span></span>  |   <span data-ttu-id="43e3a-123">いいえ</span><span class="sxs-lookup"><span data-stu-id="43e3a-123">No</span></span>  |  <span data-ttu-id="43e3a-124">関数に関する情報を提供する URL です。</span><span class="sxs-lookup"><span data-stu-id="43e3a-124">URL that provides information about the function.</span></span> <span data-ttu-id="43e3a-125">(作業ウィンドウに表示されます。) たとえば、 **http://contoso.com/help/convertcelsiustofahrenheit.html**です。</span><span class="sxs-lookup"><span data-stu-id="43e3a-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="43e3a-126">文字列</span><span class="sxs-lookup"><span data-stu-id="43e3a-126">string</span></span> | <span data-ttu-id="43e3a-127">はい</span><span class="sxs-lookup"><span data-stu-id="43e3a-127">Yes</span></span> | <span data-ttu-id="43e3a-p106">関数の一意の ID です。設定後、この ID は変更できません。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p106">A unique ID for the function. This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="43e3a-130">文字列</span><span class="sxs-lookup"><span data-stu-id="43e3a-130">string</span></span>  |  <span data-ttu-id="43e3a-131">はい</span><span class="sxs-lookup"><span data-stu-id="43e3a-131">Yes</span></span>  |  <span data-ttu-id="43e3a-p107">Excel でエンド ユーザーに表示される関数の名前です。Excel では、この関数名が XML マニフェスト ファイルで指定されているカスタム関数の名前空間で接頭辞となります。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p107">The name of the function that end users see in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="43e3a-134">object</span><span class="sxs-lookup"><span data-stu-id="43e3a-134">object</span></span>  |  <span data-ttu-id="43e3a-135">いいえ</span><span class="sxs-lookup"><span data-stu-id="43e3a-135">No</span></span>  |  <span data-ttu-id="43e3a-p108">Excel で関数を実行する方法とタイミングのいくつかの側面をカスタマイズできます。詳細については、 [オプションのオブジェクト](#options-object) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p108">Enables you to customize some aspects of how and when Excel executes the function. See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="43e3a-138">配列</span><span class="sxs-lookup"><span data-stu-id="43e3a-138">array</span></span>  |  <span data-ttu-id="43e3a-139">はい</span><span class="sxs-lookup"><span data-stu-id="43e3a-139">Yes</span></span>  |  <span data-ttu-id="43e3a-p109">関数の入力パラメーターを定義する配列。詳細については、 [パラメーター配列](#parameters-array) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p109">Array that defines the input parameters for the function. See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="43e3a-142">object</span><span class="sxs-lookup"><span data-stu-id="43e3a-142">object</span></span>  |  <span data-ttu-id="43e3a-143">はい</span><span class="sxs-lookup"><span data-stu-id="43e3a-143">Yes</span></span>  |  <span data-ttu-id="43e3a-p110">関数によって返される情報の種類を定義するオブジェクト。詳細については、 [結果のオブジェクト](#result-object) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p110">Object that defines the type of information that is returned by the function. See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="43e3a-146">options</span><span class="sxs-lookup"><span data-stu-id="43e3a-146">options</span></span>

<span data-ttu-id="43e3a-p111"> `options` オブジェクトは、Excel で関数を実行する方法とタイミングのいくつかの側面をカスタマイズすることができます。次の表に`options` オブジェクトのプロパティを一覧表示しています。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p111">The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="43e3a-149">プロパティ</span><span class="sxs-lookup"><span data-stu-id="43e3a-149">Property</span></span>  |  <span data-ttu-id="43e3a-150">データ型</span><span class="sxs-lookup"><span data-stu-id="43e3a-150">Data type</span></span>  |  <span data-ttu-id="43e3a-151">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="43e3a-151">Required</span></span>  |  <span data-ttu-id="43e3a-152">説明</span><span class="sxs-lookup"><span data-stu-id="43e3a-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="43e3a-153">boolean</span><span class="sxs-lookup"><span data-stu-id="43e3a-153">boolean</span></span>  |  <span data-ttu-id="43e3a-154">いいえ</span><span class="sxs-lookup"><span data-stu-id="43e3a-154">No</span></span><br/><br/><span data-ttu-id="43e3a-155">既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="43e3a-155">Default value is 4.</span></span>  |  <span data-ttu-id="43e3a-p112">`true` を使用する場合、関数をキャンセルすることになる操作をユーザーが実行するたびに Excel は、 `onCanceled` ハンドラーを呼び出します。例えば、手動で再計算をトリガーしたり、関数が参照しているセルを編集したりなどの操作です。このオプションを使用する場合、Excel は、`caller` パラメータを追加して、JavaScript 関数を呼び出します 。(`parameters` プロパティにこのパラメータを登録***しない***でください )。関数の本文では、`caller.onCanceled` のメンバーにハンドラーを割り当てる必要があります。詳細については、 [関数をキャンセルする](custom-functions-overview.md#canceling-a-function)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p112">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="43e3a-161">ブール値</span><span class="sxs-lookup"><span data-stu-id="43e3a-161">boolean</span></span>  |  <span data-ttu-id="43e3a-162">いいえ</span><span class="sxs-lookup"><span data-stu-id="43e3a-162">No</span></span><br/><br/><span data-ttu-id="43e3a-163">既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="43e3a-163">Default value is 4.</span></span>  |  <span data-ttu-id="43e3a-p113">`true`を使用する場合、1 回だけ呼び出される場合でも、セルに関数を繰り返し出力できます。このオプションは、急速に変化するデータ ソース、株価などに便利です。このオプションを使用する場合、Excel は、`caller`  パラメータを追加して、JavaScript 関数を呼び出します 。( `parameters` プロパティにこのパラメータを登録\*\*\* しない\*\*\*でください )。関数には `return` 文を入れてはいけません。代わりに、結果の値が`caller.setResult` コールバック メソッドの引数として渡されます。詳細については、 [ストリーミング機能](custom-functions-overview.md#streaming-functions)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p113">If `true`, the function can output repeatedly to the cell even when invoked only once. This option is useful for rapidly-changing data sources, such as a stock price. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). The function should have no `return` statement. Instead, the result value is passed as the argument of the `caller.setResult` callback method. For more information, see [Streamed functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="43e3a-171">parameters</span><span class="sxs-lookup"><span data-stu-id="43e3a-171">parameters</span></span>

<span data-ttu-id="43e3a-p114">`parameters` プロパティは、カスタム関数オブジェクトの配列です。次の表は、各オブジェクトのプロパティを一覧表示しています。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p114">The `parameters` property is an array of parameter objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="43e3a-174">プロパティ</span><span class="sxs-lookup"><span data-stu-id="43e3a-174">Property</span></span>  |  <span data-ttu-id="43e3a-175">データ型</span><span class="sxs-lookup"><span data-stu-id="43e3a-175">Data type</span></span>  |  <span data-ttu-id="43e3a-176">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="43e3a-176">Required</span></span>  |  <span data-ttu-id="43e3a-177">説明</span><span class="sxs-lookup"><span data-stu-id="43e3a-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="43e3a-178">文字列</span><span class="sxs-lookup"><span data-stu-id="43e3a-178">string</span></span>  |  <span data-ttu-id="43e3a-179">いいえ</span><span class="sxs-lookup"><span data-stu-id="43e3a-179">No</span></span> |  <span data-ttu-id="43e3a-180">パラメータの説明。</span><span class="sxs-lookup"><span data-stu-id="43e3a-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="43e3a-181">文字列</span><span class="sxs-lookup"><span data-stu-id="43e3a-181">string</span></span>  |  <span data-ttu-id="43e3a-182">いいえ</span><span class="sxs-lookup"><span data-stu-id="43e3a-182">No</span></span>  |  <span data-ttu-id="43e3a-183">**scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="43e3a-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="43e3a-184">文字列</span><span class="sxs-lookup"><span data-stu-id="43e3a-184">string</span></span>  |  <span data-ttu-id="43e3a-185">はい</span><span class="sxs-lookup"><span data-stu-id="43e3a-185">Yes</span></span>  |  <span data-ttu-id="43e3a-p115">パラメーターの名前です。この名前は、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p115">The name of the parameter. This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="43e3a-188">文字列</span><span class="sxs-lookup"><span data-stu-id="43e3a-188">string</span></span>  |  <span data-ttu-id="43e3a-189">いいえ</span><span class="sxs-lookup"><span data-stu-id="43e3a-189">No</span></span>  |  <span data-ttu-id="43e3a-p116">パラメーターのデータ型です。 **ブール値**、 **数値**、または **文字列**である必要があります。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p116">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="result"></a><span data-ttu-id="43e3a-192">result</span><span class="sxs-lookup"><span data-stu-id="43e3a-192">result</span></span>

<span data-ttu-id="43e3a-p117"> `results` オブジェクトは、関数によって返される情報の種類を定義します。次の表のプロパティの `result` オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p117">The `results` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="43e3a-195">プロパティ</span><span class="sxs-lookup"><span data-stu-id="43e3a-195">Property</span></span>  |  <span data-ttu-id="43e3a-196">データ型</span><span class="sxs-lookup"><span data-stu-id="43e3a-196">Data type</span></span>  |  <span data-ttu-id="43e3a-197">必須かどうか</span><span class="sxs-lookup"><span data-stu-id="43e3a-197">Required</span></span>  |  <span data-ttu-id="43e3a-198">説明</span><span class="sxs-lookup"><span data-stu-id="43e3a-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="43e3a-199">文字列</span><span class="sxs-lookup"><span data-stu-id="43e3a-199">string</span></span>  |  <span data-ttu-id="43e3a-200">いいえ</span><span class="sxs-lookup"><span data-stu-id="43e3a-200">No</span></span>  |  <span data-ttu-id="43e3a-201">**scholar** (非配列値) または **matrix** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="43e3a-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="43e3a-202">文字列</span><span class="sxs-lookup"><span data-stu-id="43e3a-202">string</span></span>  |  <span data-ttu-id="43e3a-203">はい</span><span class="sxs-lookup"><span data-stu-id="43e3a-203">Yes</span></span>  |  <span data-ttu-id="43e3a-p118">パラメーターのデータ型です。 **ブール値**、 **数値**、または **文字列**である必要があります。</span><span class="sxs-lookup"><span data-stu-id="43e3a-p118">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="see-also"></a><span data-ttu-id="43e3a-206">関連項目</span><span class="sxs-lookup"><span data-stu-id="43e3a-206">See also</span></span>

* [<span data-ttu-id="43e3a-207">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="43e3a-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="43e3a-208">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="43e3a-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="43e3a-209">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="43e3a-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="43e3a-210">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="43e3a-210">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)