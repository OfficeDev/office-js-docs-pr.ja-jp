---
ms.date: 06/17/2019
description: Excel のカスタム関数のメタデータを定義します。
title: Excel のカスタム関数のメタデータ
localization_priority: Normal
ms.openlocfilehash: a7715bcdd125d44ec887f8b779ac0673b4a12af0
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059861"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="57c86-103">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="57c86-103">Custom functions metadata</span></span>

<span data-ttu-id="57c86-104">Excel アドイン内で[カスタム関数](custom-functions-overview.md)を定義する場合、アドインプロジェクトには、カスタム関数を登録してエンドユーザーが使用できるようにするために excel が必要とする情報を提供する JSON メタデータファイルが含まれています。</span><span class="sxs-lookup"><span data-stu-id="57c86-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span>

<span data-ttu-id="57c86-105">このファイルは、次のいずれかの方法で生成されます。</span><span class="sxs-lookup"><span data-stu-id="57c86-105">This file is generated either:</span></span>

- <span data-ttu-id="57c86-106">手書きの JSON ファイル</span><span class="sxs-lookup"><span data-stu-id="57c86-106">By you, in a handwritten JSON file</span></span>
- <span data-ttu-id="57c86-107">関数の先頭に入力した JSDoc コメントから</span><span class="sxs-lookup"><span data-stu-id="57c86-107">From the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="57c86-108">ユーザーが初めてアドインを実行したときに、すべてのブックの同じユーザーがそのアドインを使用できるようになると、カスタム関数が登録されます。</span><span class="sxs-lookup"><span data-stu-id="57c86-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="57c86-109">この記事では、JSON メタデータファイルの形式について説明しています (手動で記述する場合を想定しています)。</span><span class="sxs-lookup"><span data-stu-id="57c86-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="57c86-110">JSDoc comment JSON ファイル生成の詳細については、「[カスタム関数の json メタデータの生成](custom-functions-json-autogeneration.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57c86-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="57c86-111">カスタム関数を有効にするためにアドイン プロジェクトに含める必要のあるその他のファイルに関する情報については、「[Excel でのカスタム関数の作成](custom-functions-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57c86-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

<span data-ttu-id="57c86-112">JSON ファイルをホストするサーバーでは、カスタム関数を Excel Online で正しく作動させるために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="57c86-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="57c86-113">メタデータの例</span><span class="sxs-lookup"><span data-stu-id="57c86-113">Example metadata</span></span>

<span data-ttu-id="57c86-114">次の例では、カスタム関数を定義するアドインの JSON メタデータ ファイルの内容を示します。</span><span class="sxs-lookup"><span data-stu-id="57c86-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="57c86-115">この例の後に続くセクションでは、JSON の例に含まれる個々のプロパティの詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="57c86-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
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
> <span data-ttu-id="57c86-116">完全なサンプル JSON ファイルは、 [Officedev/Excel-カスタム機能](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json)GitHub リポジトリのコミット履歴で入手できます。</span><span class="sxs-lookup"><span data-stu-id="57c86-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="57c86-117">JSON を自動的に生成するようにプロジェクトが調整されているため、手書きの JSON の完全なサンプルは、プロジェクトの以前のバージョンでのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="57c86-117">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="functions"></a><span data-ttu-id="57c86-118">functions</span><span class="sxs-lookup"><span data-stu-id="57c86-118">functions</span></span> 

<span data-ttu-id="57c86-119">`functions` プロパティは、カスタム関数オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="57c86-119">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="57c86-120">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="57c86-120">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="57c86-121">プロパティ</span><span class="sxs-lookup"><span data-stu-id="57c86-121">Property</span></span>  |  <span data-ttu-id="57c86-122">データ型</span><span class="sxs-lookup"><span data-stu-id="57c86-122">Data type</span></span>  |  <span data-ttu-id="57c86-123">必須</span><span class="sxs-lookup"><span data-stu-id="57c86-123">Required</span></span>  |  <span data-ttu-id="57c86-124">説明</span><span class="sxs-lookup"><span data-stu-id="57c86-124">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="57c86-125">string</span><span class="sxs-lookup"><span data-stu-id="57c86-125">string</span></span>  |  <span data-ttu-id="57c86-126">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-126">No</span></span>  |  <span data-ttu-id="57c86-127">Excel でエンド ユーザーに表示される関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="57c86-127">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="57c86-128">たとえば、「**華氏の値を摂氏に変換する**」です。</span><span class="sxs-lookup"><span data-stu-id="57c86-128">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="57c86-129">string</span><span class="sxs-lookup"><span data-stu-id="57c86-129">string</span></span>  |   <span data-ttu-id="57c86-130">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-130">No</span></span>  |  <span data-ttu-id="57c86-131">関数に関する情報を提供する URL です </span><span class="sxs-lookup"><span data-stu-id="57c86-131">URL that provides information about the function.</span></span> <span data-ttu-id="57c86-132">(作業ウィンドウに表示されます)。たとえば、`http://contoso.com/help/convertcelsiustofahrenheit.html` です。</span><span class="sxs-lookup"><span data-stu-id="57c86-132">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span> |
| `id`     | <span data-ttu-id="57c86-133">文字列</span><span class="sxs-lookup"><span data-stu-id="57c86-133">string</span></span> | <span data-ttu-id="57c86-134">はい</span><span class="sxs-lookup"><span data-stu-id="57c86-134">Yes</span></span> | <span data-ttu-id="57c86-135">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="57c86-135">A unique ID for the function.</span></span> <span data-ttu-id="57c86-136">この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。</span><span class="sxs-lookup"><span data-stu-id="57c86-136">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="57c86-137">文字列</span><span class="sxs-lookup"><span data-stu-id="57c86-137">string</span></span>  |  <span data-ttu-id="57c86-138">はい</span><span class="sxs-lookup"><span data-stu-id="57c86-138">Yes</span></span>  |  <span data-ttu-id="57c86-139">Excel でエンド ユーザーに表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="57c86-139">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="57c86-140">Excel では、この関数名は XML マニフェスト ファイルで指定されているカスタム関数の名前空間でプレフィックスされます。</span><span class="sxs-lookup"><span data-stu-id="57c86-140">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="57c86-141">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57c86-141">object</span></span>  |  <span data-ttu-id="57c86-142">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-142">No</span></span>  |  <span data-ttu-id="57c86-143">Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="57c86-143">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="57c86-144">詳細については、[options](#options) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57c86-144">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="57c86-145">配列</span><span class="sxs-lookup"><span data-stu-id="57c86-145">array</span></span>  |  <span data-ttu-id="57c86-146">はい</span><span class="sxs-lookup"><span data-stu-id="57c86-146">Yes</span></span>  |  <span data-ttu-id="57c86-147">関数の入力パラメーターを定義する配列です。</span><span class="sxs-lookup"><span data-stu-id="57c86-147">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="57c86-148">詳細については、[parameters](#parameters) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57c86-148">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="57c86-149">object</span><span class="sxs-lookup"><span data-stu-id="57c86-149">object</span></span>  |  <span data-ttu-id="57c86-150">はい</span><span class="sxs-lookup"><span data-stu-id="57c86-150">Yes</span></span>  |  <span data-ttu-id="57c86-151">関数が返す情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="57c86-151">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="57c86-152">詳細については、[result](#result) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57c86-152">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="57c86-153">options</span><span class="sxs-lookup"><span data-stu-id="57c86-153">options</span></span>

<span data-ttu-id="57c86-154">`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="57c86-154">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="57c86-155">次の表に、`options` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="57c86-155">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="57c86-156">プロパティ</span><span class="sxs-lookup"><span data-stu-id="57c86-156">Property</span></span>  |  <span data-ttu-id="57c86-157">データ型</span><span class="sxs-lookup"><span data-stu-id="57c86-157">Data type</span></span>  |  <span data-ttu-id="57c86-158">必須</span><span class="sxs-lookup"><span data-stu-id="57c86-158">Required</span></span>  |  <span data-ttu-id="57c86-159">説明</span><span class="sxs-lookup"><span data-stu-id="57c86-159">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="57c86-160">ブール</span><span class="sxs-lookup"><span data-stu-id="57c86-160">boolean</span></span>  |  <span data-ttu-id="57c86-161">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-161">No</span></span><br/><br/><span data-ttu-id="57c86-162">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="57c86-162">Default value is `false`.</span></span>  |  <span data-ttu-id="57c86-163">`true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `CancelableInvocation` ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57c86-163">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="57c86-164">通常、取り消し可能な関数は、1つの結果を返す非同期関数で、データの要求のキャンセルを処理する必要がある場合にのみ使用されます。</span><span class="sxs-lookup"><span data-stu-id="57c86-164">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="57c86-165">関数は、ストリーミングと取り消しの両方にすることはできません。</span><span class="sxs-lookup"><span data-stu-id="57c86-165">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="57c86-166">詳細については、「[ストリーミング機能を作成する](custom-functions-web-reqs.md#make-a-streaming-function)」の最後の方にあるメモを参照してください。</span><span class="sxs-lookup"><span data-stu-id="57c86-166">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="57c86-167">ブール</span><span class="sxs-lookup"><span data-stu-id="57c86-167">boolean</span></span> | <span data-ttu-id="57c86-168">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-168">No</span></span> <br/><br/><span data-ttu-id="57c86-169">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="57c86-169">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="57c86-170">True の場合、カスタム関数は、カスタム関数を呼び出したセルのアドレスにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="57c86-170">If true, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="57c86-171">カスタム関数を呼び出したセルのアドレスを取得するには、カスタム関数で context を使用します。</span><span class="sxs-lookup"><span data-stu-id="57c86-171">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="57c86-172">詳しくは、「[カスタム関数が呼び出したセルを特定する](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="57c86-172">For more information, see [Determine which cell invoked your custom function](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span></span> <span data-ttu-id="57c86-173">カスタム関数は、streaming と requiresAddress の両方として設定することはできません。</span><span class="sxs-lookup"><span data-stu-id="57c86-173">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="57c86-174">このオプションを使用する場合、' 呼び ' パラメーターは、オプションで渡された最後のパラメーターである必要があります。</span><span class="sxs-lookup"><span data-stu-id="57c86-174">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="57c86-175">ブール</span><span class="sxs-lookup"><span data-stu-id="57c86-175">boolean</span></span>  |  <span data-ttu-id="57c86-176">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-176">No</span></span><br/><br/><span data-ttu-id="57c86-177">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="57c86-177">Default value is `false`.</span></span>  |  <span data-ttu-id="57c86-178">`true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。</span><span class="sxs-lookup"><span data-stu-id="57c86-178">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="57c86-179">このオプションは、株価などの急速に変化するデータ ソースに便利です。</span><span class="sxs-lookup"><span data-stu-id="57c86-179">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="57c86-180">この関数には、`return` ステートメントは含めないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="57c86-180">The function should have no `return` statement.</span></span> <span data-ttu-id="57c86-181">代わりに、結果の値は `StreamingInvocation.setResult` コールバック メソッドの引数として渡されます。</span><span class="sxs-lookup"><span data-stu-id="57c86-181">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="57c86-182">詳細については、「[ストリーミング関数](custom-functions-web-reqs.md#make-a-streaming-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57c86-182">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `volatile`  | <span data-ttu-id="57c86-183">ブール</span><span class="sxs-lookup"><span data-stu-id="57c86-183">boolean</span></span> | <span data-ttu-id="57c86-184">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-184">No</span></span> <br/><br/><span data-ttu-id="57c86-185">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="57c86-185">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="57c86-186">`true` の場合は、数式の依存値が変更されたときのみではなく、Excel が再計算するたびに関数が再計算されます。</span><span class="sxs-lookup"><span data-stu-id="57c86-186">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="57c86-187">関数は、ストリーミングと揮発性の両方にすることはできません。</span><span class="sxs-lookup"><span data-stu-id="57c86-187">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="57c86-188">`stream` と `volatile` の両方のプロパティが `true` に設定されている場合は、揮発性のオプションが無視されます。</span><span class="sxs-lookup"><span data-stu-id="57c86-188">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="57c86-189">parameters</span><span class="sxs-lookup"><span data-stu-id="57c86-189">parameters</span></span>

<span data-ttu-id="57c86-190">`parameters` プロパティは、パラメーター オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="57c86-190">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="57c86-191">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="57c86-191">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="57c86-192">プロパティ</span><span class="sxs-lookup"><span data-stu-id="57c86-192">Property</span></span>  |  <span data-ttu-id="57c86-193">データ型</span><span class="sxs-lookup"><span data-stu-id="57c86-193">Data type</span></span>  |  <span data-ttu-id="57c86-194">必須</span><span class="sxs-lookup"><span data-stu-id="57c86-194">Required</span></span>  |  <span data-ttu-id="57c86-195">説明</span><span class="sxs-lookup"><span data-stu-id="57c86-195">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="57c86-196">string</span><span class="sxs-lookup"><span data-stu-id="57c86-196">string</span></span>  |  <span data-ttu-id="57c86-197">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-197">No</span></span> |  <span data-ttu-id="57c86-198">パラメーターの説明です。</span><span class="sxs-lookup"><span data-stu-id="57c86-198">A description of the parameter.</span></span> <span data-ttu-id="57c86-199">これは、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="57c86-199">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="57c86-200">string</span><span class="sxs-lookup"><span data-stu-id="57c86-200">string</span></span>  |  <span data-ttu-id="57c86-201">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-201">No</span></span>  |  <span data-ttu-id="57c86-202">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="57c86-202">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="57c86-203">文字列</span><span class="sxs-lookup"><span data-stu-id="57c86-203">string</span></span>  |  <span data-ttu-id="57c86-204">はい</span><span class="sxs-lookup"><span data-stu-id="57c86-204">Yes</span></span>  |  <span data-ttu-id="57c86-205">パラメーターの名前です。</span><span class="sxs-lookup"><span data-stu-id="57c86-205">The name of the parameter.</span></span> <span data-ttu-id="57c86-206">この名前は、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="57c86-206">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="57c86-207">文字列</span><span class="sxs-lookup"><span data-stu-id="57c86-207">string</span></span>  |  <span data-ttu-id="57c86-208">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-208">No</span></span>  |  <span data-ttu-id="57c86-209">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="57c86-209">The data type of the parameter.</span></span> <span data-ttu-id="57c86-210">**boolean**、**number**、**string**、または **any** が可能です。ここでは、前の 3 種類のいずれかを使用できます。</span><span class="sxs-lookup"><span data-stu-id="57c86-210">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="57c86-211">このプロパティが指定されていない場合、データ型の既定は **any** です。</span><span class="sxs-lookup"><span data-stu-id="57c86-211">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="57c86-212">ブール</span><span class="sxs-lookup"><span data-stu-id="57c86-212">boolean</span></span> | <span data-ttu-id="57c86-213">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-213">No</span></span> | <span data-ttu-id="57c86-214">`true` の場合、パラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="57c86-214">If `true`, the parameter is optional.</span></span> |

## <a name="result"></a><span data-ttu-id="57c86-215">result</span><span class="sxs-lookup"><span data-stu-id="57c86-215">result</span></span>

<span data-ttu-id="57c86-216">`result` オブジェクトは、この関数が返す情報の種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="57c86-216">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="57c86-217">次の表に、`result` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="57c86-217">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="57c86-218">プロパティ</span><span class="sxs-lookup"><span data-stu-id="57c86-218">Property</span></span>  |  <span data-ttu-id="57c86-219">データ型</span><span class="sxs-lookup"><span data-stu-id="57c86-219">Data type</span></span>  |  <span data-ttu-id="57c86-220">必須</span><span class="sxs-lookup"><span data-stu-id="57c86-220">Required</span></span>  |  <span data-ttu-id="57c86-221">説明</span><span class="sxs-lookup"><span data-stu-id="57c86-221">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="57c86-222">string</span><span class="sxs-lookup"><span data-stu-id="57c86-222">string</span></span>  |  <span data-ttu-id="57c86-223">いいえ</span><span class="sxs-lookup"><span data-stu-id="57c86-223">No</span></span>  |  <span data-ttu-id="57c86-224">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="57c86-224">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="next-steps"></a><span data-ttu-id="57c86-225">次のステップ</span><span class="sxs-lookup"><span data-stu-id="57c86-225">Next steps</span></span>
<span data-ttu-id="57c86-226">[関数に名前を付けるためのベストプラクティス](custom-functions-naming.md)、または前述の手書き JSON メソッドを使用して[関数をローカライズ](custom-functions-localize.md)する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="57c86-226">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="57c86-227">関連項目</span><span class="sxs-lookup"><span data-stu-id="57c86-227">See also</span></span>

* [<span data-ttu-id="57c86-228">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="57c86-228">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="57c86-229">カスタム関数のパラメータオプション</span><span class="sxs-lookup"><span data-stu-id="57c86-229">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* [<span data-ttu-id="57c86-230">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="57c86-230">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="57c86-231">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="57c86-231">Create custom functions in Excel</span></span>](custom-functions-overview.md)