---
ms.date: 07/15/2019
description: Excel でカスタム関数の JSON メタデータを定義し、関数 id と name プロパティを関連付けます。
title: Excel のカスタム関数のメタデータ
localization_priority: Normal
ms.openlocfilehash: b0e015cfa439651420487db4885647f5c7de7da8
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771562"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="88bb6-103">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="88bb6-103">Custom functions metadata</span></span>

<span data-ttu-id="88bb6-104">カスタム関数の[概要](custom-functions-overview.md)の記事で説明されているように、カスタム関数プロジェクトには、JSON メタデータファイルとスクリプト (JavaScript または TypeScript) の両方を含める必要があります。関数を登録するには、このファイルを使用できるようにします。</span><span class="sxs-lookup"><span data-stu-id="88bb6-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="88bb6-105">ユーザーが初めてアドインを実行したときに、すべてのブックの同じユーザーがそのアドインを使用できるようになると、カスタム関数が登録されます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="88bb6-106">`yo office`スキャフォールディングファイルを使用することをお勧めします。このプロセスは、 [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)に示されているプロセスと同様に、ユーザーエラーが発生しやすくなります。</span><span class="sxs-lookup"><span data-stu-id="88bb6-106">It is recommended that you use JSON autogeneration when possible, using the `yo office` scaffold files, similar to the process shown in the [Excel Custom Function tutorial](../tutorials/excel-tutorial-create-custom-functions.md) because this process is easier and less prone to user error.</span></span> <span data-ttu-id="88bb6-107">JSDoc comment JSON ファイル生成のプロセスの詳細については、「[カスタム関数の json メタデータの生成](custom-functions-json-autogeneration.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-107">For more information on the process of JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="88bb6-108">ただし、カスタム関数プロジェクトを最初から作成できます。そのためには、次のことを行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="88bb6-108">However, you can make a custom functions project from scratch; it requires that you:</span></span>

- <span data-ttu-id="88bb6-109">JSON ファイルを手動で記述する</span><span class="sxs-lookup"><span data-stu-id="88bb6-109">Write your JSON file by hand</span></span>
- <span data-ttu-id="88bb6-110">マニフェストファイルが手動で作成した JSON ファイルに接続されていることを確認する</span><span class="sxs-lookup"><span data-stu-id="88bb6-110">Check that your manifest file is connected to your hand-authored JSON file</span></span>
- <span data-ttu-id="88bb6-111">関数を登録する`id`ため`name`に、スクリプトファイルの関数とプロパティを関連付けます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-111">Associate your functions' `id` and `name` properties in the script file in order to register your functions</span></span>

<span data-ttu-id="88bb6-112">この記事では、これら3つの手順をすべて実行する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-112">This article will show you how to do all three of these steps.</span></span>

> [!NOTE]
> <span data-ttu-id="88bb6-113">スキャフォールディングファイルとは`yo office`異なり、XML マニフェストファイルの`<Resources>`セクションを使用して、作成した JSON ファイルを使用してマニフェストをフックする必要があります。</span><span class="sxs-lookup"><span data-stu-id="88bb6-113">In contrast with the `yo office` scaffold files, you need to hook your manifest up with the JSON file you create, through the `<Resources>` section in your XML manifest file.</span></span> <span data-ttu-id="88bb6-114">Web 上の Excel でカスタム関数が正しく動作するためには、JSON ファイルをホストするサーバー上のサーバー設定で[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS)を有効にする必要があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-114">Note that the server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel on the web.</span></span>

## <a name="authoring-metadata-and-hooking-up-to-the-manifest"></a><span data-ttu-id="88bb6-115">メタデータの作成とマニフェストへのフック</span><span class="sxs-lookup"><span data-stu-id="88bb6-115">Authoring metadata and hooking up to the manifest</span></span>

<span data-ttu-id="88bb6-116">プロジェクトで JSON ファイルを作成し、関数のパラメーターなど、関数に関するすべての詳細を提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="88bb6-116">You need to create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="88bb6-117">関数プロパティの完全なリストについては、[次のメタデータの例](#json-metadata-example)と[メタデータリファレンス](#metadata-reference)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-117">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="88bb6-118">また、次の例に示すように、XML マニフェストファイルが JSON ファイル`<Resources>`を参照していることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="88bb6-118">You also need to make sure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## <a name="json-metadata-example"></a><span data-ttu-id="88bb6-119">JSON メタデータの例</span><span class="sxs-lookup"><span data-stu-id="88bb6-119">JSON metadata example</span></span>

<span data-ttu-id="88bb6-120">次の例では、カスタム関数を定義するアドインの JSON メタデータ ファイルの内容を示します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-120">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="88bb6-121">この例の後に続くセクションでは、JSON の例に含まれる個々のプロパティの詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-121">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
      "description": "Count up from zero",
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
      "description": "Get the second highest number from a range",
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
> <span data-ttu-id="88bb6-122">完全なサンプル JSON ファイルは、 [Officedev/Excel-カスタム機能](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json)GitHub リポジトリのコミット履歴で入手できます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-122">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="88bb6-123">JSON を自動的に生成するようにプロジェクトが調整されているため、手書きの JSON の完全なサンプルは、プロジェクトの以前のバージョンでのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-123">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="88bb6-124">メタデータリファレンス</span><span class="sxs-lookup"><span data-stu-id="88bb6-124">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="88bb6-125">functions</span><span class="sxs-lookup"><span data-stu-id="88bb6-125">functions</span></span>

<span data-ttu-id="88bb6-126">`functions` プロパティは、カスタム関数オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-126">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="88bb6-127">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-127">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="88bb6-128">プロパティ</span><span class="sxs-lookup"><span data-stu-id="88bb6-128">Property</span></span>      | <span data-ttu-id="88bb6-129">データ型</span><span class="sxs-lookup"><span data-stu-id="88bb6-129">Data type</span></span> | <span data-ttu-id="88bb6-130">必須</span><span class="sxs-lookup"><span data-stu-id="88bb6-130">Required</span></span> | <span data-ttu-id="88bb6-131">説明</span><span class="sxs-lookup"><span data-stu-id="88bb6-131">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="88bb6-132">string</span><span class="sxs-lookup"><span data-stu-id="88bb6-132">string</span></span>    | <span data-ttu-id="88bb6-133">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-133">No</span></span>       | <span data-ttu-id="88bb6-134">Excel でエンド ユーザーに表示される関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-134">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="88bb6-135">たとえば、「**華氏の値を摂氏に変換する**」です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-135">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="88bb6-136">string</span><span class="sxs-lookup"><span data-stu-id="88bb6-136">string</span></span>    | <span data-ttu-id="88bb6-137">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-137">No</span></span>       | <span data-ttu-id="88bb6-138">関数に関する情報を提供する URL です </span><span class="sxs-lookup"><span data-stu-id="88bb6-138">URL that provides information about the function.</span></span> <span data-ttu-id="88bb6-139">(作業ウィンドウに表示されます)。たとえば、`http://contoso.com/help/convertcelsiustofahrenheit.html` です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-139">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="88bb6-140">文字列</span><span class="sxs-lookup"><span data-stu-id="88bb6-140">string</span></span>    | <span data-ttu-id="88bb6-141">はい</span><span class="sxs-lookup"><span data-stu-id="88bb6-141">Yes</span></span>      | <span data-ttu-id="88bb6-142">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-142">A unique ID for the function.</span></span> <span data-ttu-id="88bb6-143">この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。</span><span class="sxs-lookup"><span data-stu-id="88bb6-143">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="88bb6-144">文字列</span><span class="sxs-lookup"><span data-stu-id="88bb6-144">string</span></span>    | <span data-ttu-id="88bb6-145">はい</span><span class="sxs-lookup"><span data-stu-id="88bb6-145">Yes</span></span>      | <span data-ttu-id="88bb6-146">Excel でエンド ユーザーに表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-146">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="88bb6-147">Excel では、この関数名は XML マニフェスト ファイルで指定されているカスタム関数の名前空間でプレフィックスされます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-147">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="88bb6-148">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="88bb6-148">object</span></span>    | <span data-ttu-id="88bb6-149">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-149">No</span></span>       | <span data-ttu-id="88bb6-150">Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-150">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="88bb6-151">詳細については、[options](#options) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-151">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="88bb6-152">配列</span><span class="sxs-lookup"><span data-stu-id="88bb6-152">array</span></span>     | <span data-ttu-id="88bb6-153">はい</span><span class="sxs-lookup"><span data-stu-id="88bb6-153">Yes</span></span>      | <span data-ttu-id="88bb6-154">関数の入力パラメーターを定義する配列です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-154">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="88bb6-155">詳細については、「 [parameters](#parameters) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-155">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="88bb6-156">object</span><span class="sxs-lookup"><span data-stu-id="88bb6-156">object</span></span>    | <span data-ttu-id="88bb6-157">はい</span><span class="sxs-lookup"><span data-stu-id="88bb6-157">Yes</span></span>      | <span data-ttu-id="88bb6-158">関数が返す情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="88bb6-158">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="88bb6-159">詳細については、[result](#result) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-159">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="88bb6-160">options</span><span class="sxs-lookup"><span data-stu-id="88bb6-160">options</span></span>

<span data-ttu-id="88bb6-161">`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-161">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="88bb6-162">次の表に、`options` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-162">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="88bb6-163">プロパティ</span><span class="sxs-lookup"><span data-stu-id="88bb6-163">Property</span></span>          | <span data-ttu-id="88bb6-164">データ型</span><span class="sxs-lookup"><span data-stu-id="88bb6-164">Data type</span></span> | <span data-ttu-id="88bb6-165">必須</span><span class="sxs-lookup"><span data-stu-id="88bb6-165">Required</span></span>                               | <span data-ttu-id="88bb6-166">説明</span><span class="sxs-lookup"><span data-stu-id="88bb6-166">Description</span></span>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| :---------------- | :-------- | :------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cancelable`      | <span data-ttu-id="88bb6-167">ブール</span><span class="sxs-lookup"><span data-stu-id="88bb6-167">boolean</span></span>   | <span data-ttu-id="88bb6-168">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-168">No</span></span><br/><br/><span data-ttu-id="88bb6-169">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-169">Default value is `false`.</span></span>  | <span data-ttu-id="88bb6-170">`true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `CancelableInvocation` ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-170">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="88bb6-171">通常、取り消し可能な関数は、1つの結果を返す非同期関数で、データの要求のキャンセルを処理する必要がある場合にのみ使用されます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-171">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="88bb6-172">関数は、ストリーミングと取り消しの両方にすることはできません。</span><span class="sxs-lookup"><span data-stu-id="88bb6-172">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="88bb6-173">詳細については、「[ストリーミング機能を作成する](custom-functions-web-reqs.md#make-a-streaming-function)」の最後の方にあるメモを参照してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-173">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `requiresAddress` | <span data-ttu-id="88bb6-174">ブール</span><span class="sxs-lookup"><span data-stu-id="88bb6-174">boolean</span></span>   | <span data-ttu-id="88bb6-175">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-175">No</span></span> <br/><br/><span data-ttu-id="88bb6-176">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-176">Default value is `false`.</span></span> | <span data-ttu-id="88bb6-177">の`true`場合は、カスタム関数を呼び出したセルのアドレスにカスタム関数からアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-177">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="88bb6-178">カスタム関数を呼び出したセルのアドレスを取得するには、カスタム関数で context を使用します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-178">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="88bb6-179">詳細については、「[アドレス指定セルのコンテキストパラメーター](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-179">For more information, see [Addressing cell's context parameter](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter).</span></span> <span data-ttu-id="88bb6-180">カスタム関数は、streaming と requiresAddress の両方として設定することはできません。</span><span class="sxs-lookup"><span data-stu-id="88bb6-180">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="88bb6-181">このオプションを使用する場合、' 呼び ' パラメーターは、オプションで渡された最後のパラメーターである必要があります。</span><span class="sxs-lookup"><span data-stu-id="88bb6-181">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span>                                              |
| `stream`          | <span data-ttu-id="88bb6-182">ブール</span><span class="sxs-lookup"><span data-stu-id="88bb6-182">boolean</span></span>   | <span data-ttu-id="88bb6-183">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-183">No</span></span><br/><br/><span data-ttu-id="88bb6-184">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-184">Default value is `false`.</span></span>  | <span data-ttu-id="88bb6-185">`true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-185">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="88bb6-186">このオプションは、株価などの急速に変化するデータ ソースに便利です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-186">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="88bb6-187">この関数には、`return` ステートメントは含めないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="88bb6-187">The function should have no `return` statement.</span></span> <span data-ttu-id="88bb6-188">代わりに、結果の値は `StreamingInvocation.setResult` コールバック メソッドの引数として渡されます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-188">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="88bb6-189">詳細については、「[ストリーミング関数](custom-functions-web-reqs.md#make-a-streaming-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-189">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>                                                                                                                                                                |
| `volatile`        | <span data-ttu-id="88bb6-190">ブール</span><span class="sxs-lookup"><span data-stu-id="88bb6-190">boolean</span></span>   | <span data-ttu-id="88bb6-191">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-191">No</span></span> <br/><br/><span data-ttu-id="88bb6-192">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-192">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="88bb6-193">`true` の場合は、数式の依存値が変更されたときのみではなく、Excel が再計算するたびに関数が再計算されます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-193">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="88bb6-194">関数は、ストリーミングと揮発性の両方にすることはできません。</span><span class="sxs-lookup"><span data-stu-id="88bb6-194">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="88bb6-195">`stream` と `volatile` の両方のプロパティが `true` に設定されている場合は、揮発性のオプションが無視されます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-195">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span>                                                                                                                                                                                                                                                                                             |

### <a name="parameters"></a><span data-ttu-id="88bb6-196">parameters</span><span class="sxs-lookup"><span data-stu-id="88bb6-196">parameters</span></span>

<span data-ttu-id="88bb6-197">`parameters` プロパティは、パラメーター オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-197">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="88bb6-198">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-198">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="88bb6-199">プロパティ</span><span class="sxs-lookup"><span data-stu-id="88bb6-199">Property</span></span>  |  <span data-ttu-id="88bb6-200">データ型</span><span class="sxs-lookup"><span data-stu-id="88bb6-200">Data type</span></span>  |  <span data-ttu-id="88bb6-201">必須</span><span class="sxs-lookup"><span data-stu-id="88bb6-201">Required</span></span>  |  <span data-ttu-id="88bb6-202">説明</span><span class="sxs-lookup"><span data-stu-id="88bb6-202">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="88bb6-203">string</span><span class="sxs-lookup"><span data-stu-id="88bb6-203">string</span></span>  |  <span data-ttu-id="88bb6-204">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-204">No</span></span> |  <span data-ttu-id="88bb6-205">パラメーターの説明です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-205">A description of the parameter.</span></span> <span data-ttu-id="88bb6-206">これは、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-206">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="88bb6-207">string</span><span class="sxs-lookup"><span data-stu-id="88bb6-207">string</span></span>  |  <span data-ttu-id="88bb6-208">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-208">No</span></span>  |  <span data-ttu-id="88bb6-209">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="88bb6-209">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="88bb6-210">文字列</span><span class="sxs-lookup"><span data-stu-id="88bb6-210">string</span></span>  |  <span data-ttu-id="88bb6-211">はい</span><span class="sxs-lookup"><span data-stu-id="88bb6-211">Yes</span></span>  |  <span data-ttu-id="88bb6-212">パラメーターの名前です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-212">The name of the parameter.</span></span> <span data-ttu-id="88bb6-213">この名前は、Excel の intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-213">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="88bb6-214">文字列</span><span class="sxs-lookup"><span data-stu-id="88bb6-214">string</span></span>  |  <span data-ttu-id="88bb6-215">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-215">No</span></span>  |  <span data-ttu-id="88bb6-216">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-216">The data type of the parameter.</span></span> <span data-ttu-id="88bb6-217">**boolean**、**number**、**string**、または **any** が可能です。ここでは、前の 3 種類のいずれかを使用できます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-217">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="88bb6-218">このプロパティが指定されていない場合、データ型の既定は **any** です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-218">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="88bb6-219">ブール</span><span class="sxs-lookup"><span data-stu-id="88bb6-219">boolean</span></span> | <span data-ttu-id="88bb6-220">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-220">No</span></span> | <span data-ttu-id="88bb6-221">`true` の場合、パラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="88bb6-221">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="88bb6-222">ブール</span><span class="sxs-lookup"><span data-stu-id="88bb6-222">boolean</span></span> | <span data-ttu-id="88bb6-223">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-223">No</span></span> | <span data-ttu-id="88bb6-224">の`true`場合は、パラメーターが指定された配列から設定されます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-224">If `true`, parameters will populate from a specified array.</span></span> <span data-ttu-id="88bb6-225">すべての繰り返しパラメーターは、定義によって省略可能なパラメーターとして扱われることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-225">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="88bb6-226">result</span><span class="sxs-lookup"><span data-stu-id="88bb6-226">result</span></span>

<span data-ttu-id="88bb6-227">`result` オブジェクトは、この関数が返す情報の種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-227">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="88bb6-228">次の表に、`result` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-228">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="88bb6-229">プロパティ</span><span class="sxs-lookup"><span data-stu-id="88bb6-229">Property</span></span>         | <span data-ttu-id="88bb6-230">データ型</span><span class="sxs-lookup"><span data-stu-id="88bb6-230">Data type</span></span> | <span data-ttu-id="88bb6-231">必須</span><span class="sxs-lookup"><span data-stu-id="88bb6-231">Required</span></span> | <span data-ttu-id="88bb6-232">説明</span><span class="sxs-lookup"><span data-stu-id="88bb6-232">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="88bb6-233">string</span><span class="sxs-lookup"><span data-stu-id="88bb6-233">string</span></span>    | <span data-ttu-id="88bb6-234">いいえ</span><span class="sxs-lookup"><span data-stu-id="88bb6-234">No</span></span>       | <span data-ttu-id="88bb6-235">**スカラー** (配列以外の値) または**マトリックス** (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="88bb6-235">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="88bb6-236">関数名を JSON メタデータに関連付ける</span><span class="sxs-lookup"><span data-stu-id="88bb6-236">Associating function names with JSON metadata</span></span>

<span data-ttu-id="88bb6-237">関数が正しく動作するには、関数の`id`プロパティを JavaScript 実装に関連付ける必要があります。</span><span class="sxs-lookup"><span data-stu-id="88bb6-237">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="88bb6-238">関連付けがあることを確認してください。そうしないと、関数は登録されず、Excel では使用できません。</span><span class="sxs-lookup"><span data-stu-id="88bb6-238">Make sure there is an association, otherwise the function will not be registered and not useable in Excel.</span></span> <span data-ttu-id="88bb6-239">次のコードサンプルは、 `CustomFunctions.associate()`メソッドを使用して関連付けを行う方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="88bb6-239">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="88bb6-240">このサンプルではカスタム関数 `add` を定義し、それを `id` プロパティ値が **ADD** の、JSON メタデータ ファイル内のオブジェクトに関連付けます。</span><span class="sxs-lookup"><span data-stu-id="88bb6-240">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="88bb6-241">次の JSON は、以前のカスタム関数 JavaScript コードに関連付けられている JSON メタデータを示しています。</span><span class="sxs-lookup"><span data-stu-id="88bb6-241">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

```json
{
  "functions": [
    {
      "description": "Add two numbers",
      "id": "ADD",
      "name": "ADD",
      "parameters": [
        {
          "description": "First number",
          "name": "first",
          "type": "number"
        },
        {
          "description": "Second number",
          "name": "second",
          "type": "number"
        }
      ],
      "result": {
        "type": "number"
      }
    }
  ]
}
```

<span data-ttu-id="88bb6-242">JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-242">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="88bb6-243">JSON のメタデータ ファイルにそれぞれの `id` プロパティには、英数字とピリオドのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="88bb6-243">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="88bb6-244">JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-244">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="88bb6-245">すなわち、メタデータ ファイル内の 2 つの関数オブジェクトは同じ `id` 値であってはいけません。</span><span class="sxs-lookup"><span data-stu-id="88bb6-245">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="88bb6-246">対応する JavaScript 関数の名前に関連付けられた後では、JSON のメタデータ ファイル内の `id` プロパティの値を変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-246">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="88bb6-247">JSON のメタデータ ファイル内の `name` プロパティを更新することによって Excel でエンド ユーザーに表示される関数の名前を変更することができます。しかし、確立された後は、 `id` プロパティの値を決して変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-247">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="88bb6-248">JavaScript ファイルで、各関数の`CustomFunctions.associate`後に、カスタム関数の関連付けを指定します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-248">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="88bb6-249">次のサンプルは、JavaScript コード サンプルで定義された関数に対応する JSON メタデータを示します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-249">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="88bb6-250">プロパティ`id`と`name`プロパティの値は、大文字で記述します。これは、カスタム関数を記述するときのベストプラクティスです。</span><span class="sxs-lookup"><span data-stu-id="88bb6-250">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="88bb6-251">この JSON を追加する必要があるのは、自動生成を使用せずに、手動で独自の JSON ファイルを準備する場合だけです。</span><span class="sxs-lookup"><span data-stu-id="88bb6-251">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="88bb6-252">Autogeneration の詳細については、「[カスタム関数の JSON メタデータを作成](custom-functions-json-autogeneration.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="88bb6-252">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="next-steps"></a><span data-ttu-id="88bb6-253">次のステップ</span><span class="sxs-lookup"><span data-stu-id="88bb6-253">Next steps</span></span>

<span data-ttu-id="88bb6-254">[関数に名前を付けるためのベストプラクティス](custom-functions-naming.md)、または前述の手書き JSON メソッドを使用して[関数をローカライズ](custom-functions-localize.md)する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="88bb6-254">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="88bb6-255">関連項目</span><span class="sxs-lookup"><span data-stu-id="88bb6-255">See also</span></span>

- [<span data-ttu-id="88bb6-256">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="88bb6-256">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="88bb6-257">カスタム関数のパラメータオプション</span><span class="sxs-lookup"><span data-stu-id="88bb6-257">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="88bb6-258">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="88bb6-258">Create custom functions in Excel</span></span>](custom-functions-overview.md)
