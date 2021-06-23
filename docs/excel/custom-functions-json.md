---
ms.date: 12/22/2020
description: カスタム関数の JSON メタデータを定義し、Excel ID と名前のプロパティを関連付ける。
title: カスタム関数の JSON メタデータを手動で作成Excel
localization_priority: Normal
ms.openlocfilehash: 514eacba5045d160eb6f3d4823adbd8c2f45292a
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075902"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a><span data-ttu-id="35e6e-103">カスタム関数の JSON メタデータを手動で作成する</span><span class="sxs-lookup"><span data-stu-id="35e6e-103">Manually create JSON metadata for custom functions</span></span>

<span data-ttu-id="35e6e-104">カスタム関数の概要[](custom-functions-overview.md)に関する記事で説明したように、カスタム関数プロジェクトには、JSON メタデータ ファイルとスクリプト (JavaScript または TypeScript) ファイルの両方を含め、関数を登録して使用できる必要があります。</span><span class="sxs-lookup"><span data-stu-id="35e6e-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="35e6e-105">カスタム関数は、ユーザーが初めてアドインを実行した後、すべてのブックで同じユーザーが使用できる場合に登録されます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="35e6e-106">独自の JSON ファイルを作成する代わりに、可能な場合は JSON 自動生成を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="35e6e-106">We recommend using JSON autogeneration when possible instead of creating your own JSON file.</span></span> <span data-ttu-id="35e6e-107">自動生成はユーザー エラーの発生が少なく、スキャフォールディング `yo office` されたファイルには既にこれが含まれます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-107">Autogeneration is less prone to user error and the `yo office` scaffolded files already include this.</span></span> <span data-ttu-id="35e6e-108">JSDoc タグと JSON 自動生成プロセスの詳細については、「カスタム関数の JSON メタデータの自動生成 [」を参照してください](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="35e6e-108">For more information on JSDoc tags and the JSON autogeneration process, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="35e6e-109">ただし、カスタム関数プロジェクトを最初から作成できます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-109">However, you can make a custom functions project from scratch.</span></span> <span data-ttu-id="35e6e-110">このプロセスでは、次の操作を行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="35e6e-110">This process requires you to:</span></span>

- <span data-ttu-id="35e6e-111">JSON ファイルを書き込む。</span><span class="sxs-lookup"><span data-stu-id="35e6e-111">Write your JSON file.</span></span>
- <span data-ttu-id="35e6e-112">マニフェスト ファイルが JSON ファイルに接続されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-112">Check that your manifest file is connected to your JSON file.</span></span>
- <span data-ttu-id="35e6e-113">関数を登録するために `id` 、 `name` スクリプト ファイル内の関数とプロパティを関連付ける。</span><span class="sxs-lookup"><span data-stu-id="35e6e-113">Associate your functions' `id` and `name` properties in the script file in order to register your functions.</span></span>

<span data-ttu-id="35e6e-114">次の図は、スキャフォールディング ファイルの使用と JSON の最初 `yo office` からの書き込みとの違いを説明しています。</span><span class="sxs-lookup"><span data-stu-id="35e6e-114">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>

![Yo の使用と独自の JSON のOfficeの違いのイメージ。](../images/custom-functions-json.png)

> [!NOTE]
> <span data-ttu-id="35e6e-116">ジェネレーターを使用しない場合は、XML マニフェスト ファイルのセクションを使用して、作成した JSON ファイルにマニフェストを `<Resources>` 接続してください `yo office` 。</span><span class="sxs-lookup"><span data-stu-id="35e6e-116">Remember to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file if you do not use the `yo office` generator.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="35e6e-117">メタデータの作成とマニフェストへの接続</span><span class="sxs-lookup"><span data-stu-id="35e6e-117">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="35e6e-118">プロジェクトに JSON ファイルを作成し、関数のパラメーターなど、プロジェクト内の関数に関する詳細を提供します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-118">Create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="35e6e-119">関数プロパティ[の完全な一覧](#json-metadata-example)[については、](#metadata-reference)次のメタデータ例とメタデータ参照を参照してください。</span><span class="sxs-lookup"><span data-stu-id="35e6e-119">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="35e6e-120">次の例と同様に、XML マニフェスト ファイルでセクション内の JSON `<Resources>` ファイルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="35e6e-120">Ensure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

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

## <a name="json-metadata-example"></a><span data-ttu-id="35e6e-121">JSON メタデータの例</span><span class="sxs-lookup"><span data-stu-id="35e6e-121">JSON metadata example</span></span>

<span data-ttu-id="35e6e-122">次の例では、カスタム関数を定義するアドインの JSON メタデータ ファイルの内容を示します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-122">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="35e6e-123">この例の後に続くセクションでは、JSON の例に含まれる個々のプロパティの詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-123">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="35e6e-124">完全なサンプル JSON ファイルは、リポジトリのコミット履歴Excel [OfficeDev/GitHub-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json)で使用できます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-124">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="35e6e-125">プロジェクトが JSON を自動的に生成するために調整されたので、手書き JSON の完全なサンプルは、以前のバージョンのプロジェクトでのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-125">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="35e6e-126">メタデータ参照</span><span class="sxs-lookup"><span data-stu-id="35e6e-126">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="35e6e-127">functions</span><span class="sxs-lookup"><span data-stu-id="35e6e-127">functions</span></span>

<span data-ttu-id="35e6e-128">`functions` プロパティは、カスタム関数オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-128">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="35e6e-129">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-129">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="35e6e-130">プロパティ</span><span class="sxs-lookup"><span data-stu-id="35e6e-130">Property</span></span>      | <span data-ttu-id="35e6e-131">データ型</span><span class="sxs-lookup"><span data-stu-id="35e6e-131">Data type</span></span> | <span data-ttu-id="35e6e-132">必須</span><span class="sxs-lookup"><span data-stu-id="35e6e-132">Required</span></span> | <span data-ttu-id="35e6e-133">説明</span><span class="sxs-lookup"><span data-stu-id="35e6e-133">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="35e6e-134">string</span><span class="sxs-lookup"><span data-stu-id="35e6e-134">string</span></span>    | <span data-ttu-id="35e6e-135">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-135">No</span></span>       | <span data-ttu-id="35e6e-136">Excel でエンド ユーザーに表示される関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-136">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="35e6e-137">たとえば、「**華氏の値を摂氏に変換する**」です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-137">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="35e6e-138">文字列</span><span class="sxs-lookup"><span data-stu-id="35e6e-138">string</span></span>    | <span data-ttu-id="35e6e-139">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-139">No</span></span>       | <span data-ttu-id="35e6e-140">関数に関する情報を提供する URL です </span><span class="sxs-lookup"><span data-stu-id="35e6e-140">URL that provides information about the function.</span></span> <span data-ttu-id="35e6e-141">(作業ウィンドウに表示されます)。たとえば、`http://contoso.com/help/convertcelsiustofahrenheit.html` です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-141">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="35e6e-142">文字列</span><span class="sxs-lookup"><span data-stu-id="35e6e-142">string</span></span>    | <span data-ttu-id="35e6e-143">はい</span><span class="sxs-lookup"><span data-stu-id="35e6e-143">Yes</span></span>      | <span data-ttu-id="35e6e-144">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-144">A unique ID for the function.</span></span> <span data-ttu-id="35e6e-145">この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。</span><span class="sxs-lookup"><span data-stu-id="35e6e-145">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="35e6e-146">文字列</span><span class="sxs-lookup"><span data-stu-id="35e6e-146">string</span></span>    | <span data-ttu-id="35e6e-147">はい</span><span class="sxs-lookup"><span data-stu-id="35e6e-147">Yes</span></span>      | <span data-ttu-id="35e6e-148">Excel でエンド ユーザーに表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-148">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="35e6e-149">このExcel、この関数名には、XML マニフェスト ファイルで指定されたカスタム関数名前空間のプレフィックスが付けされます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-149">In Excel, this function name is prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="35e6e-150">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="35e6e-150">object</span></span>    | <span data-ttu-id="35e6e-151">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-151">No</span></span>       | <span data-ttu-id="35e6e-152">Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-152">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="35e6e-153">詳細については、[options](#options) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="35e6e-153">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="35e6e-154">配列</span><span class="sxs-lookup"><span data-stu-id="35e6e-154">array</span></span>     | <span data-ttu-id="35e6e-155">はい</span><span class="sxs-lookup"><span data-stu-id="35e6e-155">Yes</span></span>      | <span data-ttu-id="35e6e-156">関数の入力パラメーターを定義する配列です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-156">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="35e6e-157">詳細については [、パラメーター](#parameters) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="35e6e-157">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="35e6e-158">object</span><span class="sxs-lookup"><span data-stu-id="35e6e-158">object</span></span>    | <span data-ttu-id="35e6e-159">はい</span><span class="sxs-lookup"><span data-stu-id="35e6e-159">Yes</span></span>      | <span data-ttu-id="35e6e-160">関数が返す情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="35e6e-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="35e6e-161">詳細については、[result](#result) に関する説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="35e6e-161">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="35e6e-162">options</span><span class="sxs-lookup"><span data-stu-id="35e6e-162">options</span></span>

<span data-ttu-id="35e6e-163">`options` オブジェクトでは、Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-163">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="35e6e-164">次の表に、`options` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-164">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="35e6e-165">プロパティ</span><span class="sxs-lookup"><span data-stu-id="35e6e-165">Property</span></span>          | <span data-ttu-id="35e6e-166">データ型</span><span class="sxs-lookup"><span data-stu-id="35e6e-166">Data type</span></span> | <span data-ttu-id="35e6e-167">必須</span><span class="sxs-lookup"><span data-stu-id="35e6e-167">Required</span></span>                               | <span data-ttu-id="35e6e-168">説明</span><span class="sxs-lookup"><span data-stu-id="35e6e-168">Description</span></span> |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | <span data-ttu-id="35e6e-169">ブール</span><span class="sxs-lookup"><span data-stu-id="35e6e-169">boolean</span></span>   | <span data-ttu-id="35e6e-170">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-170">No</span></span><br/><br/><span data-ttu-id="35e6e-171">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-171">Default value is `false`.</span></span>  | <span data-ttu-id="35e6e-172">`true` の場合、手動での再計算のトリガーや、関数によって参照されているセルの編集など、関数をキャンセルする効果のある操作をユーザーが実行すると、Excel によって `CancelableInvocation` ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-172">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="35e6e-173">キャンセル可能な関数は、通常、1 つの結果を返す非同期関数でのみ使用され、データ要求の取り消しを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="35e6e-173">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="35e6e-174">関数は、プロパティとプロパティの両方 `stream` を使用 `cancelable` することはできません。</span><span class="sxs-lookup"><span data-stu-id="35e6e-174">A function can't use both the `stream` and `cancelable` properties.</span></span> |
| `requiresAddress` | <span data-ttu-id="35e6e-175">ブール</span><span class="sxs-lookup"><span data-stu-id="35e6e-175">boolean</span></span>   | <span data-ttu-id="35e6e-176">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-176">No</span></span> <br/><br/><span data-ttu-id="35e6e-177">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-177">Default value is `false`.</span></span> | <span data-ttu-id="35e6e-178">場合 `true` は、カスタム関数は、それを呼び出したセルのアドレスにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-178">If `true`, your custom function can access the address of the cell that invoked it.</span></span> <span data-ttu-id="35e6e-179">呼 `address` び出しパラメーター [のプロパティには](custom-functions-parameter-options.md#invocation-parameter) 、カスタム関数を呼び出したセルのアドレスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-179">The `address` property of the [invocation parameter](custom-functions-parameter-options.md#invocation-parameter) contains the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="35e6e-180">関数は、プロパティとプロパティの両方 `stream` を使用 `requiresAddress` することはできません。</span><span class="sxs-lookup"><span data-stu-id="35e6e-180">A function can't use both the `stream` and `requiresAddress` properties.</span></span> |
| `requiresParameterAddresses` | <span data-ttu-id="35e6e-181">ブール</span><span class="sxs-lookup"><span data-stu-id="35e6e-181">boolean</span></span>   | <span data-ttu-id="35e6e-182">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-182">No</span></span> <br/><br/><span data-ttu-id="35e6e-183">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-183">Default value is `false`.</span></span> | <span data-ttu-id="35e6e-184">場合 `true` は、カスタム関数は、関数の入力パラメーターのアドレスにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-184">If `true`, your custom function can access the addresses of the function's input parameters.</span></span> <span data-ttu-id="35e6e-185">このプロパティは、result オブジェクトのプロパティと `dimensionality` 組み[](#result)合わせて使用する必要があります。 `dimensionality` `matrix`</span><span class="sxs-lookup"><span data-stu-id="35e6e-185">This property must be used in combination with the `dimensionality` property of the [result](#result) object, and `dimensionality` must be set to `matrix`.</span></span> <span data-ttu-id="35e6e-186">詳細 [については、「パラメーターのアドレスを検出する](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="35e6e-186">See [Detect the address of a parameter](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) for more information.</span></span> |
| `stream`          | <span data-ttu-id="35e6e-187">ブール</span><span class="sxs-lookup"><span data-stu-id="35e6e-187">boolean</span></span>   | <span data-ttu-id="35e6e-188">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-188">No</span></span><br/><br/><span data-ttu-id="35e6e-189">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-189">Default value is `false`.</span></span>  | <span data-ttu-id="35e6e-190">`true` の場合、1 回のみ呼び出されたときにも、関数はセルに繰り返し出力できます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-190">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="35e6e-191">このオプションは、株価などの急速に変化するデータ ソースに便利です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-191">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="35e6e-192">この関数には、`return` ステートメントは含めないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="35e6e-192">The function should have no `return` statement.</span></span> <span data-ttu-id="35e6e-193">代わりに、結果の値は `StreamingInvocation.setResult` コールバック メソッドの引数として渡されます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-193">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="35e6e-194">詳細については、「ストリーミング関数 [を作成する」を参照してください](custom-functions-web-reqs.md#make-a-streaming-function)。</span><span class="sxs-lookup"><span data-stu-id="35e6e-194">For more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `volatile`        | <span data-ttu-id="35e6e-195">ブール</span><span class="sxs-lookup"><span data-stu-id="35e6e-195">boolean</span></span>   | <span data-ttu-id="35e6e-196">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-196">No</span></span> <br/><br/><span data-ttu-id="35e6e-197">既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-197">Default value is `false`.</span></span> | <span data-ttu-id="35e6e-198">場合は、数式の依存値が変更Excelではなく、計算が再計算されるごとに関数が `true` 再計算されます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-198">If `true`, the function recalculates each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="35e6e-199">関数は、プロパティとプロパティの両方 `stream` を使用 `volatile` することはできません。</span><span class="sxs-lookup"><span data-stu-id="35e6e-199">A function can't use both the `stream` and `volatile` properties.</span></span> <span data-ttu-id="35e6e-200">プロパティと `stream` プロパティ `volatile` の両方がに設定されている場合 `true` 、揮発性プロパティは無視されます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-200">If the `stream` and `volatile` properties are both set to `true`, the volatile property will be ignored.</span></span> |

### <a name="parameters"></a><span data-ttu-id="35e6e-201">parameters</span><span class="sxs-lookup"><span data-stu-id="35e6e-201">parameters</span></span>

<span data-ttu-id="35e6e-202">`parameters` プロパティは、パラメーター オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-202">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="35e6e-203">次の表に、各オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-203">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="35e6e-204">プロパティ</span><span class="sxs-lookup"><span data-stu-id="35e6e-204">Property</span></span>  |  <span data-ttu-id="35e6e-205">データ型</span><span class="sxs-lookup"><span data-stu-id="35e6e-205">Data type</span></span>  |  <span data-ttu-id="35e6e-206">必須</span><span class="sxs-lookup"><span data-stu-id="35e6e-206">Required</span></span>  |  <span data-ttu-id="35e6e-207">説明</span><span class="sxs-lookup"><span data-stu-id="35e6e-207">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="35e6e-208">string</span><span class="sxs-lookup"><span data-stu-id="35e6e-208">string</span></span>  |  <span data-ttu-id="35e6e-209">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-209">No</span></span> |  <span data-ttu-id="35e6e-210">パラメーターの説明です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-210">A description of the parameter.</span></span> <span data-ttu-id="35e6e-211">これは、ユーザーのExcelに表示IntelliSense。</span><span class="sxs-lookup"><span data-stu-id="35e6e-211">This is displayed in Excel's IntelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="35e6e-212">文字列</span><span class="sxs-lookup"><span data-stu-id="35e6e-212">string</span></span>  |  <span data-ttu-id="35e6e-213">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-213">No</span></span>  |  <span data-ttu-id="35e6e-214">(配列 `scalar` 以外の値) または `matrix` (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="35e6e-214">Must be either `scalar` (a non-array value) or `matrix` (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="35e6e-215">文字列</span><span class="sxs-lookup"><span data-stu-id="35e6e-215">string</span></span>  |  <span data-ttu-id="35e6e-216">はい</span><span class="sxs-lookup"><span data-stu-id="35e6e-216">Yes</span></span>  |  <span data-ttu-id="35e6e-217">パラメーターの名前です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-217">The name of the parameter.</span></span> <span data-ttu-id="35e6e-218">この名前は、ExcelのIntelliSense。</span><span class="sxs-lookup"><span data-stu-id="35e6e-218">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="35e6e-219">文字列</span><span class="sxs-lookup"><span data-stu-id="35e6e-219">string</span></span>  |  <span data-ttu-id="35e6e-220">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-220">No</span></span>  |  <span data-ttu-id="35e6e-221">パラメーターのデータ型です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-221">The data type of the parameter.</span></span> <span data-ttu-id="35e6e-222">、、、、を使用すると、前の 3 つの種類 `boolean` `number` `string` `any` の任意のを使用できます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-222">Can be `boolean`, `number`, `string`, or `any`, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="35e6e-223">このプロパティを指定しない場合、データ型の既定値は `any` .</span><span class="sxs-lookup"><span data-stu-id="35e6e-223">If this property is not specified, the data type defaults to `any`.</span></span> |
|  `optional`  | <span data-ttu-id="35e6e-224">ブール</span><span class="sxs-lookup"><span data-stu-id="35e6e-224">boolean</span></span> | <span data-ttu-id="35e6e-225">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-225">No</span></span> | <span data-ttu-id="35e6e-226">`true` の場合、パラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="35e6e-226">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="35e6e-227">ブール</span><span class="sxs-lookup"><span data-stu-id="35e6e-227">boolean</span></span> | <span data-ttu-id="35e6e-228">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-228">No</span></span> | <span data-ttu-id="35e6e-229">If , parameters populate from a specified `true` array.</span><span class="sxs-lookup"><span data-stu-id="35e6e-229">If `true`, parameters populate from a specified array.</span></span> <span data-ttu-id="35e6e-230">関数のすべての繰り返しパラメーターは、定義によって省略可能なパラメーターと見なされます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-230">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="35e6e-231">result</span><span class="sxs-lookup"><span data-stu-id="35e6e-231">result</span></span>

<span data-ttu-id="35e6e-232">`result` オブジェクトは、この関数が返す情報の種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-232">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="35e6e-233">次の表に、`result` オブジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-233">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="35e6e-234">プロパティ</span><span class="sxs-lookup"><span data-stu-id="35e6e-234">Property</span></span>         | <span data-ttu-id="35e6e-235">データ型</span><span class="sxs-lookup"><span data-stu-id="35e6e-235">Data type</span></span> | <span data-ttu-id="35e6e-236">必須</span><span class="sxs-lookup"><span data-stu-id="35e6e-236">Required</span></span> | <span data-ttu-id="35e6e-237">説明</span><span class="sxs-lookup"><span data-stu-id="35e6e-237">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="35e6e-238">string</span><span class="sxs-lookup"><span data-stu-id="35e6e-238">string</span></span>    | <span data-ttu-id="35e6e-239">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-239">No</span></span>       | <span data-ttu-id="35e6e-240">(配列 `scalar` 以外の値) または `matrix` (2 次元配列) のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="35e6e-240">Must be either `scalar` (a non-array value) or `matrix` (a 2-dimensional array).</span></span> |
| `type` | <span data-ttu-id="35e6e-241">文字列</span><span class="sxs-lookup"><span data-stu-id="35e6e-241">string</span></span>    | <span data-ttu-id="35e6e-242">いいえ</span><span class="sxs-lookup"><span data-stu-id="35e6e-242">No</span></span>       | <span data-ttu-id="35e6e-243">結果のデータ型。</span><span class="sxs-lookup"><span data-stu-id="35e6e-243">The data type of the result.</span></span> <span data-ttu-id="35e6e-244">、、、または (これは、前の 3 つの種類の任意の `boolean` `number` `string` `any` 使用を可能にする) を指定できます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-244">Can be `boolean`, `number`, `string`, or `any` (which allows you to use of any of the previous three types).</span></span> <span data-ttu-id="35e6e-245">このプロパティを指定しない場合、データ型の既定値は `any` .</span><span class="sxs-lookup"><span data-stu-id="35e6e-245">If this property is not specified, the data type defaults to `any`.</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="35e6e-246">関数名を JSON メタデータに関連付ける</span><span class="sxs-lookup"><span data-stu-id="35e6e-246">Associating function names with JSON metadata</span></span>

<span data-ttu-id="35e6e-247">関数が正しく動作するには、関数のプロパティを `id` JavaScript 実装に関連付ける必要があります。</span><span class="sxs-lookup"><span data-stu-id="35e6e-247">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="35e6e-248">関連付けがある場合、それ以外の場合は関数は登録されないので、この関数で使用Excel。</span><span class="sxs-lookup"><span data-stu-id="35e6e-248">Make sure there is an association, otherwise the function won't be registered and isn't useable in Excel.</span></span> <span data-ttu-id="35e6e-249">次のコード サンプルは、メソッドを使用して関連付けを行う方法を示 `CustomFunctions.associate()` しています。</span><span class="sxs-lookup"><span data-stu-id="35e6e-249">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="35e6e-250">このサンプルではカスタム関数 `add` を定義し、それを `id` プロパティ値が **ADD** の、JSON メタデータ ファイル内のオブジェクトに関連付けます。</span><span class="sxs-lookup"><span data-stu-id="35e6e-250">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

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

<span data-ttu-id="35e6e-251">次の JSON は、前のカスタム関数 JavaScript コードに関連付けられている JSON メタデータを示しています。</span><span class="sxs-lookup"><span data-stu-id="35e6e-251">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

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

<span data-ttu-id="35e6e-252">JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。</span><span class="sxs-lookup"><span data-stu-id="35e6e-252">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="35e6e-253">JSON のメタデータ ファイルにそれぞれの `id` プロパティには、英数字とピリオドのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="35e6e-253">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="35e6e-254">JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-254">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="35e6e-255">すなわち、メタデータ ファイル内の 2 つの関数オブジェクトは同じ `id` 値であってはいけません。</span><span class="sxs-lookup"><span data-stu-id="35e6e-255">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="35e6e-256">対応する JavaScript 関数の名前に関連付けられた後では、JSON のメタデータ ファイル内の `id` プロパティの値を変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="35e6e-256">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="35e6e-257">JSON のメタデータ ファイル内の `name` プロパティを更新することによって Excel でエンド ユーザーに表示される関数の名前を変更することができます。しかし、確立された後は、 `id` プロパティの値を決して変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="35e6e-257">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="35e6e-258">JavaScript ファイルで、各関数の後に使用するカスタム関数 `CustomFunctions.associate` の関連付けを指定します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-258">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="35e6e-259">次のサンプルは、前の JavaScript コード サンプルで定義されている関数に対応する JSON メタデータを示しています。</span><span class="sxs-lookup"><span data-stu-id="35e6e-259">The following sample shows the JSON metadata that corresponds to the functions defined in the preceding JavaScript code sample.</span></span> <span data-ttu-id="35e6e-260">プロパティ `id` の値は大文字で、カスタム関数を記述する場合の `name` ベスト プラクティスです。</span><span class="sxs-lookup"><span data-stu-id="35e6e-260">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="35e6e-261">独自の JSON ファイルを手動で準備し、自動生成を使用しない場合にのみ、この JSON を追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="35e6e-261">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="35e6e-262">自動生成の詳細については、「カスタム関数の [JSON メタデータの自動生成」を参照してください](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="35e6e-262">For more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="35e6e-263">次の手順</span><span class="sxs-lookup"><span data-stu-id="35e6e-263">Next steps</span></span>

<span data-ttu-id="35e6e-264">関数に[名前を付けるベスト](custom-functions-naming.md)プラクティスを説明するか[](custom-functions-localize.md)、前に説明した手書き JSON メソッドを使用して関数をローカライズする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="35e6e-264">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="35e6e-265">関連項目</span><span class="sxs-lookup"><span data-stu-id="35e6e-265">See also</span></span>

- [<span data-ttu-id="35e6e-266">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="35e6e-266">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="35e6e-267">カスタム関数パラメーター のオプション</span><span class="sxs-lookup"><span data-stu-id="35e6e-267">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="35e6e-268">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="35e6e-268">Create custom functions in Excel</span></span>](custom-functions-overview.md)
