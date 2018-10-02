---
ms.date: 09/27/2018
description: JavaScript を使用して Excel でカスタム関数を作成します。
title: Excel でカスタム関数を作成する (プレビュー)
ms.openlocfilehash: 98e418f843f6f5574088cea9c7393afc4a42060b
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348802"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="22ee7-103">Excel でカスタム関数を作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="22ee7-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="22ee7-p101">JavaScript で関数をアドインの一部として定義することにより、開発者はカスタム関数を使用して Excel に新しい関数を追加することができます。Excel 内のユーザーは、`SUM()` などの Excel のネイティブ関数にアクセスするのと同様に、カスタム関数にアクセスできます。この記事では、Excel でカスタム関数を作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-p101">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions like any other native function in Excel (such as `SUM()`). This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="22ee7-107">次の図では、エンド ユーザーが Excel ワークシートのセルにカスタム関数を挿入する例を示します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="22ee7-108">`CONTOSO.ADD42` カスタム関数は、ユーザーが関数への入力パラメーターとして指定する数値ペアに、42 を足すように設計されています。</span><span class="sxs-lookup"><span data-stu-id="22ee7-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="22ee7-109">次のコードは、`ADD42` カスタム関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="22ee7-110">この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="22ee7-111">カスタム関数アドイン プロジェクトのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="22ee7-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="22ee7-112">[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Excel カスタム関数アドイン プロジェクトを作成する場合は、ジェネレーターが作成するプロジェクトに以下のようなファイルが表示されます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="22ee7-113">ファイル</span><span class="sxs-lookup"><span data-stu-id="22ee7-113">File</span></span> | <span data-ttu-id="22ee7-114">ファイル形式</span><span class="sxs-lookup"><span data-stu-id="22ee7-114">File format</span></span> | <span data-ttu-id="22ee7-115">説明</span><span class="sxs-lookup"><span data-stu-id="22ee7-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="22ee7-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="22ee7-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="22ee7-117">または</span><span class="sxs-lookup"><span data-stu-id="22ee7-117">or</span></span><br/><span data-ttu-id="22ee7-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="22ee7-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="22ee7-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="22ee7-119">JavaScript</span></span><br/><span data-ttu-id="22ee7-120">または</span><span class="sxs-lookup"><span data-stu-id="22ee7-120">or</span></span><br/><span data-ttu-id="22ee7-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="22ee7-121">TypeScript</span></span> | <span data-ttu-id="22ee7-122">カスタム関数を定義するコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="22ee7-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="22ee7-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="22ee7-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="22ee7-124">JSON</span><span class="sxs-lookup"><span data-stu-id="22ee7-124">JSON</span></span> | <span data-ttu-id="22ee7-125">カスタム関数を説明するメタデータが含まれており、Excel でカスタム関数を登録してエンドユーザーが使用できるようにします。</span><span class="sxs-lookup"><span data-stu-id="22ee7-125">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="22ee7-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="22ee7-126">**./index.html**</span></span> | <span data-ttu-id="22ee7-127">HTML</span><span class="sxs-lookup"><span data-stu-id="22ee7-127">HTML</span></span> | <span data-ttu-id="22ee7-128">カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="22ee7-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="22ee7-129">**Manifest.xml**</span></span> | <span data-ttu-id="22ee7-130">XML</span><span class="sxs-lookup"><span data-stu-id="22ee7-130">XML</span></span> | <span data-ttu-id="22ee7-131">アドイン内のすべてのカスタム関数の名前空間と、このテーブルで前に一覧表示した JavaScript、JSON、HTML ファイルの位置を指定します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="22ee7-132">次のセクションでは、これらのファイルの詳細についてを説明します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-132">The following sections provide more details about these settings.</span></span>

### <a name="script-file"></a><span data-ttu-id="22ee7-133">スクリプト ファイル</span><span class="sxs-lookup"><span data-stu-id="22ee7-133">Script file</span></span> 

<span data-ttu-id="22ee7-134">スクリプト ファイル (Yo Office ジェネレーターが作成するプロジェクト内の **./src/customfunctions.js** または **/src/customfunctions.ts**) には、カスタム関数を定義して、カスタム関数の名前を [JSON メタデータ ファイル](#json-metadata-file)のオブジェクトにマップするコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="22ee7-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="22ee7-135">たとえば、以下のコードでは、`add` と `increment` というカスタム関数を定義して、次に両方の関数のマッピング情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-135">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions.</span></span> <span data-ttu-id="22ee7-136">`add` 関数は、`id` プロパティの値が **ADD** である JSON メタデータ ファイルのオブジェクトにマップされ、`increment` 関数は、`id` プロパティの値が **INCREMENT** であるメタデータ ファイルのオブジェクトにマップされます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-136">The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="22ee7-137">スクリプト ファイルの関数名を JSON メタデータ ファイルのオブジェクトにマップする方法の詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-137">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

```js
function add(first, second){
  return first + second;
}

function increment(incrementBy, callback) {
  var result = 0;
  var timer = setInterval(function() {
    result += incrementBy;
    callback.setResult(result);
  }, 1000);

  callback.onCanceled = function() {
    clearInterval(timer);
  };
}

// map `id` values in the JSON metadata file to the JavaScript function names
CustomFunctionMappings.ADD = add;
CustomFunctionMappings.INCREMENT = increment;
```

### <a name="json-metadata-file"></a><span data-ttu-id="22ee7-138">JSON メタデータ ファイル</span><span class="sxs-lookup"><span data-stu-id="22ee7-138">JSON metadata file</span></span> 

<span data-ttu-id="22ee7-139">カスタム関数のメタデータ ファイル (Yo Office ジェネレーターが作成するプロジェクト内の **./config/customfunctions.json**) は、Excel でカスタム関数を登録してエンドユーザーが使用できるようにするのに必要な情報を示しています。</span><span class="sxs-lookup"><span data-stu-id="22ee7-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="22ee7-140">カスタム関数は、ユーザーがはじめてアドインを実行したときに登録されます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-140">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="22ee7-141">その後、その同じユーザーは、最初にアドインが実行されたブックだけでなく、すべてのブックでそれらのカスタム関数を使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-141">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

> [!TIP]
> <span data-ttu-id="22ee7-142">JSON ファイルをホストするサーバーのサーバー設定では、カスタム関数が Excel Online で正しく作動するために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-142">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="22ee7-143">以下の **customfunctions.json** のコードでは、この記事で前述した `add` 関数と `increment` 関数のメタデータを指定します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-143">The following code in **customfunctions.json** specifies the metadata for the `add` function that was described previously in this article.</span></span> <span data-ttu-id="22ee7-144">このコード サンプルの次の表では、この JSON オブジェクト内の個々のプロパティについての詳細情報を示しています。</span><span class="sxs-lookup"><span data-stu-id="22ee7-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="22ee7-145">JSON メタデータ ファイルの `id` および `name` プロパティの値を指定する方法の詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-145">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com",
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
      "id": "INCREMENT",
      "name": "INCREMENT",
      "description": "Periodically increment a value",
      "helpUrl": "http://www.contoso.com",
      "result": {
          "type": "number",
          "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "increment",
            "description": "Amount to increment",
            "type": "number",
            "dimensionality": "scalar"
        }
    ],
    "options": {
        "cancelable": true,
        "stream": true
      }
    }
  ]
}
```

<span data-ttu-id="22ee7-146">以下の表では、通常 JSON メタデータ ファイルに格納されているプロパティを一覧表示しています。</span><span class="sxs-lookup"><span data-stu-id="22ee7-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="22ee7-147">JSON メタデータ ファイルの詳細については、「[カスタム関数のメタデータ](custom-functions-json.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-147">For more detailed information about the JSON metadata file, including options not used in the previous example, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="22ee7-148">プロパティ</span><span class="sxs-lookup"><span data-stu-id="22ee7-148">Property</span></span>  | <span data-ttu-id="22ee7-149">説明</span><span class="sxs-lookup"><span data-stu-id="22ee7-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="22ee7-150">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="22ee7-150">A unique ID for the group.</span></span> <span data-ttu-id="22ee7-151">設定後は、この ID は変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-151">This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="22ee7-152">Excel でエンド ユーザーに対して表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="22ee7-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="22ee7-153">Excel では、 [XML マニフェスト ファイル](#manifest-file)で指定されているカスタム関数の名前空間が、関数名に接頭辞として付きます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-153">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `helpUrl` | <span data-ttu-id="22ee7-154">ユーザーがヘルプを要求したときに表示されるページの URL です。</span><span class="sxs-lookup"><span data-stu-id="22ee7-154">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="22ee7-155">関数が実行することについて説明します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-155">Describes what the function does.</span></span> <span data-ttu-id="22ee7-156">この値は、関数が Excel 内のオートコンプリート メニューで選択された項目となっている場合に、ツールヒントとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="22ee7-157">関数によって返される情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="22ee7-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="22ee7-158">`type` 子プロパティには、**文字列**、**数値**、または**ブール値**を使用できます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-158">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="22ee7-159">`dimensionality` 子プロパティの値には、**スカラー**または**マトリックス** (指定された `type` の値の 2 次元配列) が使用できます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-159">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `parameters` | <span data-ttu-id="22ee7-160">関数の入力パラメーターを定義する配列。</span><span class="sxs-lookup"><span data-stu-id="22ee7-160">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="22ee7-161">`name` および `description` 子プロパティが Excel intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-161">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="22ee7-162">`type` 子プロパティ値には、**文字列**、**数値**、または**ブール値**を使用できます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-162">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="22ee7-163">`dimensionality` 子プロパティの値には、**スカラー**または **マトリックス** (指定された `type` の値の 2 次元配列) が使用できます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-163">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `options` | <span data-ttu-id="22ee7-164">Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-164">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="22ee7-165">このプロパティの使用方法の詳細については、この記事で後述する「[ストリーム関数](#streamed-functions)」および「[関数のキャンセル](#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-165">For more information about how this property can be used, see [Streamed functions](#streamed-functions) and [Cancellation](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="22ee7-166">マニフェスト ファイル</span><span class="sxs-lookup"><span data-stu-id="22ee7-166">Manifest file</span></span>

<span data-ttu-id="22ee7-167">カスタム関数 (Yo Office ジェネレーターが作成するプロジェクト内の **./manifest.xml**) は、アドイン内のすべてのカスタム関数の名前空間と、JavaScript、JSON、HTML ファイルの位置を指定します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-167">The XML manifest file for an add-in that defines custom functions specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="22ee7-168">以下の XML マークアップでは、カスタム関数を有効にするためにアドインのマニフェストに含める必要のある `<ExtensionPoint>` および `<Resources>` 要素の例を示します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-168">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest in order to enable Excel to run custom functions.</span></span>  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="JS-URL" /> <!--resid points to location of JavaScript file-->
                    </Script>
                    <Page>
                        <SourceLocation resid="HTML-URL"/> <!--resid points to location of HTML file-->
                    </Page>
                    <Metadata>
                        <SourceLocation resid="JSON-URL" /> <!--resid points to location of JSON file-->
                    </Metadata>
                    <Namespace resid="namespace" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="JSON-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.json" /> <!--specifies the location of your JSON file-->
            <bt:Url id="JS-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.js" /> <!--specifies the location of your JavaScript file-->
            <bt:Url id="HTML-URL" DefaultValue="http://127.0.0.1:8080/index.html" /> <!--specifies the location of your HTML file-->
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. -->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="22ee7-169">Excel 内の関数の先頭には、XML マニフェスト ファイルで指定される名前空間が追加されます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-169">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="22ee7-170">関数の名前空間は関数名の前に配置され、それらはピリオドで区切られます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-170">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="22ee7-171">たとえば、Excel ワークシートのセル内の関数 `ADD42` を呼び出すには、`=CONTOSO.ADD42` と入力します。これは、CONTOSO が名前空間で、`ADD42` が JSON ファイルで指定された関数の名前であるからです。</span><span class="sxs-lookup"><span data-stu-id="22ee7-171">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="22ee7-172">名前空間は、所属する会社またはアドインの識別子として使用することを想定しています。</span><span class="sxs-lookup"><span data-stu-id="22ee7-172">The prefix is intended to be used as an identifier for your add-in.</span></span> 

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="22ee7-173">外部ソースからデータを返す関数</span><span class="sxs-lookup"><span data-stu-id="22ee7-173">Functions that return data from external sources</span></span>

<span data-ttu-id="22ee7-174">カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-174">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="22ee7-175">JavaScript Promise を Excel に返す。</span><span class="sxs-lookup"><span data-stu-id="22ee7-175">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="22ee7-176">コールバック関数を使用して Promise を最終値で解決します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-176">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="22ee7-177">カスタム関数は、Excel が最終結果を待つ間、セルに `#GETTING_DATA` の一時的な結果を表示します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-177">Asynchronous functions display a `#GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="22ee7-178">ユーザーは、カスタム関数が結果を待つ間、ワークシートの他の部分を通常通り操作することができます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-178">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

<span data-ttu-id="22ee7-179">以下のコード サンプルでは、`getTemperature()` カスタム関数が温度計の現在の温度を取得します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-179">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="22ee7-180">`sendWebRequest` は [XHR](custom-functions-runtime.md#xhr) を使用して温度 Web サービスを呼び出す仮想関数 (ここでは説明していません) であることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-180">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a><span data-ttu-id="22ee7-181">ストリーム関数</span><span class="sxs-lookup"><span data-stu-id="22ee7-181">Streamed functions</span></span>

<span data-ttu-id="22ee7-182">ストリーム カスタム関数を使用すると、時間の経過とともにセルに繰り返しデータを出力でき、ユーザーが再計算を要求することは特に必要ありません。</span><span class="sxs-lookup"><span data-stu-id="22ee7-182">Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request recalculation.</span></span> <span data-ttu-id="22ee7-183">以下のコード サンプルは、1 秒おきに結果に数値を追加するカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="22ee7-183">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="22ee7-184">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-184">Note the following about this code:</span></span>

- <span data-ttu-id="22ee7-185">Excel は、`setResult`コールバックを使用して自動的に新しい値を表示します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-185">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="22ee7-186">2 番目のパラメーター `handler` は、[オートコンプリート] メニューから関数を選択する場合には、エンドユーザーに対して表示されません。</span><span class="sxs-lookup"><span data-stu-id="22ee7-186">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="22ee7-187">`onCanceled` コールバックは、関数がキャンセルされた場合に実行される関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-187">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="22ee7-188">すべてのストリーム関数に対して、このようなキャンセル ハンドラーを実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-188">You must implement a cancellation handler like this for any streamed function.</span></span> <span data-ttu-id="22ee7-189">詳細については、 「[関数のキャンセル](#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-189">For more information, see [Canceling a function](#canceling-a-function).</span></span> 

```js
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}
```

<span data-ttu-id="22ee7-190">JSON メタデータ ファイルでストリーム関数にメタデータを指定する場合には、以下の例に示すように、`options` オブジェクトにプロパティ `"cancelable": true` および `"stream": true` を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-190">When you specify metadata for a streamed function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

```json
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="22ee7-191">関数のキャンセル</span><span class="sxs-lookup"><span data-stu-id="22ee7-191">Canceling a function</span></span>

<span data-ttu-id="22ee7-192">状況によっては、帯域幅の消費量、作業メモリ、UPC への負荷を減らすために、ストリーム カスタム関数の実行をキャンセルする必要がある場合もあります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-192">In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="22ee7-193">Excel は、以下のような状況では関数の実行をキャンセルします。</span><span class="sxs-lookup"><span data-stu-id="22ee7-193">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="22ee7-194">ユーザーが、関数への参照があるセルを編集または削除した場合。</span><span class="sxs-lookup"><span data-stu-id="22ee7-194">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="22ee7-195">関数の引数 (入力) のいずれかが変更された場合。</span><span class="sxs-lookup"><span data-stu-id="22ee7-195">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="22ee7-196">この場合、キャンセルに続いて新しい関数の呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-196">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="22ee7-197">ユーザーが手動で再計算をトリガーした場合。</span><span class="sxs-lookup"><span data-stu-id="22ee7-197">When the user triggers recalculation manually.</span></span> <span data-ttu-id="22ee7-198">この場合、キャンセルに続いて新しい関数の呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-198">In this case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="22ee7-199">関数をキャンセルする機能を有効にするには、JavaScript 関数内にキャンセル ハンドラーを実装して、関数を記述する JSON メタデータの `options` オブジェクト内にプロパティ `"cancelable": true` を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-199">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="22ee7-200">この記事の前のセクションのコード サンプルは、これらの手法の例を示しています。</span><span class="sxs-lookup"><span data-stu-id="22ee7-200">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="22ee7-201">状態の保存と共有</span><span class="sxs-lookup"><span data-stu-id="22ee7-201">Saving and sharing state</span></span>

<span data-ttu-id="22ee7-202">カスタム関数では、JavaScript のグローバル変数にデータを保存できます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-202">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="22ee7-203">後続の呼び出しでは、カスタム関数はこれらの変数に保存されている値を使用できます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-203">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="22ee7-204">保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を追加する場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="22ee7-204">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="22ee7-205">たとえば、Web リソースへの呼び出しから返されたデータを保存しておけば、同じ Web リソースへ繰り返し呼び出しを行わなくて済みます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-205">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="22ee7-206">以下のコード サンプルは、 状態をグローバルで保存する温度ストリーミング関数の実装を示しています。</span><span class="sxs-lookup"><span data-stu-id="22ee7-206">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="22ee7-207">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-207">Note the following about this code:</span></span>

- <span data-ttu-id="22ee7-208">`refreshTemperature` は、1 秒おきに特定の温度計の温度を読み取るストリーム関数です。</span><span class="sxs-lookup"><span data-stu-id="22ee7-208">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="22ee7-209">新しい温度は `savedTemperatures` 変数に保存されますが、セルの値を直接更新することはありません。</span><span class="sxs-lookup"><span data-stu-id="22ee7-209">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="22ee7-210">ワークシート・セルから直接呼び出されません。\*したがって、JSON ファイルには登録されません \*</span><span class="sxs-lookup"><span data-stu-id="22ee7-210">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>

- <span data-ttu-id="22ee7-211">`streamTemperature` 1 秒おきにセルに表示される温度値を更新します。また、 `savedTemperatures` 変数をデータ ソースとして使用します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-211">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="22ee7-212">JSON ファイルに登録し、すべて大文字で `STREAMTEMPERATURE` という名前をつける必要があります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-212">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>

- <span data-ttu-id="22ee7-213">ユーザーは、Excel UI の複数のセルから `streamTemperature` を呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-213">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="22ee7-214">呼び出すたびに、同じ `savedTemperatures` 変数からデータを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-214">Each call reads data from the same `savedTemperatures` variable.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
  }
  getNextTemperature();
}

function refreshTemperature(thermometerID){
  sendWebRequest(thermometerID, function(data){
    savedTemperatures[thermometerID] = data.temperature;
  });
  setTimeout(function(){
    refreshTemperature(thermometerID);
  }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="22ee7-215">データの範囲を使用する</span><span class="sxs-lookup"><span data-stu-id="22ee7-215">Working with ranges of data</span></span>

<span data-ttu-id="22ee7-216">カスタム関数は、入力パラメーターとしてデータの範囲を受け取ることができます。または、データの範囲を返すこともできます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-216">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="22ee7-217">JavaScript では、データの範囲は、2 次元配列として表されます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-217">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="22ee7-218">たとえば、関数が Excel に格納されている数値の範囲から 2 番目に高い値を返すとします。</span><span class="sxs-lookup"><span data-stu-id="22ee7-218">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="22ee7-219">以下の関数が、タイプ `Excel.CustomFunctionDimensionality.matrix` のものである `values` パラメーターを受け取ります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-219">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="22ee7-220">この関数の JSON メタデータでは、パラメーターの `type` プロパティを `matrix` に設定するように注意してください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-220">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 1; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="handling-errors"></a><span data-ttu-id="22ee7-221">エラーの処理</span><span class="sxs-lookup"><span data-stu-id="22ee7-221">Handling errors</span></span>

<span data-ttu-id="22ee7-222">カスタム関数を定義するアドインをビルドする場合には、実行時エラーに対処するエラー処理ロジックを含めるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-222">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="22ee7-223">カスタム関数のエラー処理は、[一般的な Excel JavaScript API のエラー処理](excel-add-ins-error-handling.md) と同じです。</span><span class="sxs-lookup"><span data-stu-id="22ee7-223">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="22ee7-224">以下のコード サンプルでは、`.catch` がコード内で発生するエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-224">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="known-issues"></a><span data-ttu-id="22ee7-225">既知の問題</span><span class="sxs-lookup"><span data-stu-id="22ee7-225">Known issues</span></span>

- <span data-ttu-id="22ee7-226">ヘルプの URL とパラメーターの説明。Excel ではまだ使用されていません。</span><span class="sxs-lookup"><span data-stu-id="22ee7-226">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="22ee7-227">カスタム関数は現在、モバイル クライアント用の Excel では使用できません。</span><span class="sxs-lookup"><span data-stu-id="22ee7-227">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="22ee7-228">揮発性関数（スプレッドシート内の無関係なデータが変更されたときに自動的に再計算する関数）はまだサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="22ee7-228">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="22ee7-229">Office 365 管理ポータルと AppSource による展開はまだ有効になっていません。</span><span class="sxs-lookup"><span data-stu-id="22ee7-229">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="22ee7-230">Excel Online のカスタム関数は、一定期間使用しないとセッション中に機能しなくなることがあります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-230">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="22ee7-231">ブラウザページを更新（F5）し、カスタム関数を再入力して機能を復元します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-231">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="22ee7-232">Excel for Windows で実行されている複数のアドインがある場合には、ワークシートのセル内に **#GETTING_DATA** の一時的な結果が表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-232">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="22ee7-233">すべての Excel ウィンドウを閉じて、Excel を再起動します。</span><span class="sxs-lookup"><span data-stu-id="22ee7-233">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="22ee7-234">将来的には、カスタム関数用のデバッグ ツールが利用可能となる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="22ee7-234">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="22ee7-235">それまでは、F12 開発者ツールを使用して Excel オンラインでデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="22ee7-235">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="22ee7-236">詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-236">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="22ee7-237">変更ログ</span><span class="sxs-lookup"><span data-stu-id="22ee7-237">Changelog</span></span>

- <span data-ttu-id="22ee7-238">**2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開\*</span><span class="sxs-lookup"><span data-stu-id="22ee7-238">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="22ee7-239">**2017 年 11 月 20 日**: ビルド 8801 以降を使用しているユーザー向けに互換性バグを修正</span><span class="sxs-lookup"><span data-stu-id="22ee7-239">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="22ee7-240">**2017 年 11 月 28 日**: 非同期関数のキャンセルへのサポートを公開\* (ストリーム関数への変更が必要)</span><span class="sxs-lookup"><span data-stu-id="22ee7-240">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="22ee7-241">**2018 年 5 月 7 日**: Mac、Excel Online、およびインプロセスで実行される同期関数へのサポートを公開\*</span><span class="sxs-lookup"><span data-stu-id="22ee7-241">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="22ee7-242">**2018 年 9 月 20日**: JavaScript 実行時のカスタム関数へのサポートを公開</span><span class="sxs-lookup"><span data-stu-id="22ee7-242">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="22ee7-243">詳細については、「[Excel カスタム関数のランタイム](custom-functions-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="22ee7-243">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="22ee7-244">\* Office Insiders チャネル対象</span><span class="sxs-lookup"><span data-stu-id="22ee7-244">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="22ee7-245">関連項目</span><span class="sxs-lookup"><span data-stu-id="22ee7-245">See also</span></span>

* [<span data-ttu-id="22ee7-246">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="22ee7-246">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="22ee7-247">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="22ee7-247">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="22ee7-248">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="22ee7-248">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="22ee7-249">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="22ee7-249">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)