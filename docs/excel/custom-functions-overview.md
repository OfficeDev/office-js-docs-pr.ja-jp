---
ms.date: 12/14/2018
description: JavaScript を使用して Excel でカスタム関数を作成する。
title: Excel でのカスタム関数の作成 (プレビュー)
ms.openlocfilehash: 87f56f4c697d19296fe1b539e4071c8e79fbed6a
ms.sourcegitcommit: 09f124fac7b2e711e1a8be562a99624627c0699e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/15/2018
ms.locfileid: "27283117"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="57188-103">Excel でのカスタム関数の作成 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="57188-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="57188-104">開発者は、カスタム関数を使用して関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="57188-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="57188-105">ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="57188-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="57188-106">この記事では、Excel でカスタム関数を作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="57188-106">This article explains how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="57188-107">次の図は、エンドユーザーが Excel ワークシートのセルにカスタム関数を挿入する様子を示します。</span><span class="sxs-lookup"><span data-stu-id="57188-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="57188-108">`CONTOSO.ADD42` カスタム関数は、関数への入力パラメーターとしてユーザーが指定した数値のペアに 42 を追加するように設計されています。</span><span class="sxs-lookup"><span data-stu-id="57188-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="57188-109">`ADD42` カスタム関数は次のコードにより定義されます。</span><span class="sxs-lookup"><span data-stu-id="57188-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="57188-110">この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。</span><span class="sxs-lookup"><span data-stu-id="57188-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="57188-111">カスタム関数 アドイン プロジェクトのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="57188-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="57188-112">[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Excel のカスタム関数アドイン プロジェクトを作成する場合、ジェネレーターが作成するプロジェクトに以下のようなファイルが表示されます。</span><span class="sxs-lookup"><span data-stu-id="57188-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="57188-113">ファイル</span><span class="sxs-lookup"><span data-stu-id="57188-113">File</span></span> | <span data-ttu-id="57188-114">ファイル形式</span><span class="sxs-lookup"><span data-stu-id="57188-114">File format</span></span> | <span data-ttu-id="57188-115">説明</span><span class="sxs-lookup"><span data-stu-id="57188-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="57188-116">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="57188-116">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="57188-117">または</span><span class="sxs-lookup"><span data-stu-id="57188-117">or</span></span><br/><span data-ttu-id="57188-118">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="57188-118">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="57188-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="57188-119">JavaScript</span></span><br/><span data-ttu-id="57188-120">または</span><span class="sxs-lookup"><span data-stu-id="57188-120">or</span></span><br/><span data-ttu-id="57188-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="57188-121">TypeScript</span></span> | <span data-ttu-id="57188-122">カスタム関数を定義するコードが含みます。</span><span class="sxs-lookup"><span data-stu-id="57188-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="57188-123">**./src/functions/functions.json**</span><span class="sxs-lookup"><span data-stu-id="57188-123">**./src/functions/functions.json**</span></span> | <span data-ttu-id="57188-124">JSON</span><span class="sxs-lookup"><span data-stu-id="57188-124">JSON</span></span> | <span data-ttu-id="57188-125">カスタム関数を定義し、Excel に関数を登録してエンドユーザーが使用できるようにするためのメタデータを含みます。</span><span class="sxs-lookup"><span data-stu-id="57188-125">Contains metadata that describes custom functions and enables Excel to register the custom functions and make them available to end users.</span></span> |
| <span data-ttu-id="57188-126">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="57188-126">**./src/functions/functions.html**</span></span> | <span data-ttu-id="57188-127">HTML</span><span class="sxs-lookup"><span data-stu-id="57188-127">HTML</span></span> | <span data-ttu-id="57188-128">カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。</span><span class="sxs-lookup"><span data-stu-id="57188-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="57188-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="57188-129">**Manifest.xml**</span></span> | <span data-ttu-id="57188-130">XML</span><span class="sxs-lookup"><span data-stu-id="57188-130">XML</span></span> | <span data-ttu-id="57188-131">アドイン内のすべてのカスタム関数の名前空間と、この表で前述した JavaScript、JSON、HTML ファイルの位置を指定します。</span><span class="sxs-lookup"><span data-stu-id="57188-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="57188-132">次のセクションでは、これらのファイルに関する詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="57188-132">The following sections provide more information about these files.</span></span>

### <a name="script-file"></a><span data-ttu-id="57188-133">スクリプト ファイル</span><span class="sxs-lookup"><span data-stu-id="57188-133">Script file</span></span>

<span data-ttu-id="57188-134">スクリプト ファイル (Yo Office ジェネレーターが作成するプロジェクト内の **./src/customfunctions.js** または **/src/customfunctions.ts**) には、カスタム関数を定義して、カスタム関数の名前を [JSON メタデータ ファイル](#json-metadata-file)のオブジェクトにマップするコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="57188-134">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="57188-135">例えば、次のコードはカスタム関数 `add` と `increment` を定義し、両方の関数のマッピング情報を指定します。 </span><span class="sxs-lookup"><span data-stu-id="57188-135">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span> <span data-ttu-id="57188-136">`add` 関数は、`id` プロパティの値が **ADD** の JSON メタデータ ファイル内のオブジェクトにマップされ、`increment` 関数は、`id` プロパティの値が **INCREMENT** のメタデータ ファイル内のオブジェクトにマップされます。</span><span class="sxs-lookup"><span data-stu-id="57188-136">The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="57188-137">JSON メタデータ ファイル内のオブジェクトへのスクリプト ファイル内関数名のマッピングの詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57188-137">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

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

### <a name="json-metadata-file"></a><span data-ttu-id="57188-138">JSON メタデータ ファイル</span><span class="sxs-lookup"><span data-stu-id="57188-138">JSON metadata file</span></span> 

<span data-ttu-id="57188-139">カスタム関数のメタデータ ファイル (Yo Office ジェネレーターが作成するプロジェクトでは **./config/customfunctions.json**) は、Excel がカスタム関数の登録し、エンドユーザーが利用できるようするために必要な情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="57188-139">The custom functions metadata file (**./src/functions/functions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="57188-140">カスタム関数は、ユーザーがアドインを初めて実行するときに登録されます。</span><span class="sxs-lookup"><span data-stu-id="57188-140">Custom functions are registered when a user runs an add-in for the first time.</span></span> <span data-ttu-id="57188-141">その後は、同じユーザーに対しては、(アドインが最初に実行されたワークブック内のみでなく) すべてのワークブック内で利用が可能になります。</span><span class="sxs-lookup"><span data-stu-id="57188-141">After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="57188-142">JSON ファイルをホストするサーバーでは、カスタム関数を Excel Online で正しく作動させるために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="57188-142">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="57188-143">**functions.json** の次のコードは、`add` 関数のメタデータと上述の `increment` 関数を指定します。</span><span class="sxs-lookup"><span data-stu-id="57188-143">The following code in **functions.json** specifies the metadata for the `add` function and the `increment` function that were described previously.</span></span> <span data-ttu-id="57188-144">このコード サンプルに続く表では、JSON オブジェクト内の個別のプロパティについての詳細情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="57188-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="57188-145">JSON メタデータ ファイル内の `id` と `name` 各プロパティーの値の指定に関する詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57188-145">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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
        "stream": true,
        "volatile": false
      }
    }
  ]
}
```

<span data-ttu-id="57188-146">次の表は、JSON メタデータ ファイルに通常格納されているプロパティの一覧表示です。</span><span class="sxs-lookup"><span data-stu-id="57188-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="57188-147">JSON メタデータ ファイルの詳細については、「[カスタム関数のメタデータ](custom-functions-json.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57188-147">For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="57188-148">プロパティ</span><span class="sxs-lookup"><span data-stu-id="57188-148">Property</span></span>  | <span data-ttu-id="57188-149">説明</span><span class="sxs-lookup"><span data-stu-id="57188-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="57188-150">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="57188-150">A unique ID for the function.</span></span> <span data-ttu-id="57188-151">この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。</span><span class="sxs-lookup"><span data-stu-id="57188-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="57188-152">Excel でエンド ユーザーに表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="57188-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="57188-153">Excel では、この関数名は [XML マニフェスト ファイル](#manifest-file)で指定されているカスタム関数の名前空間でプレフィックスされます。</span><span class="sxs-lookup"><span data-stu-id="57188-153">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `helpUrl` | <span data-ttu-id="57188-154">ユーザーがヘルプを要求したときに表示されるページの URL です。</span><span class="sxs-lookup"><span data-stu-id="57188-154">URL for the page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="57188-155">関数について説明します。</span><span class="sxs-lookup"><span data-stu-id="57188-155">Describes what the function does.</span></span> <span data-ttu-id="57188-156">この値は、関数が Excel 内のオートコンプリート メニューで選択された項目となっている場合に、ツールヒントとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="57188-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="57188-157">関数が返す情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="57188-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="57188-158">このオブジェクトに関する詳細情報については [result](custom-functions-json.md#result) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57188-158">For detailed information about this object, see [result](custom-functions-json.md#result).</span></span> |
| `parameters` | <span data-ttu-id="57188-159">関数の入力パラメーターを定義する配列です。</span><span class="sxs-lookup"><span data-stu-id="57188-159">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="57188-160">このオブジェクトに関する詳細情報については [parameters](custom-functions-json.md#parameters) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57188-160">For detailed information about this object, see [parameters](custom-functions-json.md#parameters).</span></span> |
| `options` | <span data-ttu-id="57188-161">Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="57188-161">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="57188-162">このプロパティの使用方法の詳細については、この記事で後述する[ストリーム関数](#streaming-functions)、[関数のキャンセル](#canceling-a-function)、および[揮発性関数の宣言](#declaring-a-volatile-function)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57188-162">For more information about how this property can be used, see [Streaming functions](#streaming-functions), [Canceling a function](#canceling-a-function), and [Declaring a volatile function](#declaring-a-volatile-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="57188-163">マニフェスト ファイル</span><span class="sxs-lookup"><span data-stu-id="57188-163">Manifest file</span></span>

<span data-ttu-id="57188-164">カスタム関数 (Yo Office ジェネレーターが作成するプロジェクトでは **./manifest.xml**) を定義するアドインの XML マニフェスト ファイルは、アドイン内のすべてのカスタム関数の名前空間と、 JavaScript、JSON、および HTML の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="57188-164">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="57188-165">次の XML マークアップでは、`<ExtensionPoint>` と `<Resources>` カスタム関数を有効にするアドインのマニフェストに含める必要がある要素の例を示します。</span><span class="sxs-lookup"><span data-stu-id="57188-165">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
        <Hosts>
            <Host xsi:type="Workbook">
                <AllFormFactors>
                    <ExtensionPoint xsi:type="CustomFunctions">
                        <Script>
                            <SourceLocation resid="Contoso.Functions.Script.Url" />
                        </Script>
                        <Page>
                            <SourceLocation resid="Contoso.Functions.Page.Url"/>
                        </Page>
                        <Metadata>
                            <SourceLocation resid="Contoso.Functions.Metadata.Url" />
                        </Metadata>
                        <Namespace resid="Contoso.Functions.Namespace" />
                    </ExtensionPoint>
                </AllFormFactors>
            </Host>
        </Hosts>
        <Resources>
            <bt:Images>
                <bt:Image id="Contoso.tpicon_16x16" DefaultValue="https://localhost:3000/assets/icon-16.png" />
                <bt:Image id="Contoso.tpicon_32x32" DefaultValue="https://localhost:3000/assets/icon-32.png" />
                <bt:Image id="Contoso.tpicon_80x80" DefaultValue="https://localhost:3000/assets/icon-80.png" />
            </bt:Images>
            <bt:Urls>
                <bt:Url id="Contoso.Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js" />
                <bt:Url id="Contoso.Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json" />
                <bt:Url id="Contoso.Functions.Page.Url" DefaultValue="https://localhost:3000/dist/functions.html" />
                <bt:Url id="Contoso.Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
            </bt:Urls>
            <bt:ShortStrings>
                <bt:String id="Contoso.Functions.Namespace" DefaultValue="CONTOSO" />
            </bt:ShortStrings>
        </Resources>
    </VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="57188-166">Excel の関数は、XML マニフェスト ファイルで指定された名前空間が接頭辞として付加されます。</span><span class="sxs-lookup"><span data-stu-id="57188-166">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="57188-167">関数の名前空間は、関数名の前に付けられ、ピリオドで区切られます。</span><span class="sxs-lookup"><span data-stu-id="57188-167">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="57188-168">例えば、Excel ワークシートのセル内で、`ADD42` 関数を呼び出すためには、`=CONTOSO.ADD42` と入力します。これは、`CONTOSO` が名前空間で、`ADD42` が JSON ファイルで指定された関数の名前だからです。</span><span class="sxs-lookup"><span data-stu-id="57188-168">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="57188-169">名前空間は、会社またはアドインの識別子としての使用を目的としています。</span><span class="sxs-lookup"><span data-stu-id="57188-169">The prefix is intended to be used as an identifier for your add-in.</span></span> <span data-ttu-id="57188-170">名前空間にはアルファベットとピリオドのみを含めることが出来ます。</span><span class="sxs-lookup"><span data-stu-id="57188-170">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="57188-171">外部ソースからデータを返す関数</span><span class="sxs-lookup"><span data-stu-id="57188-171">Functions that return data from external sources</span></span>

<span data-ttu-id="57188-172">カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="57188-172">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="57188-173">JavaScript Promise を Excel に返します。</span><span class="sxs-lookup"><span data-stu-id="57188-173">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="57188-174">コールバック関数を使用して Promise を最終値で解決します。</span><span class="sxs-lookup"><span data-stu-id="57188-174">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="57188-175">カスタム関数は、Excel での最終結果を待つ間、`#GETTING_DATA` という一時的な結果をセルに表示します。</span><span class="sxs-lookup"><span data-stu-id="57188-175">Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="57188-176">ユーザーは、結果を待つ間もワークシートの残りの部分を通常通り操作することができます。</span><span class="sxs-lookup"><span data-stu-id="57188-176">Users can interact normally with the rest of the worksheet while they wait for the result.</span></span>

<span data-ttu-id="57188-177">次のコード例では、`getTemperature()` カスタム関数が温度計の現在の温度を取得します。 </span><span class="sxs-lookup"><span data-stu-id="57188-177">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="57188-178">`sendWebRequest` は、[XHR](custom-functions-runtime.md#xhr-example) を使用して温度 Web サービスを呼び出す仮想の関数 (ここでは指定なし) であることに留意してください。</span><span class="sxs-lookup"><span data-stu-id="57188-178">Note that `sendWebRequest` is a hypothetical function (not specified here) that uses [XHR](custom-functions-runtime.md#xhr-example) to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a><span data-ttu-id="57188-179">ストリーミング関数</span><span class="sxs-lookup"><span data-stu-id="57188-179">Streaming functions</span></span>

<span data-ttu-id="57188-180">ストリーム カスタム関数を使用すると、セルに繰り返しデータを長期的に出力でき、ユーザーが再計算を明示的に要求することは特に必要ありません。</span><span class="sxs-lookup"><span data-stu-id="57188-180">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="57188-181">以下のコード サンプルは、毎秒ごとに結果に数値を追加するカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="57188-181">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="57188-182">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="57188-182">Note the following about this code:</span></span>

- <span data-ttu-id="57188-183">Excel は、`setResult` コールバックを使用して自動的に新しい値を表示します。</span><span class="sxs-lookup"><span data-stu-id="57188-183">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="57188-184">2 番目の入力パラメーターの `handler` は、[オートコンプリート] メニューから関数が選択された場合、Excel のエンドユーザーに表示されません。</span><span class="sxs-lookup"><span data-stu-id="57188-184">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="57188-185">`onCanceled` コールバックは、関数がキャンセルされた場合に実行される関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="57188-185">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="57188-186">すべてのストリーム関数には、このようなキャンセル ハンドラーの実装が必要です。</span><span class="sxs-lookup"><span data-stu-id="57188-186">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="57188-187">詳細については、「[関数をキャンセルする](#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57188-187">For more information, see [Canceling a function](#canceling-a-function).</span></span>

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

<span data-ttu-id="57188-188">JSON メタデータ ファイルでストリーミング関数にメタデータを指定する場合には、`options` オブジェクト内のプロパティ`"cancelable": true` および `"stream": true` を以下の例のように設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="57188-188">When you specify metadata for a streaming function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

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

## <a name="canceling-a-function"></a><span data-ttu-id="57188-189">関数をキャンセルする</span><span class="sxs-lookup"><span data-stu-id="57188-189">For more information, see Canceling a function.</span></span>

<span data-ttu-id="57188-190">状況によっては、帯域幅の消費量、作業メモリ、UPC への負荷を軽減するために、ストリーム カスタム関数の実行をキャンセルする必要があります。</span><span class="sxs-lookup"><span data-stu-id="57188-190">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="57188-191">Excel では、次のような状況で関数の実行をキャンセルします。</span><span class="sxs-lookup"><span data-stu-id="57188-191">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="57188-192">ユーザーが、関数を参照するセルを編集または削除した場合。</span><span class="sxs-lookup"><span data-stu-id="57188-192">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="57188-193">関数の引数 (入力) の 1 つが変更されたとき。</span><span class="sxs-lookup"><span data-stu-id="57188-193">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="57188-194">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="57188-194">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="57188-195">ユーザーが手動で再計算をトリガーしたとき。</span><span class="sxs-lookup"><span data-stu-id="57188-195">When the user triggers recalculation manually.</span></span> <span data-ttu-id="57188-196">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="57188-196">In this case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="57188-197">関数をキャンセルする機能を有効にするには、JavaScript 関数内にキャンセル ハンドラーを実装し、関数を記述するJSONのメタデータの `options` オブジェクト内のプロパティ `"cancelable": true` を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="57188-197">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="57188-198">この記事の前のセクションのコード サンプルに、これらの手法の例が示されています。</span><span class="sxs-lookup"><span data-stu-id="57188-198">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="declaring-a-volatile-function"></a><span data-ttu-id="57188-199">揮発性関数の宣言</span><span class="sxs-lookup"><span data-stu-id="57188-199">Declaring a volatile function</span></span>

<span data-ttu-id="57188-200">[揮発性関数](https://docs.microsoft.com/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)とは、関数のいずれの引数にも変更がない場合でも、値が刻々と変化する関数のことです。</span><span class="sxs-lookup"><span data-stu-id="57188-200">[Volatile functions](https://docs.microsoft.com/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed.</span></span> <span data-ttu-id="57188-201">これらの関数は、Excel が再計算するたびに再計算を行います。</span><span class="sxs-lookup"><span data-stu-id="57188-201">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="57188-202">たとえば、`NOW` 関数を呼び出すセルがあるとします。</span><span class="sxs-lookup"><span data-stu-id="57188-202">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="57188-203">`NOW` が呼び出される度に、現在の日付と時刻を自動的に返します。</span><span class="sxs-lookup"><span data-stu-id="57188-203">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="57188-204">Excel には、`RAND` や `TODAY` などの組み込み揮発性関数がいくつか含まれています。</span><span class="sxs-lookup"><span data-stu-id="57188-204">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="57188-205">Excel のすべての揮発性関数の一覧は、「[揮発性および非揮発性関数](https://docs.microsoft.com/ja-JP/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="57188-205">For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](https://docs.microsoft.com/ja-JP/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>  
  
<span data-ttu-id="57188-206">カスタム関数を使用すると独自の揮発性関数を作成することができ、日時、時間、乱数、およびモデルを処理するときに役立つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="57188-206">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling.</span></span> <span data-ttu-id="57188-207">たとえば、モンテカルロ シミュレーションでは、最適なソリューションを決定するにはランダムな入力値の生成が必要です。</span><span class="sxs-lookup"><span data-stu-id="57188-207">For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.</span></span>  
  
<span data-ttu-id="57188-208">関数を揮発性であると宣言するには、次のコードで示されるように、JSON メタデータファイルの関数で、`options` オブジェクトに`"volatile": true` を追加します。</span><span class="sxs-lookup"><span data-stu-id="57188-208">To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample.</span></span> <span data-ttu-id="57188-209">関数で `"streaming": true`と`"volatile": true` の両方をマークすることはできません。両方とも `true` とマークされている場合、揮発性のオプションは無視されます。</span><span class="sxs-lookup"><span data-stu-id="57188-209">Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.</span></span>  

```json
{
  "name": "TOMORROW",
  "description":  "Returns tomorrow’s date",
  "helpUrl": "http://www.contoso.com",
  "result": {
      "type": "string",
      "dimensionality": "scalar"
  },
  "options": {
      "volatile": true
  }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="57188-210">状態の保存と共有</span><span class="sxs-lookup"><span data-stu-id="57188-210">Saving and sharing state</span></span>

<span data-ttu-id="57188-211">カスタム関数は、グローバル JavaScript 変数にデータを保存でき、以降の呼び出しで使用することができます。</span><span class="sxs-lookup"><span data-stu-id="57188-211">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="57188-212">保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を呼び出す場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="57188-212">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="57188-213">たとえば、同じ Web リソースへの追加呼び出しを避けるために、呼び出しから返されたデータを Web リソースに保存することができます。</span><span class="sxs-lookup"><span data-stu-id="57188-213">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="57188-214">次のコード サンプルでは、状態をグローバルに保存する温度ストリーミング関数の実装を示します。</span><span class="sxs-lookup"><span data-stu-id="57188-214">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="57188-215">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="57188-215">Note the following about this code:</span></span>

- <span data-ttu-id="57188-216">`streamTemperature` 関数がセルに表示される温度の値を毎秒更新し、`savedTemperatures` 変数をデータ ソースとして使用します。</span><span class="sxs-lookup"><span data-stu-id="57188-216">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="57188-217">`streamTemperature` はストリーム関数であるため、その関数がキャンセルされたときに実行されるキャンセル ハンドラーを実装します。</span><span class="sxs-lookup"><span data-stu-id="57188-217">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="57188-218">ユーザーが `streamTemperature` 関数を Excel の複数のセルから呼び出す場合、`streamTemperature` 関数は実行のたびに、同じ `savedTemperatures` 変数からのデータを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="57188-218">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="57188-219">`refreshTemperature` 関数は、特定の温度計の温度を毎秒読み取り、結果を `savedTemperatures` 変数に格納します。</span><span class="sxs-lookup"><span data-stu-id="57188-219">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="57188-220">`refreshTemperature` 関数は、Excel でエンド ユーザーには公開されないので、JSON ファイルに登録する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="57188-220">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="57188-221">データの範囲を使用する</span><span class="sxs-lookup"><span data-stu-id="57188-221">Working with ranges of data</span></span>

<span data-ttu-id="57188-222">カスタム関数は、データの範囲を入力パラメーターとして受け入れることができ、また、データの範囲を返すこともできます。</span><span class="sxs-lookup"><span data-stu-id="57188-222">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="57188-223">JavaScript では、データの範囲は 2 次元配列として表されます。</span><span class="sxs-lookup"><span data-stu-id="57188-223">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="57188-224">例えば、関数が Excel に保存されている数値の範囲から 2 番目に大きい値を返すとします。</span><span class="sxs-lookup"><span data-stu-id="57188-224">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="57188-225">次の関数は、`Excel.CustomFunctionDimensionality.matrix` 型の `values` パラメーターを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="57188-225">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="57188-226">なお、この関数の JSON メタデータでは、パラメーターの`type`プロパティを`matrix` と設定します。</span><span class="sxs-lookup"><span data-stu-id="57188-226">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="discovering-cells-that-invoke-custom-functions"></a><span data-ttu-id="57188-227">カスタム関数を呼び出すセルを検出する</span><span class="sxs-lookup"><span data-stu-id="57188-227">Discovering cells that invoke custom functions</span></span>

<span data-ttu-id="57188-228">カスタム関数を使用すると、範囲の書式設定、キャッシュされた値の表示、およびを `caller.address` を使用しての値の調整を行うこともでき、カスタム関数を呼び出すセルを検出することができます。</span><span class="sxs-lookup"><span data-stu-id="57188-228">Custom funtions also allows you to format ranges, display cached values, and reconcile values using `caller.address`, which makes it possible to discover the cell that invoked a custom function.</span></span> <span data-ttu-id="57188-229">次のシナリオの一部で `caller.address` を使用します。</span><span class="sxs-lookup"><span data-stu-id="57188-229">You might use `caller.address` in some of the following scenarios:</span></span>

- <span data-ttu-id="57188-230">範囲の書式設定: [AsyncStorage](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data)で情報を格納するセルのキーとして `caller.address` を使用します。</span><span class="sxs-lookup"><span data-stu-id="57188-230">Formatting ranges: Use `caller.address` as the key of the cell to store information in [AsyncStorage](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="57188-231">Excel で [onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet#oncalculated) を使用して`AsyncStorage` からキーを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="57188-231">Then, use [onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="57188-232">キャッシュされた値を表示させる: 関数がオフラインで使用される場合、`onCalculated` を使用して `AsyncStorage` に格納されているキャッシュされた値を表示します。</span><span class="sxs-lookup"><span data-stu-id="57188-232">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="57188-233">調整: `caller.address` を使用して元のセルを検出し、処理が発生している場所での調整を行えます。</span><span class="sxs-lookup"><span data-stu-id="57188-233">Reconciliation: Use `caller.address` to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="57188-234">セルのアドレスに関する情報は、関数の JSON メタデータ ファイルで `requiresAddress` が`true` とマークされている場合にのみ公開されます。</span><span class="sxs-lookup"><span data-stu-id="57188-234">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="57188-235">これの例を次のサンプルに示します。</span><span class="sxs-lookup"><span data-stu-id="57188-235">The following sample gives an example of this:</span></span>

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

<span data-ttu-id="57188-236">セルのアドレスを検索するために、スクリプト ファイル (**./src/customfunctions.js**または **./src/customfunctions.ts**) に `getAddress` 関数を追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="57188-236">In the script file (**./src/customfunctions.js** or **./src/customfunctions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="57188-237">この関数は、次のサンプルで示される `parameter1` のようなパラメーターを受け取ることができます。</span><span class="sxs-lookup"><span data-stu-id="57188-237">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="57188-238">最後のパラメーターは常に `invocationContext` で、これはJSON メタデータ ファイルで `requiresAddress` が `true` とマークされているときに Excel が返すセルの位置が格納されているオブジェクトのことです。</span><span class="sxs-lookup"><span data-stu-id="57188-238">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="57188-239">既定では、`getAddress` 関数が返す値は次の形式に従います: `SheetName!CellNumber`。</span><span class="sxs-lookup"><span data-stu-id="57188-239">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="57188-240">たとえば、ある関数が Expenses という名前のシートのセル B2 から呼び出される場合の戻り値は `Expenses!B2` になります。</span><span class="sxs-lookup"><span data-stu-id="57188-240">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="handling-errors"></a><span data-ttu-id="57188-241">エラーの処理</span><span class="sxs-lookup"><span data-stu-id="57188-241">Handling errors</span></span>

<span data-ttu-id="57188-242">カスタム関数を定義するアドインをビルドする場合は、実行時エラーを考慮して、エラー処理ロジックを含めるようにします。</span><span class="sxs-lookup"><span data-stu-id="57188-242">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="57188-243">カスタム関数のエラー処理は、[全体的な Excel の JavaScript API のエラー処理](excel-add-ins-error-handling.md)と同じです。</span><span class="sxs-lookup"><span data-stu-id="57188-243">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="57188-244">次のコード サンプルでは、`.catch` がコード内で以前に発生したエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="57188-244">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="57188-245">既知の問題</span><span class="sxs-lookup"><span data-stu-id="57188-245">Known issues</span></span>

- <span data-ttu-id="57188-246">ヘルプの URL とパラメーターの説明。Excel ではまだ使用されていません。</span><span class="sxs-lookup"><span data-stu-id="57188-246">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="57188-247">カスタム関数は現在、モバイル クライアント用の Excel では使用できません。</span><span class="sxs-lookup"><span data-stu-id="57188-247">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="57188-248">揮発性関数 (スプレッドシート内の無関係なデータが変更されたときに自動的に再計算する関数) はまだサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57188-248">Volatile functions (those that recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="57188-249">Office 365 管理ポータルと AppSource による展開は、まだ有効になっていません。</span><span class="sxs-lookup"><span data-stu-id="57188-249">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="57188-250">Excel Onlineでのカスタム関数は、一定期間動作していないと、セッション中に停止することがあります。</span><span class="sxs-lookup"><span data-stu-id="57188-250">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="57188-251">ブラウザーのページを更新 (F5) し、機能を復元するカスタム関数を再入力します。</span><span class="sxs-lookup"><span data-stu-id="57188-251">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="57188-252">Windows 版 Excel で複数のアドインが実行されている場合、ワークシートのセル内に **#GETTING_DATA** という一時的な結果が表示されることがあります。</span><span class="sxs-lookup"><span data-stu-id="57188-252">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="57188-253">その場合には、Excel のウィンドウをすべて閉じ、Excel を再起動します。</span><span class="sxs-lookup"><span data-stu-id="57188-253">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="57188-254">今後、カスタム関数向けのデバッグ ツールが利用できるようになる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="57188-254">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="57188-255">それまでは、F12 開発者ツールを使用して Excel Online をデバッグすることができます。</span><span class="sxs-lookup"><span data-stu-id="57188-255">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="57188-256">詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57188-256">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="57188-257">変更ログ</span><span class="sxs-lookup"><span data-stu-id="57188-257">Changelog</span></span>

- <span data-ttu-id="57188-258">**2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開\*</span><span class="sxs-lookup"><span data-stu-id="57188-258">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="57188-259">**2017 年 11 月 20 日**: ビルド 8801 以降を使用する場合の互換性バグを修正</span><span class="sxs-lookup"><span data-stu-id="57188-259">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="57188-260">**2017 年 11 月 28 日**: 非同期関数のキャンセルのサポートを公開\* (ストリーミング機能の変更が必要)</span><span class="sxs-lookup"><span data-stu-id="57188-260">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="57188-261">**2018 年 5 月 7 日**: Mac、Excel Online、およびインプロセスで実行される同期関数へのサポートを公開\*</span><span class="sxs-lookup"><span data-stu-id="57188-261">**May 7, 2018**: Shipped\* support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="57188-262">**2018 年 9 月 20日**: JavaScript ランタイムのカスタム関数へのサポートを公開。</span><span class="sxs-lookup"><span data-stu-id="57188-262">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="57188-263">詳細については、「[Excel カスタム関数のランタイム](custom-functions-runtime.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="57188-263">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>
- <span data-ttu-id="57188-264">**2018 年 10 月 20 日**: [10 月の Insider ビルド](https://support.office.com/ja-JP/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24)では、カスタム関数は、 Windows デスクトップ用およびオンライン用の[カスタム定義メタデータ](custom-functions-json.md)で 'id' パラメーターが必要になりました。</span><span class="sxs-lookup"><span data-stu-id="57188-264">**October 20, 2018**: With the [October Insiders build](https://support.office.com/ja-JP/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), Custom Functions now requires the 'id' parameter in your [custom functions metadata](custom-functions-json.md) for Windows Desktop and Online.</span></span> <span data-ttu-id="57188-265">Mac では、このパラメーターは無視します。</span><span class="sxs-lookup"><span data-stu-id="57188-265">On Mac, this parameter should be ignored.</span></span>


<span data-ttu-id="57188-266">\* は、[Office Insider](https://products.office.com/office-insider) チャンネル (旧称 "Insider Fast") </span><span class="sxs-lookup"><span data-stu-id="57188-266">\* to the [Office Insider](https://products.office.com/office-insider) channel (formerly called "Insider Fast")</span></span>

## <a name="see-also"></a><span data-ttu-id="57188-267">関連項目</span><span class="sxs-lookup"><span data-stu-id="57188-267">See also</span></span>

* [<span data-ttu-id="57188-268">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="57188-268">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="57188-269">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="57188-269">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="57188-270">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="57188-270">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="57188-271">チュートリアル: Excel でカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="57188-271">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
