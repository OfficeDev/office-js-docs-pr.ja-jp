---
ms.date: 03/19/2019
description: JavaScript を使用して Excel でカスタム関数を作成する。
title: Excel でのカスタム関数の作成 (プレビュー)
localization_priority: Priority
ms.openlocfilehash: ac3410267da415c4d567092da2e653fcffd10b72
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870451"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="256cd-103">Excel でのカスタム関数の作成 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="256cd-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="256cd-104">開発者は、カスタム関数を使用して関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="256cd-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="256cd-105">ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="256cd-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="256cd-106">この記事では、Excel でカスタム関数を作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="256cd-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="256cd-107">次の図は、エンドユーザーが Excel ワークシートのセルにカスタム関数を挿入する様子を示します。</span><span class="sxs-lookup"><span data-stu-id="256cd-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="256cd-108">`CONTOSO.ADD42` カスタム関数は、関数への入力パラメーターとしてユーザーが指定した数値のペアに 42 を追加するように設計されています。</span><span class="sxs-lookup"><span data-stu-id="256cd-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="256cd-109">`ADD42` カスタム関数は次のコードにより定義されます。</span><span class="sxs-lookup"><span data-stu-id="256cd-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="256cd-110">この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。</span><span class="sxs-lookup"><span data-stu-id="256cd-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="256cd-111">カスタム関数 アドイン プロジェクトのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="256cd-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="256cd-112">[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Excel のカスタム関数アドイン プロジェクトを作成する場合、ジェネレーターが作成するプロジェクトに以下のようなファイルが表示されます。</span><span class="sxs-lookup"><span data-stu-id="256cd-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="256cd-113">ファイル</span><span class="sxs-lookup"><span data-stu-id="256cd-113">File</span></span> | <span data-ttu-id="256cd-114">ファイル形式</span><span class="sxs-lookup"><span data-stu-id="256cd-114">File format</span></span> | <span data-ttu-id="256cd-115">説明</span><span class="sxs-lookup"><span data-stu-id="256cd-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="256cd-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="256cd-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="256cd-117">または</span><span class="sxs-lookup"><span data-stu-id="256cd-117">or</span></span><br/><span data-ttu-id="256cd-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="256cd-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="256cd-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="256cd-119">JavaScript</span></span><br/><span data-ttu-id="256cd-120">または</span><span class="sxs-lookup"><span data-stu-id="256cd-120">or</span></span><br/><span data-ttu-id="256cd-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="256cd-121">TypeScript</span></span> | <span data-ttu-id="256cd-122">カスタム関数を定義するコードが含みます。</span><span class="sxs-lookup"><span data-stu-id="256cd-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="256cd-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="256cd-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="256cd-124">JSON</span><span class="sxs-lookup"><span data-stu-id="256cd-124">JSON</span></span> | <span data-ttu-id="256cd-125">カスタム関数を定義し、Excel に関数を登録してエンドユーザーが使用できるようにするためのメタデータを含みます。</span><span class="sxs-lookup"><span data-stu-id="256cd-125">Contains metadata that describes custom functions and enables Excel to register the custom functions and make them available to end users.</span></span> |
| <span data-ttu-id="256cd-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="256cd-126">**./index.html**</span></span> | <span data-ttu-id="256cd-127">HTML</span><span class="sxs-lookup"><span data-stu-id="256cd-127">HTML</span></span> | <span data-ttu-id="256cd-128">カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。</span><span class="sxs-lookup"><span data-stu-id="256cd-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="256cd-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="256cd-129">**./manifest.xml**</span></span> | <span data-ttu-id="256cd-130">XML</span><span class="sxs-lookup"><span data-stu-id="256cd-130">XML</span></span> | <span data-ttu-id="256cd-131">アドイン内のすべてのカスタム関数の名前空間と、この表で前述した JavaScript、JSON、HTML ファイルの位置を指定します。</span><span class="sxs-lookup"><span data-stu-id="256cd-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="256cd-132">次のセクションでは、これらのファイルに関する詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="256cd-132">The following sections provide more information about these files.</span></span>

### <a name="script-file"></a><span data-ttu-id="256cd-133">スクリプト ファイル</span><span class="sxs-lookup"><span data-stu-id="256cd-133">Script file</span></span>

<span data-ttu-id="256cd-134">スクリプト ファイル (Yo Office ジェネレーターが作成するプロジェクト内の **./src/customfunctions.js** または **./src/customfunctions.ts**) には、カスタム関数を定義して、カスタム関数の名前を [JSON メタデータ ファイル](#json-metadata-file)のオブジェクトにマップするコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="256cd-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="256cd-135">たとえば、次のコードはカスタム関数 `add` と `increment` を定義し、両方の関数の関連付け情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="256cd-135">For example, the following code defines the custom functions `add` and `increment` and then specifies association information for both functions.</span></span> <span data-ttu-id="256cd-136">`add` 関数は、`id` プロパティの値が **ADD** の JSON メタデータ ファイル内のオブジェクトに関連付けられ、`increment` 関数は、`id` プロパティの値が **INCREMENT** のメタデータ ファイル内のオブジェクトに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="256cd-136">The `add` function is associated with the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is associated with the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="256cd-137">JSON メタデータ ファイル内のオブジェクトへのスクリプト ファイル内関数名の関連付けの詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#associating-function-names-with-json-metadata)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="256cd-137">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about associating function names in the script file to objects in the JSON metadata file.</span></span>

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

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
 CustomFunctions.associate("INCREMENT", increment);
```

### <a name="json-metadata-file"></a><span data-ttu-id="256cd-138">JSON メタデータ ファイル</span><span class="sxs-lookup"><span data-stu-id="256cd-138">JSON metadata file</span></span>

<span data-ttu-id="256cd-139">カスタム関数のメタデータ ファイル (Yo Office ジェネレーターが作成するプロジェクトでは **./config/customfunctions.json**) は、Excel がカスタム関数の登録し、エンドユーザーが利用できるようするために必要な情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="256cd-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="256cd-140">カスタム関数は、ユーザーがアドインを初めて実行するときに登録されます。</span><span class="sxs-lookup"><span data-stu-id="256cd-140">Custom functions are registered when a user runs an add-in for the first time.</span></span> <span data-ttu-id="256cd-141">その後は、同じユーザーに対しては、(アドインが最初に実行されたワークブック内のみでなく) すべてのワークブック内で利用が可能になります。</span><span class="sxs-lookup"><span data-stu-id="256cd-141">After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="256cd-142">JSON ファイルをホストするサーバーでは、カスタム関数を Excel Online で正しく作動させるために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="256cd-142">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="256cd-143">**customfunctions.json** の次のコードは、`add` 関数のメタデータと上述の `increment` 関数を指定します。</span><span class="sxs-lookup"><span data-stu-id="256cd-143">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously.</span></span> <span data-ttu-id="256cd-144">このコード サンプルに続く表では、JSON オブジェクト内の個別のプロパティについての詳細情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="256cd-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="256cd-145">JSON メタデータ ファイル内の `id` と `name` 各プロパティーの値の指定に関する詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#associating-function-names-with-json-metadata)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="256cd-145">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="256cd-146">次の表は、JSON メタデータ ファイルに通常格納されているプロパティの一覧表示です。</span><span class="sxs-lookup"><span data-stu-id="256cd-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="256cd-147">JSON メタデータ ファイルの詳細については、「[カスタム関数のメタデータ](custom-functions-json.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="256cd-147">For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="256cd-148">プロパティ</span><span class="sxs-lookup"><span data-stu-id="256cd-148">Property</span></span>  | <span data-ttu-id="256cd-149">説明</span><span class="sxs-lookup"><span data-stu-id="256cd-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="256cd-150">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="256cd-150">A unique ID for the function.</span></span> <span data-ttu-id="256cd-151">この ID には、英数字とピリオドしか使用できません。また、設定後に変更してはいけません。</span><span class="sxs-lookup"><span data-stu-id="256cd-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="256cd-152">Excel でエンド ユーザーに表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="256cd-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="256cd-153">Excel では、この関数名は [XML マニフェスト ファイル](#manifest-file)で指定されているカスタム関数の名前空間でプレフィックスされます。</span><span class="sxs-lookup"><span data-stu-id="256cd-153">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="256cd-154">ユーザーがヘルプを要求したときに表示されるページの URL です。</span><span class="sxs-lookup"><span data-stu-id="256cd-154">URL for the page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="256cd-155">関数について説明します。</span><span class="sxs-lookup"><span data-stu-id="256cd-155">Describes what the function does.</span></span> <span data-ttu-id="256cd-156">この値は、関数が Excel 内のオートコンプリート メニューで選択された項目となっている場合に、ツールヒントとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="256cd-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="256cd-157">関数が返す情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="256cd-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="256cd-158">このオブジェクトに関する詳細情報については [result](custom-functions-json.md#result) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="256cd-158">For detailed information about this object, see [result](custom-functions-json.md#result).</span></span> |
| `parameters` | <span data-ttu-id="256cd-159">関数の入力パラメーターを定義する配列です。</span><span class="sxs-lookup"><span data-stu-id="256cd-159">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="256cd-160">このオブジェクトに関する詳細情報については [parameters](custom-functions-json.md#parameters) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="256cd-160">For detailed information about this object, see [parameters](custom-functions-json.md#parameters).</span></span> |
| `options` | <span data-ttu-id="256cd-161">Excel で関数を実行する方法とタイミングの一部をユーザーがカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="256cd-161">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="256cd-162">このプロパティの使用方法の詳細については、「[ストリーム関数](#streaming-functions)」および「[関数のキャンセル](#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="256cd-162">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [canceling a function](#canceling-a-function).</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="256cd-163">マニフェスト ファイル</span><span class="sxs-lookup"><span data-stu-id="256cd-163">Manifest file</span></span>

<span data-ttu-id="256cd-164">カスタム関数 (Yo Office ジェネレーターが作成するプロジェクトでは **./manifest.xml**) を定義するアドインの XML マニフェスト ファイルは、アドイン内のすべてのカスタム関数の名前空間と、 JavaScript、JSON、および HTML の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="256cd-164">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="256cd-165">次の XML マークアップでは、`<ExtensionPoint>` と `<Resources>` カスタム関数を有効にするアドインのマニフェストに含める必要がある要素の例を示します。</span><span class="sxs-lookup"><span data-stu-id="256cd-165">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://localhost:8081/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://localhost:8081/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://localhost:8081/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="256cd-166">Excel の関数は、XML マニフェスト ファイルで指定された名前空間が接頭辞として付加されます。</span><span class="sxs-lookup"><span data-stu-id="256cd-166">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="256cd-167">関数の名前空間は、関数名の前に付けられ、ピリオドで区切られます。</span><span class="sxs-lookup"><span data-stu-id="256cd-167">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="256cd-168">例えば、Excel ワークシートのセル内で、`ADD42` 関数を呼び出すためには、`=CONTOSO.ADD42` と入力します。これは、`CONTOSO` が名前空間で、`ADD42` が JSON ファイルで指定された関数の名前だからです。</span><span class="sxs-lookup"><span data-stu-id="256cd-168">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="256cd-169">名前空間は、会社またはアドインの識別子としての使用を目的としています。</span><span class="sxs-lookup"><span data-stu-id="256cd-169">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="256cd-170">名前空間にはアルファベットとピリオドのみを含めることが出来ます。</span><span class="sxs-lookup"><span data-stu-id="256cd-170">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="declaring-a-volatile-function"></a><span data-ttu-id="256cd-171">揮発性関数の宣言</span><span class="sxs-lookup"><span data-stu-id="256cd-171">Declaring a volatile function</span></span>

<span data-ttu-id="256cd-172">[揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)とは、関数のいずれの引数にも変更がない場合でも、値が刻々と変化する関数のことです。</span><span class="sxs-lookup"><span data-stu-id="256cd-172">[Volatile functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed.</span></span> <span data-ttu-id="256cd-173">これらの関数は、Excel が再計算するたびに再計算を行います。</span><span class="sxs-lookup"><span data-stu-id="256cd-173">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="256cd-174">たとえば、`NOW` 関数を呼び出すセルがあるとします。</span><span class="sxs-lookup"><span data-stu-id="256cd-174">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="256cd-175">`NOW` が呼び出される度に、現在の日付と時刻を自動的に返します。</span><span class="sxs-lookup"><span data-stu-id="256cd-175">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="256cd-176">Excel には、`RAND` や `TODAY` などの組み込み揮発性関数がいくつか含まれています。</span><span class="sxs-lookup"><span data-stu-id="256cd-176">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="256cd-177">Excel の揮発性関数の完全なリストは、「[揮発性および非揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="256cd-177">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="256cd-178">カスタム関数を使用すると独自の揮発性関数を作成することができ、日時、時間、乱数、およびモデルを処理するときに役立つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="256cd-178">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling.</span></span> <span data-ttu-id="256cd-179">たとえば、モンテカルロ シミュレーションでは、最適なソリューションを決定するにはランダムな入力値の生成が必要です。</span><span class="sxs-lookup"><span data-stu-id="256cd-179">For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="256cd-180">関数を揮発性であると宣言するには、次のコードで示されるように、JSON メタデータファイルの関数で、`options` オブジェクトに`"volatile": true` を追加します。</span><span class="sxs-lookup"><span data-stu-id="256cd-180">To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample.</span></span> <span data-ttu-id="256cd-181">関数で `"streaming": true`と`"volatile": true` の両方をマークすることはできません。両方とも `true` とマークされている場合、揮発性のオプションは無視されます。</span><span class="sxs-lookup"><span data-stu-id="256cd-181">Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.</span></span>

```json
{
 "id": "TOMORROW",
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

## <a name="saving-and-sharing-state"></a><span data-ttu-id="256cd-182">状態の保存と共有</span><span class="sxs-lookup"><span data-stu-id="256cd-182">Saving and sharing state</span></span>

<span data-ttu-id="256cd-183">カスタム関数は、グローバル JavaScript 変数にデータを保存でき、以降の呼び出しで使用することができます。</span><span class="sxs-lookup"><span data-stu-id="256cd-183">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="256cd-184">保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を呼び出す場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="256cd-184">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="256cd-185">たとえば、同じ Web リソースへの追加呼び出しを避けるために、呼び出しから返されたデータを Web リソースに保存することができます。</span><span class="sxs-lookup"><span data-stu-id="256cd-185">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="256cd-186">次のコード サンプルでは、状態をグローバルに保存する温度ストリーミング関数の実装を示します。</span><span class="sxs-lookup"><span data-stu-id="256cd-186">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="256cd-187">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="256cd-187">Note the following about this code:</span></span>

- <span data-ttu-id="256cd-188">`streamTemperature` 関数がセルに表示される温度の値を毎秒更新し、`savedTemperatures` 変数をデータ ソースとして使用します。</span><span class="sxs-lookup"><span data-stu-id="256cd-188">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="256cd-189">`streamTemperature` はストリーム関数であるため、その関数がキャンセルされたときに実行されるキャンセル ハンドラーを実装します。</span><span class="sxs-lookup"><span data-stu-id="256cd-189">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="256cd-190">ユーザーが `streamTemperature` 関数を Excel の複数のセルから呼び出す場合、`streamTemperature` 関数は実行のたびに、同じ `savedTemperatures` 変数からのデータを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="256cd-190">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="256cd-191">`refreshTemperature` 関数は、特定の温度計の温度を毎秒読み取り、結果を `savedTemperatures` 変数に格納します。</span><span class="sxs-lookup"><span data-stu-id="256cd-191">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="256cd-192">`refreshTemperature` 関数は、Excel でエンド ユーザーには公開されないので、JSON ファイルに登録する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="256cd-192">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="coauthoring"></a><span data-ttu-id="256cd-193">共同編集</span><span class="sxs-lookup"><span data-stu-id="256cd-193">Coauthoring</span></span>

<span data-ttu-id="256cd-194">Excel Online と Excel for Windows で Office 365 サブスクリプションを利用している場合、ドキュメントの共同編集を行うことができ、カスタム関数を使用できます。</span><span class="sxs-lookup"><span data-stu-id="256cd-194">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="256cd-195">ブックでカスタム関数を使用している場合、仕事仲間はカスタム関数のアドインを読み込むように要求されます。</span><span class="sxs-lookup"><span data-stu-id="256cd-195">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="256cd-196">双方がアドインを読み込むと、共同編集によりカスタム関数は結果を共有します。</span><span class="sxs-lookup"><span data-stu-id="256cd-196">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="256cd-197">共同編集の詳細については、「[Excel での共同編集](/office/vba/excel/concepts/about-coauthoring-in-excel)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="256cd-197">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="256cd-198">データの範囲を使用する</span><span class="sxs-lookup"><span data-stu-id="256cd-198">Working with ranges of data</span></span>

<span data-ttu-id="256cd-199">カスタム関数は、データの範囲を入力パラメーターとして受け入れることができ、また、データの範囲を返すこともできます。</span><span class="sxs-lookup"><span data-stu-id="256cd-199">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="256cd-200">JavaScript では、データの範囲は 2 次元配列として表されます。</span><span class="sxs-lookup"><span data-stu-id="256cd-200">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="256cd-201">例えば、関数が Excel に保存されている数値の範囲から 2 番目に大きい値を返すとします。</span><span class="sxs-lookup"><span data-stu-id="256cd-201">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="256cd-202">次の関数は、`Excel.CustomFunctionDimensionality.matrix` 型の `values` パラメーターを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="256cd-202">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="256cd-203">なお、この関数の JSON メタデータでは、パラメーターの`type`プロパティを`matrix` と設定します。</span><span class="sxs-lookup"><span data-stu-id="256cd-203">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="determine-which-cell-invoked-your-custom-function"></a><span data-ttu-id="256cd-204">カスタム関数が呼び出したセルを特定する</span><span class="sxs-lookup"><span data-stu-id="256cd-204">Determine which cell invoked your custom function</span></span>

<span data-ttu-id="256cd-205">場合によっては、カスタム関数が呼び出したセルのアドレスを取得する必要が生じます。</span><span class="sxs-lookup"><span data-stu-id="256cd-205">In some cases you'll need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="256cd-206">これは、次の種類のシナリオで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="256cd-206">This may be useful in the following types of scenarios:</span></span>

- <span data-ttu-id="256cd-207">範囲の書式設定: [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data) で情報を格納するキーとしてセル アドレスを使用します。</span><span class="sxs-lookup"><span data-stu-id="256cd-207">Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="256cd-208">Excel で [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) を使用して`AsyncStorage` からキーを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="256cd-208">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="256cd-209">キャッシュされた値を表示させる: 関数がオフラインで使用される場合、`onCalculated` を使用して `AsyncStorage` に格納されているキャッシュされた値を表示します。</span><span class="sxs-lookup"><span data-stu-id="256cd-209">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="256cd-210">調整: セル アドレスを使用して元のセルを検出し、処理が発生している場所での調整を行えます。</span><span class="sxs-lookup"><span data-stu-id="256cd-210">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="256cd-211">セルのアドレスに関する情報は、関数の JSON メタデータ ファイルで `requiresAddress` が`true` とマークされている場合にのみ公開されます。</span><span class="sxs-lookup"><span data-stu-id="256cd-211">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="256cd-212">これの例を次のサンプルに示します。</span><span class="sxs-lookup"><span data-stu-id="256cd-212">The following sample gives an example of this:</span></span>

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

<span data-ttu-id="256cd-213">セルのアドレスを検索するために、スクリプト ファイル (**./src/customfunctions.js**または **./src/customfunctions.ts**) に `getAddress` 関数を追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="256cd-213">In the script file (**./src/customfunctions.js** or **./src/customfunctions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="256cd-214">この関数は、次のサンプルで示される `parameter1` のようなパラメーターを受け取ることができます。</span><span class="sxs-lookup"><span data-stu-id="256cd-214">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="256cd-215">最後のパラメーターは常に `invocationContext` で、これはJSON メタデータ ファイルで `requiresAddress` が `true` とマークされているときに Excel が返すセルの位置が格納されているオブジェクトのことです。</span><span class="sxs-lookup"><span data-stu-id="256cd-215">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="256cd-216">既定では、`getAddress` 関数が返す値は次の形式に従います: `SheetName!CellNumber`。</span><span class="sxs-lookup"><span data-stu-id="256cd-216">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="256cd-217">たとえば、ある関数が Expenses という名前のシートのセル B2 から呼び出される場合の戻り値は `Expenses!B2` になります。</span><span class="sxs-lookup"><span data-stu-id="256cd-217">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="known-issues"></a><span data-ttu-id="256cd-218">既知の問題</span><span class="sxs-lookup"><span data-stu-id="256cd-218">Known issues</span></span>

<span data-ttu-id="256cd-219">既知の問題については、[Excel カスタム関数についての GitHub のレポート](https://github.com/OfficeDev/Excel-Custom-Functions/issues)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="256cd-219">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="see-also"></a><span data-ttu-id="256cd-220">関連項目</span><span class="sxs-lookup"><span data-stu-id="256cd-220">See also</span></span>

* [<span data-ttu-id="256cd-221">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="256cd-221">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="256cd-222">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="256cd-222">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="256cd-223">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="256cd-223">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="256cd-224">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="256cd-224">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="256cd-225">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="256cd-225">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
