---
ms.date: 10/09/2018
description: JavaScript を使用して Excel でカスタム関数を作成します。
title: Excel でのカスタム関数の作成 (プレビュー)
ms.openlocfilehash: e52039f2618f793f688cd89c5d62bac0a8632667
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506120"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="cd57a-103">Excel でカスタム関数を作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="cd57a-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="cd57a-104">カスタム関数を使用すると、開発者は JavaScript でこれらの関数をアドインの一部として定義することにより、 Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="cd57a-105">Excel 内のユーザーは、Excel の他のネイティブ関数（`SUM()` など）とまったく同様に、カスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="cd57a-106">この記事では、Excel でカスタム関数を作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="cd57a-p102">次の図は、Excel ワークシートのセルにカスタム関数を挿入する、エンド ユーザーを示します。 `CONTOSO.ADD42` カスタム関数は、関数への入力パラメーターとしてユーザーが指定した数値のペアに 42 を追加するように設計されています。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p102">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet. The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="cd57a-109">次のコードは、`ADD42` カスタム関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="cd57a-110">この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="cd57a-111">カスタム関数アドインプロジェクトのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="cd57a-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="cd57a-112">[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office) を使用して Excel カスタム関数アドイン プロジェクトを作成する場合は、ジェネレーターが作成するプロジェクトに以下のようなファイルが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="cd57a-113">ファイル</span><span class="sxs-lookup"><span data-stu-id="cd57a-113">File</span></span> | <span data-ttu-id="cd57a-114">ファイル形式</span><span class="sxs-lookup"><span data-stu-id="cd57a-114">File format</span></span> | <span data-ttu-id="cd57a-115">説明</span><span class="sxs-lookup"><span data-stu-id="cd57a-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="cd57a-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="cd57a-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="cd57a-117">または</span><span class="sxs-lookup"><span data-stu-id="cd57a-117">or</span></span><br/><span data-ttu-id="cd57a-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="cd57a-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="cd57a-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="cd57a-119">JavaScript</span></span><br/><span data-ttu-id="cd57a-120">または</span><span class="sxs-lookup"><span data-stu-id="cd57a-120">or</span></span><br/><span data-ttu-id="cd57a-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="cd57a-121">TypeScript</span></span> | <span data-ttu-id="cd57a-122">カスタム関数を定義するコードを含みます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="cd57a-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="cd57a-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="cd57a-124">JSON</span><span class="sxs-lookup"><span data-stu-id="cd57a-124">JSON</span></span> | <span data-ttu-id="cd57a-125">カスタム関数を定義し、Excel に関数を登録してエンドユーザーが使用できるようにするためのメタデータを含みます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-125">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="cd57a-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="cd57a-126">**./index.html**</span></span> | <span data-ttu-id="cd57a-127">HTML</span><span class="sxs-lookup"><span data-stu-id="cd57a-127">HTML</span></span> | <span data-ttu-id="cd57a-128">カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="cd57a-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="cd57a-129">**Manifest.xml**</span></span> | <span data-ttu-id="cd57a-130">XML</span><span class="sxs-lookup"><span data-stu-id="cd57a-130">XML</span></span> | <span data-ttu-id="cd57a-131">アドイン内のすべてのカスタム関数の名前空間と、この表で前述した JavaScript、JSON、HTML ファイルの位置を指定します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="cd57a-132">次のセクションでは、これらのファイルに関する詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-132">The following sections provide more details about these settings.</span></span>

### <a name="script-file"></a><span data-ttu-id="cd57a-133">スクリプト ファイル</span><span class="sxs-lookup"><span data-stu-id="cd57a-133">Script file</span></span> 

<span data-ttu-id="cd57a-134">スクリプト ファイル (Yo Office ジェネレーターが作成するプロジェクト内の **./src/customfunctions.js** または **./src/customfunctions.ts**) には、カスタム関数を定義して、カスタム関数の名前を [JSON メタデータ ファイル](#json-metadata-file)のオブジェクトにマップするコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="cd57a-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="cd57a-p103">例えば、次のコードでカスタム関数 `add` と `increment` を定義し、両方の関数のマッピング情報を指定します。 `add` 関数は、JSON メタデータ ファイル内のオブジェクトにマップされ、 この場所に`id` プロパティの値が**追加**されます。`increment` 関数は、メタデータ ファイル内のオブジェクトにマップされ、この場所に`id` プロパティの値が**インクリメント**します。JSON メタデータ ファイル内のオブジェクトへのスクリプト ファイル内関数名のマッピングの詳細については、 「 [カスタム関数のベスト プラクティス](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) 」 を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p103">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions. The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**. See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

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

### <a name="json-metadata-file"></a><span data-ttu-id="cd57a-138">JSON メタデータ ファイル</span><span class="sxs-lookup"><span data-stu-id="cd57a-138">JSON metadata file</span></span> 

<span data-ttu-id="cd57a-p104">カスタム関数のメタデータ ファイル (Yo Office ジェネレーターが作成するプロジェクトでは **./config/customfunctions.json** ) は、Excel がカスタム関数の登録を要求し、エンドユーザーが利用できるよう、情報を提供します。カスタム関数は、ユーザーがアドインを初めて実行するときに登録されます。その後は、同じユーザーに対しては、（アドインが最初に実行されたブック内のみでなく）すべてのブック内で利用が可能になります。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p104">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users. Custom functions are registered when a user runs an add-in for the first time. After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="cd57a-142">JSON ファイルをホストするサーバーは、カスタム関数が Excel Online で正しく作動するために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS)  を有効に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cd57a-142">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="cd57a-p105">**Customfunctions.json** の次のコードは、上述の `add` 関数および `increment` 関数のメタデータを指定します。このコード サンプルを基にした表は、この JSON オブジェクト内の個別のプロパティについての詳細情報を提供します。JSON メタデータ ファイル内の `id` および `name` プロパティの指定に関する詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p105">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously. The table that follows this code sample provides detailed information about the individual properties within this JSON object. See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="cd57a-p106">次の表は、通常、JSON メタデータ ファイルに格納されているプロパティの一覧表示です。JSON メタデータ ファイルの詳細については、[カスタム関数のメタデータ](custom-functions-json.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p106">The following table lists the properties that are typically present in the JSON metadata file. For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="cd57a-148">プロパティ</span><span class="sxs-lookup"><span data-stu-id="cd57a-148">Property</span></span>  | <span data-ttu-id="cd57a-149">説明</span><span class="sxs-lookup"><span data-stu-id="cd57a-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="cd57a-p107">関数のユニーク ID です。設定後、この ID は変更できません。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p107">A unique ID for the function. This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="cd57a-p108">Excel でエンドユーザーに表示される関数の名前です。Excel では、この関数名の前に、[XML マニフェスト ファイル](#manifest-file)で指定されているカスタム関数の名前空間が接頭辞として付されます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p108">Name of the function that the end user sees in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="cd57a-154">ユーザーがヘルプを要求したときに表示されるページの URL です。</span><span class="sxs-lookup"><span data-stu-id="cd57a-154">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="cd57a-p109">関数について説明します。この値は、関数が Excel 内のオートコンプリート メニューで選択された項目となっている場合に、ツールヒントとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p109">Describes what the function does. This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="cd57a-p110">関数によって返される情報の種類を定義するオブジェクト。`type` 子プロパティの値は、 **文字列**、 **数値**、または **ブール値**を使用できます。子プロパティの値は、 `dimensionality` **スカラー** または **マトリックス** を使用できます (指定された `type`の値の 2 次元配列)。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p110">Object that defines the type of information that is returned by the function. The value of the `type` child property can be **string**, **number**, or **boolean**. The value of the `dimensionality` child property can be **scalar** or **matrix** (a two-dimensional array of values of the specified `type`).</span></span> |
| `parameters` | <span data-ttu-id="cd57a-p111">関数の入力パラメーターを定義する配列。 `name` と `description` Excel の intelliSense の子のプロパティが表示されます。 `type` 子プロパティの値には、 **文字列**、 **数値**、または **ブール値**を使用できます。`dimensionality` 子プロパティの値には、**スカラー** または **マトリックス** を使用できます (指定された `type`の値の 2 次元配列)。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p111">Array that defines the input parameters for the function. The `name` and `description` child properties appear in the Excel intelliSense. The value of the `type` child property can be **string**, **number**, or **boolean**. The value of the `dimensionality` child property can be **scalar** or **matrix** (a two-dimensional array of values of the specified `type`).</span></span> |
| `options` | <span data-ttu-id="cd57a-p112">Excel で関数を実行する方法とタイミングのいくつかの側面をカスタマイズできます。このプロパティの使用方法の詳細については、この記事で後述する 「[ストリーム関数](#streaming-functions)」と「[関数のキャンセル](#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p112">Enables you to customize some aspects of how and when Excel executes the function. For more information about how this property can be used, see [Streamed functions](#streaming-functions) and [Canceling a function](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="cd57a-166">マニフェスト ファイル</span><span class="sxs-lookup"><span data-stu-id="cd57a-166">Manifest file</span></span>

<span data-ttu-id="cd57a-p113">カスタム関数 (Yo Office ジェネレーターが作成するプロジェクトでは **./manifest.xml** ) を定義するアドインの XML マニフェスト ファイルは、アドインとJavaScript、JSON、および HTML のロケーション内のすべてのカスタム関数の名前空間を指定します。次の XML マークアップでは、 `<ExtensionPoint>` と `<Resources>` カスタム関数を有効にするアドインのマニフェストに含める必要がある要素の一例を示します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p113">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files. The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

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
> <span data-ttu-id="cd57a-p114">Excel の関数には、XML マニフェスト ファイルで指定された名前空間が接頭辞として付加されます。関数の名前空間は、関数名の前に付けられ、ピリオドで区切られます。例えば、Excel ワークシートのセル内で、関数を呼び出すためには `ADD42` 、 `=CONTOSO.ADD42`を入力します。これは、CONTOSO が、名前空間であり、 `ADD42` JSON ファイルで指定された関数の名前であるためです。名前空間は、会社またはアドインの識別子としての使用を目的としています。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p114">Functions in Excel are prepended by the namespace specified in your XML manifest file. A function's namespace comes before the function name and they are separated by a period. For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file. The namespace is intended to be used as an identifier for your company or the add-in.</span></span> 

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="cd57a-173">外部ソースからデータを返す関数</span><span class="sxs-lookup"><span data-stu-id="cd57a-173">Functions that return data from external sources</span></span>

<span data-ttu-id="cd57a-174">カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cd57a-174">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="cd57a-175">JavaScript Promise を Excel に返す。</span><span class="sxs-lookup"><span data-stu-id="cd57a-175">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="cd57a-176">コールバック関数を使用して Promise を最終値で解決する。</span><span class="sxs-lookup"><span data-stu-id="cd57a-176">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="cd57a-p115">カスタム関数は、 Excel が `#GETTING_DATA` セルの最終結果を待っている間、一時的な結果を表示します。ユーザーは、結果待機中も通常はワークシートの残りの部分を操作することができます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p115">Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result. Users can interact normally with the rest of the worksheet while they wait for the result.</span></span>

<span data-ttu-id="cd57a-p116">次のコード例は、 温度計の現在温度を取得する `getTemperature()` カスタム関数です。 `sendWebRequest` は、温度 web サービスを呼び出す [XHR](custom-functions-runtime.md#xhr-example) を使用した仮想関数 (ここでは指定なし) であることに留意してください。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p116">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer. Note that `sendWebRequest` is a hypothetical function (not specified here) that uses [XHR](custom-functions-runtime.md#xhr-example) to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a><span data-ttu-id="cd57a-181">ストリーミング関数</span><span class="sxs-lookup"><span data-stu-id="cd57a-181">Streaming functions</span></span>

<span data-ttu-id="cd57a-p117">ストリーミングのカスタム関数を使用すると、データ更新を明確に要求するユーザーを必要とせず、時間の経過と共に繰り返しセルにデータを出力します。次のコード サンプルは、1 秒ごとの結果の数値を追加するカスタム関数です。このコードについては、以下のことに留意してください。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p117">Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh. The following code sample is a custom function that adds a number to the result every second. Note the following about this code:</span></span>

- <span data-ttu-id="cd57a-185">Excel は、`setResult`コールバックを使用して自動的に新しい値を表示します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-185">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="cd57a-186">[オートコンプリート] メニューから関数を選択する場合、2 番目の入力パラメータ `handler` は、Excel のエンドユーザーに表示されません。</span><span class="sxs-lookup"><span data-stu-id="cd57a-186">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="cd57a-p118">`onCanceled` コールバックは、関数がキャンセルされたときに実行される関数を定義します。どのストリーミング関数に対してもキャンセル ハンドラーを実装する必要があります。詳細については、 「[関数のキャンセル](#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p118">The `onCanceled` callback defines the function that executes when the function is canceled. You must implement a cancellation handler like this for any streamed function. For more information, see [Canceling a function](#canceling-a-function).</span></span>

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

<span data-ttu-id="cd57a-190">JSON メタデータ ファイルでストリーミング関数にメタデータを指定する場合には、以下の例のように、プロパティ`"cancelable": true` および `options` オブジェクト内の `"stream": true` を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cd57a-190">When you specify metadata for a streamed function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

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

## <a name="canceling-a-function"></a><span data-ttu-id="cd57a-191">関数のキャンセル
</span><span class="sxs-lookup"><span data-stu-id="cd57a-191">Canceling a function</span></span>

<span data-ttu-id="cd57a-p119">状況によっては、帯域幅の消費、作業メモリ、および CPU への負荷を縮小するために、ストリーミング カスタム関数の実行をキャンセルする必要が生じます。Excel では、次のような場合、関数の実行をキャンセルします。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p119">In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load. Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="cd57a-194">ユーザーが、関数への参照があるセルを編集または削除した場合。</span><span class="sxs-lookup"><span data-stu-id="cd57a-194">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="cd57a-p120">関数の引数 (入力) のいずれかが変更されたとき。この例では、キャンセルに続いて新しい関数呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p120">When one of the arguments (inputs) for the function changes. In this case, a new function call is triggered following the cancellation.</span></span>

- <span data-ttu-id="cd57a-p121">ユーザーが手動で再計算をトリガーしたとき。この例では、キャンセルに続いて新しい関数呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p121">When the user triggers recalculation manually. In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="cd57a-p122">関数をキャンセルする機能を有効にするには、JavaScript 関数内にキャンセル ハンドラーを実装し、関数を定義する JSON のメタデータの`options` オブジェクト内のプロパティ `"cancelable": true` を指定する必要があります。この記事の前のセクションのコード サンプルに、これらの手法の例が示されています。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p122">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function. The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="cd57a-201">状態の保存と共有</span><span class="sxs-lookup"><span data-stu-id="cd57a-201">Saving and sharing state</span></span>

<span data-ttu-id="cd57a-p123">カスタム関数は、以降の呼び出しで使用できるJavaScript のグローバル変数にデータを保存できます。保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を呼び出す場合に便利です。例えば、同一の Web リソースへの追加呼び出しを避けるため、Web リソースへの呼び出しから返されたデータを保存することができます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p123">Custom functions can save data in global JavaScript variables. In subsequent calls, your custom function may use the values saved in these variables. Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state. For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="cd57a-205">以下のコード サンプルは、 状態をグローバルで保存する温度ストリーミング関数の実装を示しています。</span><span class="sxs-lookup"><span data-stu-id="cd57a-205">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="cd57a-206">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="cd57a-206">Note the following about this code:</span></span>

- <span data-ttu-id="cd57a-207">`streamTemperature` 関数が 毎秒セルに表示される温度の値を更新し、 `savedTemperatures` 変数をデータ ソースとして使用します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-207">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="cd57a-208">`streamTemperature` は、ストリーム関数であるため、その関数がキャンセルされたときに実行されるキャンセル ハンドラーを実装します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-208">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="cd57a-209">ユーザーが `streamTemperature` 関数を Excel の複数のセルから呼び出す場合、 `streamTemperature` 関数は実行のたびに、同じ `savedTemperatures` 変数からのデータを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="cd57a-209">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="cd57a-210">`refreshTemperature` 関数は、毎秒特定の温度計の温度を読み取り、結果を `savedTemperatures` 変数に格納します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-210">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="cd57a-211">`refreshTemperature` 関数は、Excel でのエンド ユーザーには公開されないので、JSON ファイルに登録する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="cd57a-211">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="cd57a-212">データの範囲の操作</span><span class="sxs-lookup"><span data-stu-id="cd57a-212">Working with ranges of data</span></span>

<span data-ttu-id="cd57a-213">カスタム関数は、入力パラメーターとしてデータの範囲を受け取ることができます。または、データの範囲を返すこともできます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-213">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="cd57a-214">JavaScript では、データの範囲は、2 次元配列として表されます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-214">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="cd57a-215">たとえば、関数が Excel に格納されている数値の範囲から 2 番目に高い値を返すとします。</span><span class="sxs-lookup"><span data-stu-id="cd57a-215">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="cd57a-216">以下の関数が、タイプ `Excel.CustomFunctionDimensionality.matrix` のものである `values` パラメーターを受け取ります。</span><span class="sxs-lookup"><span data-stu-id="cd57a-216">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="cd57a-217">この関数の JSON メタデータでは、パラメーターの `type` プロパティを `matrix` に設定するように注意してください。</span><span class="sxs-lookup"><span data-stu-id="cd57a-217">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="cd57a-218">エラーの処理</span><span class="sxs-lookup"><span data-stu-id="cd57a-218">Handling errors</span></span>

<span data-ttu-id="cd57a-p128">カスタム関数を定義するアドインをビルドする場合は、ランタイム エラーを考慮するためのエラー処理 ロジックを含めるようにしてください。カスタム関数のエラー処理は、 [大規模な Excel の JavaScript API のエラー処理](excel-add-ins-error-handling.md)と同じです。次のコード サンプルでは、 `.catch`がコード内で以前に発生したエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-p128">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="cd57a-222">既知の問題</span><span class="sxs-lookup"><span data-stu-id="cd57a-222">Known issues</span></span>

- <span data-ttu-id="cd57a-223">ヘルプの URL とパラメーターの説明。Excel ではまだ使用されていません。</span><span class="sxs-lookup"><span data-stu-id="cd57a-223">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="cd57a-224">カスタム関数は現在、モバイル クライアント用の Excel では使用できません。</span><span class="sxs-lookup"><span data-stu-id="cd57a-224">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="cd57a-225">揮発性関数（スプレッドシート内の無関係なデータが変更されたときに自動的に再計算する関数）はまだサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cd57a-225">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="cd57a-226">Office 365 管理ポータルと AppSource による展開はまだ有効になっていません。</span><span class="sxs-lookup"><span data-stu-id="cd57a-226">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="cd57a-227">Excel Online のカスタム関数は、一定期間使用しないとセッション中に機能しなくなることがあります。</span><span class="sxs-lookup"><span data-stu-id="cd57a-227">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="cd57a-228">ブラウザページを更新（F5）し、カスタム関数を再入力して機能を復元します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-228">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="cd57a-229">Excel for Windows で実行されている複数のアドインがある場合には、ワークシートのセル内に **#GETTING_DATA** の一時的な結果が表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="cd57a-229">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="cd57a-230">すべての Excel ウィンドウを閉じて、Excel を再起動します。</span><span class="sxs-lookup"><span data-stu-id="cd57a-230">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="cd57a-231">将来的には、カスタム関数用のデバッグ ツールが利用可能となる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="cd57a-231">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="cd57a-232">それまでは、F12 開発者ツールを使用して Excel オンラインでデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="cd57a-232">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="cd57a-233">詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cd57a-233">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="cd57a-234">変更ログ</span><span class="sxs-lookup"><span data-stu-id="cd57a-234">Changelog</span></span>

- <span data-ttu-id="cd57a-235">**2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開\*</span><span class="sxs-lookup"><span data-stu-id="cd57a-235">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="cd57a-236">**2017 年 11 月 20 日**: ビルド 8801 以降を使用しているユーザー向けに互換性バグを修正</span><span class="sxs-lookup"><span data-stu-id="cd57a-236">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="cd57a-237">**2017 年 11 月 28 日**: 非同期関数のキャンセルへのサポートを公開\* (ストリーム関数への変更が必要)</span><span class="sxs-lookup"><span data-stu-id="cd57a-237">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="cd57a-238">**2018 年 5 月 7 日**: Mac、Excel Online、およびインプロセスで実行される同期関数へのサポートを公開\*</span><span class="sxs-lookup"><span data-stu-id="cd57a-238">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="cd57a-239">**2018 年 9 月 20日**: JavaScript 実行時のカスタム関数へのサポートを公開</span><span class="sxs-lookup"><span data-stu-id="cd57a-239">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="cd57a-240">詳細については、「[Excel カスタム関数のランタイム](custom-functions-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cd57a-240">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="cd57a-241">\* Office Insiders チャネル対象</span><span class="sxs-lookup"><span data-stu-id="cd57a-241">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="cd57a-242">関連項目</span><span class="sxs-lookup"><span data-stu-id="cd57a-242">See also</span></span>

* [<span data-ttu-id="cd57a-243">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="cd57a-243">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="cd57a-244">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="cd57a-244">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="cd57a-245">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="cd57a-245">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="cd57a-246">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="cd57a-246">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)