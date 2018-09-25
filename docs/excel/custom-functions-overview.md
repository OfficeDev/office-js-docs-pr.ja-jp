---
ms.date: 09/20/2018
description: JavaScript を使用して Excel でカスタム関数を作成します。
title: Excel でのカスタム関数の作成 (プレビュー)
ms.openlocfilehash: b214329fe50955d0f39d50f674152f475ca24b4d
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/25/2018
ms.locfileid: "25005044"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="c1848-103">Excel でのカスタム関数の作成 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="c1848-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="c1848-104">開発者はカスタム関数を使用すれば、アドインの一部として新しい関数を定義して、これらの関数を追加することができます。</span><span class="sxs-lookup"><span data-stu-id="c1848-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="c1848-105">Excel 内のユーザーは、Excel の他のネイティブ関数 (`SUM()` など) と同様に、カスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="c1848-105">Users within Excel can access custom functions like any other native function in Excel (such as `SUM()`).</span></span> <span data-ttu-id="c1848-106">この記事では、Excel でカスタム関数を作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="c1848-106">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="c1848-107">次の図では、エンド ユーザーが Excel ワークシートのセルにカスタム関数を挿入する例を示します。</span><span class="sxs-lookup"><span data-stu-id="c1848-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="c1848-108">カスタム関数は、ユーザーが関数への入力パラメーターとして指定する数値ペアに、42 を足すように設計されています。`CONTOSO.ADD42`</span><span class="sxs-lookup"><span data-stu-id="c1848-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="c1848-109">次のコードは、`ADD42` カスタム関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="c1848-109">The following code defines the `ADD42` custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="c1848-110">カスタム機能は、Windows、Mac、および Excel Online の開発者プレビューで利用できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="c1848-110">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="c1848-111">これらを試すには、以下の手順を行います。</span><span class="sxs-lookup"><span data-stu-id="c1848-111">To try them, complete these steps:</span></span>

1. <span data-ttu-id="c1848-112">Office (Windows はビルド 10827、Mac は 13.329) をインストールし、 [Office Insider](https://products.office.com/office-insider) プログラムに参加します。</span><span class="sxs-lookup"><span data-stu-id="c1848-112">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/office-insider) program.</span></span> <span data-ttu-id="c1848-113">カスタム関数へのアクセスを取得するには、Office Insider プログラムに参加する必要があります。現時点では、Office Insider プログラムのメンバーでない限り、カスタム関数はすべての Office のビルド間で無効となっています。</span><span class="sxs-lookup"><span data-stu-id="c1848-113">You must join the Office Insider program in order to have access to custom functions; currently, custom functions are disabled across all Office builds unless you are a member of the Office Insider program.</span></span>

2. <span data-ttu-id="c1848-114">[Yo Office](https://github.com/OfficeDev/generator-office) を使用して Excel カスタム関数のアドイン プロジェクトを作成し、[OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) の指示に従ってプロジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="c1848-114">Use [Yo Office](https://github.com/OfficeDev/generator-office) to create an Excel Custom Functions add-in project, and then follow the instructions in the [OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) to use the project.</span></span>

3. <span data-ttu-id="c1848-115">Excel ワークシートの任意のセルに `=CONTOSO.ADD42(1,2)` と入力し、**Enter** キーを押してカスタム関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="c1848-115">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

> [!NOTE]
> <span data-ttu-id="c1848-116">この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。</span><span class="sxs-lookup"><span data-stu-id="c1848-116">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="c1848-117">基本操作の説明</span><span class="sxs-lookup"><span data-stu-id="c1848-117">Learn the basics</span></span>

<span data-ttu-id="c1848-118">[Yo Office](https://github.com/OfficeDev/generator-office) を使用して作成したカスタム関数プロジェクトに、以下のファイルが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c1848-118">In the custom functions project that you've created using [Yo Office](https://github.com/OfficeDev/generator-office), you’ll see the following files:</span></span>

| <span data-ttu-id="c1848-119">ファイル</span><span class="sxs-lookup"><span data-stu-id="c1848-119">File</span></span> | <span data-ttu-id="c1848-120">ファイル形式</span><span class="sxs-lookup"><span data-stu-id="c1848-120">File format</span></span> | <span data-ttu-id="c1848-121">説明</span><span class="sxs-lookup"><span data-stu-id="c1848-121">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="c1848-122">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="c1848-122">**./src/customfunctions.js**</span></span> | <span data-ttu-id="c1848-123">JavaScript</span><span class="sxs-lookup"><span data-stu-id="c1848-123">JavaScript</span></span> | <span data-ttu-id="c1848-124">カスタム関数を定義するコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="c1848-124">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="c1848-125">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="c1848-125">**./config/customfunctions.json**</span></span> | <span data-ttu-id="c1848-126">JSON</span><span class="sxs-lookup"><span data-stu-id="c1848-126">JSON</span></span> | <span data-ttu-id="c1848-127">カスタム関数について説明し、エンドユーザーが使用可能なように、Excel で関数を登録できるようにするメタデータが含まれています。</span><span class="sxs-lookup"><span data-stu-id="c1848-127">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="c1848-128">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="c1848-128">**./index.html**</span></span> | <span data-ttu-id="c1848-129">HTML</span><span class="sxs-lookup"><span data-stu-id="c1848-129">HTML</span></span> | <span data-ttu-id="c1848-130">カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。</span><span class="sxs-lookup"><span data-stu-id="c1848-130">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="c1848-131">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="c1848-131">**Manifest.xml**</span></span> | <span data-ttu-id="c1848-132">XML</span><span class="sxs-lookup"><span data-stu-id="c1848-132">XML</span></span> | <span data-ttu-id="c1848-133">アドイン内のすべてのカスタム関数の名前空間と、このテーブルで前に一覧表示した JavaScript、JSON、HTML ファイルの位置を指定します。</span><span class="sxs-lookup"><span data-stu-id="c1848-133">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

### <a name="manifest-file-manifestxml"></a><span data-ttu-id="c1848-134">マニフェスト ファイル (./manifest.xml)</span><span class="sxs-lookup"><span data-stu-id="c1848-134">Manifest file (manifest.xml)</span></span>

<span data-ttu-id="c1848-135">カスタム関数を定義するアドイン用の XML マニフェスト ファイルでは、アドイン内のすべてのカスタム関数の名前空間と、JavaScript、JSON、HTML ファイルの位置を定義します。</span><span class="sxs-lookup"><span data-stu-id="c1848-135">The XML manifest file for an add-in that defines custom functions specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="c1848-136">次の XML マークアップの例では、Excel がカスタム関数を実行できるようにするための、アドインのマニフェストに含める必要のある `<ExtensionPoint>` および `<Resources>` 要素の例を示します。</span><span class="sxs-lookup"><span data-stu-id="c1848-136">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest in order to enable Excel to run custom functions.</span></span>  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. For example, a function named "ADD42" is invoked as `=CONTOSO.ADD42` in Excel.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="c1848-137">Excel 内の関数は、XML マニフェスト ファイルで指定される名前空間の先頭に追加されます。</span><span class="sxs-lookup"><span data-stu-id="c1848-137">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="c1848-138">関数の名前空間は関数名の前に配置され、それらはピリオドで区切られます。</span><span class="sxs-lookup"><span data-stu-id="c1848-138">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="c1848-139">たとえば、Excel ワークシートのセル内の関数 `ADD42()` を呼び出すには、`=CONTOSO.ADD42` と入力します。これは、CONTOSO が名前空間で、`ADD42` が JSON ファイルで指定された関数の名前であるからです。</span><span class="sxs-lookup"><span data-stu-id="c1848-139">For example, to call the function `ADD42()` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="c1848-140">名前空間は、所属する会社またはアドインの識別子として使用することを想定しています。</span><span class="sxs-lookup"><span data-stu-id="c1848-140">The prefix is intended to be used as an identifier for your add-in.</span></span> 

### <a name="json-file-configcustomfunctionsjson"></a><span data-ttu-id="c1848-141">JSON ファイル (./config/customfunctions.json)</span><span class="sxs-lookup"><span data-stu-id="c1848-141">JSON file (./config/customfunctions.json)</span></span>

<span data-ttu-id="c1848-142">カスタム関数のメタデータ ファイルは、Excel がカスタム関数を登録し、エンドユーザーが使用できるようにするために必要とする情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="c1848-142">A custom functions metadata file provides the information that Excel requires to register the custom functions and make them available to end-users.</span></span> <span data-ttu-id="c1848-143">カスタム関数は、ユーザーがはじめてアドインを実行したときに登録されます。</span><span class="sxs-lookup"><span data-stu-id="c1848-143">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="c1848-144">その後、その同じユーザーは、最初にアドインが実行されたブックだけでなく、すべてのブックでそれらのカスタム関数を使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="c1848-144">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

> [!TIP]
> <span data-ttu-id="c1848-145">JSON ファイルをホストするサーバーのサーバー設定では、カスタム関数が Excel Online で正しく作動するために、[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1848-145">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="c1848-146">以下の **customfunctions.json** のコードでは、この記事で前述した `ADD42` 関数のメタデータを指定します。</span><span class="sxs-lookup"><span data-stu-id="c1848-146">The following code in **customfunctions.json** specifies the metadata for the `ADD42` function that was described previously in this article.</span></span> <span data-ttu-id="c1848-147">このメタデータでは、関数の名前、説明、戻り値、入力パラメーターその他を定義します。</span><span class="sxs-lookup"><span data-stu-id="c1848-147">This metadata defines the function's name, description, return value, input parameters, and more.</span></span> <span data-ttu-id="c1848-148">このコード サンプルの次の表では、この JSON オブジェクト内の個々のプロパティについての詳細情報を示しています。</span><span class="sxs-lookup"><span data-stu-id="c1848-148">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span>

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [                {
                    "name": "number 1",
                    "description": "the first number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                },
                {
                    "name": "number 2",
                    "description": "the second number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
        }
    ]
}
```

<span data-ttu-id="c1848-149">以下の表では、通常 JSON メタデータ ファイルに格納されているプロパティを一覧表示しています。</span><span class="sxs-lookup"><span data-stu-id="c1848-149">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="c1848-150">前の例で使用されていないオプションを含む、JSON メタデータ ファイルの詳細情報については、「[カスタム関数のメタデータ](custom-functions-json.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c1848-150">For more detailed information about the JSON metadata file, including options not used in the previous example, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="c1848-151">プロパティ</span><span class="sxs-lookup"><span data-stu-id="c1848-151">Property</span></span>  | <span data-ttu-id="c1848-152">説明</span><span class="sxs-lookup"><span data-stu-id="c1848-152">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="c1848-153">関数の一意の ID です。</span><span class="sxs-lookup"><span data-stu-id="c1848-153">A unique ID for the group.</span></span> <span data-ttu-id="c1848-154">設定後は、この ID は変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="c1848-154">This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="c1848-155">ユーザーがセルに数式を入力した際に、オートコンプリート メニューに表示される関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="c1848-155">Name of the function that is shown in the autocomplete menu as a user types a formula within a cell.</span></span> <span data-ttu-id="c1848-156">オートコンプリート メニューでは、XML マニフェスト ファイルで指定されるカスタム関数の名前空間が、この値に接頭辞としてつきます。</span><span class="sxs-lookup"><span data-stu-id="c1848-156">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `helpUrl` | <span data-ttu-id="c1848-157">ユーザーがヘルプを要求したときに表示されるページの Url です。</span><span class="sxs-lookup"><span data-stu-id="c1848-157">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="c1848-158">関数が実行することについて説明します。</span><span class="sxs-lookup"><span data-stu-id="c1848-158">Describes what the function does.</span></span> <span data-ttu-id="c1848-159">この値は、関数が Excel 内のオートコンプリート メニューで選択された項目となっている場合に、ツールヒントとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="c1848-159">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="c1848-160">関数によって返される情報の種類を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c1848-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="c1848-161">子プロパティには、**文字列**、**数値**、または**ブール値**を使用できます。`type`</span><span class="sxs-lookup"><span data-stu-id="c1848-161">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="c1848-162">`dimensionality` 子プロパティの値には、**スカラー**または**マトリックス** (指定された `type` の値の 2 次元配列) が使用できます。</span><span class="sxs-lookup"><span data-stu-id="c1848-162">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `parameters` | <span data-ttu-id="c1848-163">関数の入力パラメーターを定義する配列。</span><span class="sxs-lookup"><span data-stu-id="c1848-163">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="c1848-164">`name` および `description` 子プロパティが Excel intelliSense に表示されます。</span><span class="sxs-lookup"><span data-stu-id="c1848-164">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="c1848-165">および `dimensionality` 子プロパティは、この表で前述した `result` オブジェクトの子プロパティと同じです。`type`</span><span class="sxs-lookup"><span data-stu-id="c1848-165">The `type` and `dimensionality` child properties are identical to the child properties of the `result` object that is described previously in this table.</span></span> |
| `options` | <span data-ttu-id="c1848-166">Excel がいつどのように関数を実行するのかについて、いくつかの機能をカスタマイズできるようになります。</span><span class="sxs-lookup"><span data-stu-id="c1848-166">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="c1848-167">このプロパティの使用方法の詳細については、この記事で後述する「[ストリーム関数](#streamed-functions)」および「[キャンセル](#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c1848-167">For more information about how this property can be used, see [Streamed functions](#streamed-functions) and [Cancellation](#canceling-a-function) later in this article.</span></span> |

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="c1848-168">外部ソースからデータを返す関数</span><span class="sxs-lookup"><span data-stu-id="c1848-168">Functions that return data from external sources</span></span>

<span data-ttu-id="c1848-169">カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1848-169">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="c1848-170">JavaScript Promise を Excel に返します。</span><span class="sxs-lookup"><span data-stu-id="c1848-170">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="c1848-171">コールバック関数を使用して Promise を最終値で解決します。</span><span class="sxs-lookup"><span data-stu-id="c1848-171">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="c1848-172">カスタム関数は、Excel が最終結果を待つ間、セルに `#GETTING_DATA` の一時的な結果を表示します。</span><span class="sxs-lookup"><span data-stu-id="c1848-172">Asynchronous functions display a `#GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="c1848-173">ユーザーは、カスタム関数が結果を待つ間、ワークシートの他の部分を通常通り操作することができます。</span><span class="sxs-lookup"><span data-stu-id="c1848-173">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

<span data-ttu-id="c1848-174">以下のコード サンプルでは、`getTemperature()` カスタム関数が温度計の現在の温度を取得します。</span><span class="sxs-lookup"><span data-stu-id="c1848-174">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="c1848-175">`sendWebRequest` は XHR を使用して温度 Web サービスを呼び出す仮想関数 (ここでは説明していません) であることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c1848-175">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a><span data-ttu-id="c1848-176">ストリーム関数</span><span class="sxs-lookup"><span data-stu-id="c1848-176">Streamed functions</span></span>

<span data-ttu-id="c1848-177">ストリーム カスタム関数を使用すると、時間の経過とともにセルに繰り返しデータを出力でき、ユーザーが再計算を要求することは特に必要ありません。</span><span class="sxs-lookup"><span data-stu-id="c1848-177">Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request recalculation.</span></span> <span data-ttu-id="c1848-178">以下のコード サンプルは、1 秒おきに結果に数値を追加するカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="c1848-178">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="c1848-179">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="c1848-179">Note the following about this code:</span></span>

- <span data-ttu-id="c1848-180">Excel は、`setResult`コールバックを使用して自動的に新しい値を表示します。</span><span class="sxs-lookup"><span data-stu-id="c1848-180">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="c1848-181">最終的なパラメーター `handler` は登録コードでは指定されず、Excel ユーザーが関数を入力するときにオートコンプリート メニューに表示されません。</span><span class="sxs-lookup"><span data-stu-id="c1848-181">For streamed functions, the final parameter, `handler`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="c1848-182">これは、関数のデータを Excel に渡してセルの値を更新するために使用される `setResult` コールバック関数を含むオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c1848-182">It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.</span></span>

- <span data-ttu-id="c1848-183">Excel が `handler` オブジェクトの `setResult` 関数を渡すには、関数の登録の際に、JSON メタデータ ファイル内のカスタム関数の `options` プロパティでオプション `"stream": true` を設定して、ストリーミングへのサポートを宣言する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1848-183">In order for Excel to pass the `setResult` function in the `handler` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="c1848-184">関数をキャンセルする</span><span class="sxs-lookup"><span data-stu-id="c1848-184">Canceling a function</span></span>

<span data-ttu-id="c1848-185">状況によっては、帯域幅の消費量、作業メモリ、UPC への負荷を減らすために、ストリーム カスタム関数の実行をキャンセルする必要がある場合もあります。</span><span class="sxs-lookup"><span data-stu-id="c1848-185">In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="c1848-186">Excel は、以下のような状況では関数の実行をキャンセルします。</span><span class="sxs-lookup"><span data-stu-id="c1848-186">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="c1848-187">ユーザーが、関数への参照があるセルを編集または削除した場合。</span><span class="sxs-lookup"><span data-stu-id="c1848-187">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="c1848-188">関数の引数 (入力) のいずれかが変更された場合。</span><span class="sxs-lookup"><span data-stu-id="c1848-188">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="c1848-189">この場合、キャンセルに続いて新しい関数の呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="c1848-189">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="c1848-190">ユーザーが手動で再計算をトリガーする。</span><span class="sxs-lookup"><span data-stu-id="c1848-190">The user triggers recalculation manually.</span></span> <span data-ttu-id="c1848-191">この場合、キャンセルに続いて新しい関数の呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="c1848-191">In this case, a new function call is triggered in addition to the cancelation.</span></span>

> [!NOTE]
> <span data-ttu-id="c1848-192">すべてのストリーミング関数に対してキャンセル ハンドラを実装することが 必須 です。</span><span class="sxs-lookup"><span data-stu-id="c1848-192">You must implement a cancellation handler for every streaming function.</span></span>

<span data-ttu-id="c1848-193">関数をキャンセル可能にするには、JSON メタデータ ファイルのカスタム関数の `options` プロパティで、オプション `"cancelable": true` を設定してください。</span><span class="sxs-lookup"><span data-stu-id="c1848-193">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="c1848-194">以下のコードは、前述したのと同じ `incrementValue` 関数を示していますが、今回はキャンセル ハンドラが実装されています。</span><span class="sxs-lookup"><span data-stu-id="c1848-194">The following code shows the same `incrementValue` function that was described previously, but this time with a cancellation handler implemented.</span></span> <span data-ttu-id="c1848-195">この例では、`incrementValue` 関数がキャンセルされたときに `clearInterval()` が実行されます。</span><span class="sxs-lookup"><span data-stu-id="c1848-195">In this example, `clearInterval()` will run when the `incrementValue` function is canceled.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);

    handler.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="c1848-196">状態の保存と共有</span><span class="sxs-lookup"><span data-stu-id="c1848-196">Saving and sharing state</span></span>

<span data-ttu-id="c1848-197">カスタム関数では、JavaScript のグローバル変数にデータを保存できます。</span><span class="sxs-lookup"><span data-stu-id="c1848-197">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="c1848-198">後続の呼び出しでは、カスタム関数はこれらの変数に保存されている値を使用できます。</span><span class="sxs-lookup"><span data-stu-id="c1848-198">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="c1848-199">保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を追加する場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="c1848-199">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="c1848-200">たとえば、同じ Web リソースへの追加呼び出しを避けるために、呼び出しから返されたデータを Web リソースに保存することができます。</span><span class="sxs-lookup"><span data-stu-id="c1848-200">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="c1848-201">次のコード サンプルは、 状態をグローバルで保存する前述の温度ストリーミング関数の実装を示しています。</span><span class="sxs-lookup"><span data-stu-id="c1848-201">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="c1848-202">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="c1848-202">Note the following about this code:</span></span>

- <span data-ttu-id="c1848-203">`refreshTemperature` は、1 秒おきに特定の温度計の温度を読み取るストリーム関数です。</span><span class="sxs-lookup"><span data-stu-id="c1848-203">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="c1848-204">新しい温度は `savedTemperatures` 変数に保存されますが、セルの値を直接更新することはありません。</span><span class="sxs-lookup"><span data-stu-id="c1848-204">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="c1848-205">ワークシート・セルから直接呼び出されません。\*したがって、JSON ファイルには登録されません \*</span><span class="sxs-lookup"><span data-stu-id="c1848-205">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>

- <span data-ttu-id="c1848-206">`streamTemperature` 1 秒おきにセルに表示される温度値を更新します。また、 `savedTemperatures` 変数をデータソースとして使用します。</span><span class="sxs-lookup"><span data-stu-id="c1848-206">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="c1848-207">JSON ファイルに登録し、すべて大文字で `STREAMTEMPERATURE` という名前をつける必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1848-207">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>

- <span data-ttu-id="c1848-208">ユーザーは、Excel UI の複数のセルから `streamTemperature` を呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="c1848-208">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="c1848-209">呼び出すたびに、同じ `savedTemperatures` 変数からデータを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="c1848-209">Each call reads data from the same `savedTemperatures` variable.</span></span>

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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="c1848-210">データの範囲を使用する</span><span class="sxs-lookup"><span data-stu-id="c1848-210">Working with ranges of data</span></span>

<span data-ttu-id="c1848-211">カスタム関数は、入力パラメーターとしてデータの範囲を受け取ることができます。または、データの範囲を返すこともできます。</span><span class="sxs-lookup"><span data-stu-id="c1848-211">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="c1848-212">JavaScript では、データの範囲は、2 次元配列として表されます。</span><span class="sxs-lookup"><span data-stu-id="c1848-212">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="c1848-213">たとえば、関数が Excel に格納されている数値の範囲から 2 番目に高い値を返すとします。</span><span class="sxs-lookup"><span data-stu-id="c1848-213">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="c1848-214">以下の関数が、タイプ `Excel.CustomFunctionDimensionality.matrix` のものである `values` パラメーターを受け取ります。</span><span class="sxs-lookup"><span data-stu-id="c1848-214">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="c1848-215">この関数の JSON メタデータでは、パラメーターの `type` プロパティを `matrix` に設定するように注意してください。</span><span class="sxs-lookup"><span data-stu-id="c1848-215">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="c1848-216">エラーの処理</span><span class="sxs-lookup"><span data-stu-id="c1848-216">Handling errors</span></span>

<span data-ttu-id="c1848-217">カスタム関数を定義するアドインをビルドする場合には、実行時エラーに対処するエラー処理ロジックを含めるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="c1848-217">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="c1848-218">カスタム関数のエラー処理は、[一般的な Excel JavaScript API のエラー処理](excel-add-ins-error-handling.md) と同じです。</span><span class="sxs-lookup"><span data-stu-id="c1848-218">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="c1848-219">以下のコード サンプルでは、`.catch` がコード内で発生するエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="c1848-219">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
    let url = "https://yourhypotheticalapi/comments/" + x;

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

## <a name="known-issues"></a><span data-ttu-id="c1848-220">既知の問題</span><span class="sxs-lookup"><span data-stu-id="c1848-220">Known issues</span></span>

- <span data-ttu-id="c1848-221">ヘルプの URL とパラメーターの説明。Excel ではまだ使用されていません。</span><span class="sxs-lookup"><span data-stu-id="c1848-221">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="c1848-222">カスタム機能は現在、モバイル クライアント用の Excel では使用できません。</span><span class="sxs-lookup"><span data-stu-id="c1848-222">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="c1848-223">揮発性関数（スプレッドシート内の無関係なデータが変更されたときに自動的に再計算する関数）はまだサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c1848-223">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="c1848-224">Office 365 管理ポータルと AppSource による展開はまだ有効になっていません。</span><span class="sxs-lookup"><span data-stu-id="c1848-224">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="c1848-225">Excel Online のカスタム関数は、一定期間使用しないとセッション中に機能しなくなることがあります。</span><span class="sxs-lookup"><span data-stu-id="c1848-225">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="c1848-226">ブラウザページを更新（F5）し、カスタム関数を再入力して機能を復元します。</span><span class="sxs-lookup"><span data-stu-id="c1848-226">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="c1848-227">Excel for Windows で実行されている複数のアドインがある場合には、ワークシートのセル内に **#GETTING_DATA** の一時的な結果が表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="c1848-227">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="c1848-228">すべての Excel ウィンドウを閉じて、Excel を再起動します。</span><span class="sxs-lookup"><span data-stu-id="c1848-228">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="c1848-229">将来的には、カスタム関数用のデバッグ ツールが利用可能となる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="c1848-229">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="c1848-230">それまでは、F12 開発者ツールを使用して Excel オンラインでデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="c1848-230">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="c1848-231">詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c1848-231">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="c1848-232">変更ログ</span><span class="sxs-lookup"><span data-stu-id="c1848-232">Changelog</span></span>

- <span data-ttu-id="c1848-233">**2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開\*</span><span class="sxs-lookup"><span data-stu-id="c1848-233">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="c1848-234">**2017 年 11 月 20 日**: ビルド 8801 以降を使用する場合の互換性バグを修正</span><span class="sxs-lookup"><span data-stu-id="c1848-234">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="c1848-235">**2017 年 11 月 28 日**: 非同期関数のキャンセルへのサポートを公開\* (ストリーム関数への変更が必要)</span><span class="sxs-lookup"><span data-stu-id="c1848-235">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="c1848-236">**2018 年 5 月 7 日**: Mac、Excel Online、およびインプロセスで実行される同期関数へのサポートを公開\*</span><span class="sxs-lookup"><span data-stu-id="c1848-236">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="c1848-237">**2018 年 9 月 20日**: JavaScript 実行時のカスタム関数へのサポートを公開</span><span class="sxs-lookup"><span data-stu-id="c1848-237">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="c1848-238">詳細については、「[Excel カスタム関数のランタイム](custom-functions-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c1848-238">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="c1848-239">\* Office Insiders チャネル対象</span><span class="sxs-lookup"><span data-stu-id="c1848-239">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="c1848-240">関連項目</span><span class="sxs-lookup"><span data-stu-id="c1848-240">See also</span></span>

* [<span data-ttu-id="c1848-241">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="c1848-241">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="c1848-242">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="c1848-242">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="c1848-243">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="c1848-243">Custom functions best practices</span></span>](custom-functions-best-practices.md)
