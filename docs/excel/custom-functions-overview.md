---
ms.date: 03/29/2019
description: JavaScript を使用して Excel でカスタム関数を作成する。
title: Excel でのカスタム関数の作成 (プレビュー)
localization_priority: Priority
ms.openlocfilehash: 7a461728061ace532a11a8473d27ec4340eebb97
ms.sourcegitcommit: fbe2a799fda71aab73ff1c5546c936edbac14e47
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/10/2019
ms.locfileid: "31764412"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="6031e-103">Excel でのカスタム関数の作成 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="6031e-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="6031e-104">開発者は、カスタム関数を使用して関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="6031e-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="6031e-105">ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="6031e-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="6031e-106">この記事では、Excel でカスタム関数を作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="6031e-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="6031e-107">次の図は、エンドユーザーが Excel ワークシートのセルにカスタム関数を挿入する様子を示します。</span><span class="sxs-lookup"><span data-stu-id="6031e-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="6031e-108">`CONTOSO.ADD42` カスタム関数は、関数への入力パラメーターとしてユーザーが指定した数値のペアに 42 を追加するように設計されています。</span><span class="sxs-lookup"><span data-stu-id="6031e-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="6031e-109">`ADD42` カスタム関数は次のコードにより定義されます。</span><span class="sxs-lookup"><span data-stu-id="6031e-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="6031e-110">この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。</span><span class="sxs-lookup"><span data-stu-id="6031e-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="6031e-111">カスタム関数 アドイン プロジェクトのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="6031e-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="6031e-112">[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Excel のカスタム関数アドイン プロジェクトを作成する場合、使用する関数、作業ウィンドウ、およびアドイン全体をこのジェネレーターが作成します。</span><span class="sxs-lookup"><span data-stu-id="6031e-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall.</span></span> <span data-ttu-id="6031e-113">このため、カスタム関数に重要なファイルに注意を集中できます。</span><span class="sxs-lookup"><span data-stu-id="6031e-113">We'll concentrate on the files that are important to custom functions:</span></span> 

| <span data-ttu-id="6031e-114">ファイル</span><span class="sxs-lookup"><span data-stu-id="6031e-114">File</span></span> | <span data-ttu-id="6031e-115">ファイル形式</span><span class="sxs-lookup"><span data-stu-id="6031e-115">File format</span></span> | <span data-ttu-id="6031e-116">説明</span><span class="sxs-lookup"><span data-stu-id="6031e-116">Description</span></span> |
|------|-------------|-------------|
| **<span data-ttu-id="6031e-117">./src/functions/functions.js</span><span class="sxs-lookup"><span data-stu-id="6031e-117">./src/functions/functions.js</span></span>**<br/><span data-ttu-id="6031e-118">または</span><span class="sxs-lookup"><span data-stu-id="6031e-118">or</span></span><br/>**<span data-ttu-id="6031e-119">./src/functions/functions.ts</span><span class="sxs-lookup"><span data-stu-id="6031e-119">./src/functions/functions.ts</span></span>** | <span data-ttu-id="6031e-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="6031e-120">JavaScript</span></span><br/><span data-ttu-id="6031e-121">または</span><span class="sxs-lookup"><span data-stu-id="6031e-121">or</span></span><br/><span data-ttu-id="6031e-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="6031e-122">TypeScript</span></span> | <span data-ttu-id="6031e-123">カスタム関数を定義するコードが含みます。</span><span class="sxs-lookup"><span data-stu-id="6031e-123">Contains the code that defines custom functions.</span></span> |
| **<span data-ttu-id="6031e-124">./src/functions/functions.html</span><span class="sxs-lookup"><span data-stu-id="6031e-124">./src/functions/functions.html</span></span>** | <span data-ttu-id="6031e-125">HTML</span><span class="sxs-lookup"><span data-stu-id="6031e-125">HTML</span></span> | <span data-ttu-id="6031e-126">カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。</span><span class="sxs-lookup"><span data-stu-id="6031e-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| **<span data-ttu-id="6031e-127">./manifest.xml</span><span class="sxs-lookup"><span data-stu-id="6031e-127">./manifest.xml</span></span>** | <span data-ttu-id="6031e-128">XML</span><span class="sxs-lookup"><span data-stu-id="6031e-128">XML</span></span> | <span data-ttu-id="6031e-129">アドイン内のすべてのカスタム関数の名前空間と、この表で前述した JavaScript ファイルと HTML ファイルの位置を指定します。</span><span class="sxs-lookup"><span data-stu-id="6031e-129">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> <span data-ttu-id="6031e-130">また、作業ウィンドウ ファイルやコマンド ファイルなど、アドインで使用する可能性のある他のファイルの位置もリストされます。</span><span class="sxs-lookup"><span data-stu-id="6031e-130">It also lists the locations of other files your add-in might make use of, such as the task pane files and command files.</span></span> |

### <a name="script-file"></a><span data-ttu-id="6031e-131">スクリプト ファイル</span><span class="sxs-lookup"><span data-stu-id="6031e-131">Script file</span></span>

<span data-ttu-id="6031e-132">スクリプト ファイル (Yo Office ジェネレーターが作成するプロジェクト内の **./src/customfunctions.js** または **/src/customfunctions.ts**) は、カスタム関数を定義し、どのコードがその関数を定義するかをコメントし、カスタム関数の名前を JSON メタデータ ファイルのオブジェクトに関連付けるコードを格納しています。</span><span class="sxs-lookup"><span data-stu-id="6031e-132">The script file (./src/customfunctions.js or ./src/customfunctions.ts in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the JSON metadata file.</span></span>

<span data-ttu-id="6031e-133">次のコードはカスタム関数 `add` を定義し、その関数の関連付け情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="6031e-133">The following code defines the custom function `add`  and then specifies association information for the function.</span></span> <span data-ttu-id="6031e-134">関数の関連付けに関する詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#associating-function-names-with-json-metadata)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6031e-134">For more information on associating functions, see [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span></span>

<span data-ttu-id="6031e-135">次のコードも、関数を定義するコード コメントを提供します。</span><span class="sxs-lookup"><span data-stu-id="6031e-135">The following code also provides code comments which define the function.</span></span> <span data-ttu-id="6031e-136">必須の `@customfunction` コメントが最初に宣言されて、これがカスタム関数であることを示します。</span><span class="sxs-lookup"><span data-stu-id="6031e-136">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="6031e-137">さらに、お気付きのように `first` と `second` の 2 つのパラメーターが宣言されており、その後にそれらの `description` プロパティが記述されます。</span><span class="sxs-lookup"><span data-stu-id="6031e-137">Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties.</span></span> <span data-ttu-id="6031e-138">最後に `returns` の説明が記述されます。</span><span class="sxs-lookup"><span data-stu-id="6031e-138">Finally, a `returns` description is given.</span></span> <span data-ttu-id="6031e-139">カスタム関数で必要になるコメントに関する詳細については、「[カスタム関数の JSON メタデータを作成する](custom-functions-json-autogeneration.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6031e-139">For more information about what comments are required for your custom function, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

### <a name="manifest-file"></a><span data-ttu-id="6031e-140">マニフェスト ファイル</span><span class="sxs-lookup"><span data-stu-id="6031e-140">Manifest file</span></span>

<span data-ttu-id="6031e-141">カスタム関数 (Yo Office ジェネレーターが作成するプロジェクトでは **./manifest.xml**) を定義するアドインの XML マニフェスト ファイルは、アドイン内のすべてのカスタム関数の名前空間と、 JavaScript、JSON、および HTML の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="6031e-141">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> 

<span data-ttu-id="6031e-142">次の基本的な XML マークアップは、カスタム関数を有効にするアドインのマニフェストに含める必要がある要素`<ExtensionPoint>` と `<Resources>` の例を示しています。</span><span class="sxs-lookup"><span data-stu-id="6031e-142">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span> <span data-ttu-id="6031e-143">Yo Office ジェネレーターを使用する場合、生成されたカスタム関数ファイルには、さらに複雑なマニフェスト ファイルが格納されます。こちらの[Github リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml)で比較できます。</span><span class="sxs-lookup"><span data-stu-id="6031e-143">If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span></span>

> [!NOTE] 
> <span data-ttu-id="6031e-144">カスタム関数のJavaScript、JSON、HTML ファイルのマニフェスト ファイルで指定した URL はだれでもアクセスでき、同じサブドメインを持つ必要があります。</span><span class="sxs-lookup"><span data-stu-id="6031e-144">The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.</span></span>

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
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="6031e-145">Excel の関数は、XML マニフェスト ファイルで指定された名前空間が接頭辞として付加されます。</span><span class="sxs-lookup"><span data-stu-id="6031e-145">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="6031e-146">関数の名前空間は、関数名の前に付けられ、ピリオドで区切られます。</span><span class="sxs-lookup"><span data-stu-id="6031e-146">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="6031e-147">例えば、Excel ワークシートのセル内で、`ADD42` 関数を呼び出すためには、`=CONTOSO.ADD42` と入力します。これは、`CONTOSO` が名前空間で、`ADD42` が JSON ファイルで指定された関数の名前だからです。</span><span class="sxs-lookup"><span data-stu-id="6031e-147">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="6031e-148">名前空間は、会社またはアドインの識別子としての使用を目的としています。</span><span class="sxs-lookup"><span data-stu-id="6031e-148">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="6031e-149">名前空間にはアルファベットとピリオドのみを含めることが出来ます。</span><span class="sxs-lookup"><span data-stu-id="6031e-149">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="declaring-a-volatile-function"></a><span data-ttu-id="6031e-150">揮発性関数の宣言</span><span class="sxs-lookup"><span data-stu-id="6031e-150">Declaring a volatile function</span></span>

<span data-ttu-id="6031e-151">[揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)とは、関数のいずれの引数にも変更がない場合でも、値が刻々と変化する関数のことです。</span><span class="sxs-lookup"><span data-stu-id="6031e-151">[Volatile functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed.</span></span> <span data-ttu-id="6031e-152">これらの関数は、Excel が再計算するたびに再計算を行います。</span><span class="sxs-lookup"><span data-stu-id="6031e-152">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="6031e-153">たとえば、`NOW` 関数を呼び出すセルがあるとします。</span><span class="sxs-lookup"><span data-stu-id="6031e-153">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="6031e-154">`NOW` が呼び出される度に、現在の日付と時刻を自動的に返します。</span><span class="sxs-lookup"><span data-stu-id="6031e-154">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="6031e-155">Excel には、`RAND` や `TODAY` などの組み込み揮発性関数がいくつか含まれています。</span><span class="sxs-lookup"><span data-stu-id="6031e-155">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="6031e-156">Excel の揮発性関数の完全なリストは、「[揮発性および非揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6031e-156">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="6031e-157">カスタム関数を使用すると独自の揮発性関数を作成することができ、日時、時間、乱数、およびモデルを処理するときに役立つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="6031e-157">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling.</span></span> <span data-ttu-id="6031e-158">たとえば、モンテカルロ シミュレーションでは、最適なソリューションを決定するにはランダムな入力値の生成が必要です。</span><span class="sxs-lookup"><span data-stu-id="6031e-158">For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="6031e-159">関数を揮発性であると宣言するには、次のコードで示されるように、JSON メタデータファイルの関数で、`options` オブジェクトに`"volatile": true` を追加します。</span><span class="sxs-lookup"><span data-stu-id="6031e-159">To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample.</span></span> <span data-ttu-id="6031e-160">関数で `"streaming": true`と`"volatile": true` の両方をマークすることはできません。両方とも `true` とマークされている場合、揮発性のオプションは無視されます。</span><span class="sxs-lookup"><span data-stu-id="6031e-160">Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.</span></span>

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

## <a name="saving-and-sharing-state"></a><span data-ttu-id="6031e-161">状態の保存と共有</span><span class="sxs-lookup"><span data-stu-id="6031e-161">Saving and sharing state</span></span>

<span data-ttu-id="6031e-162">カスタム関数は、グローバル JavaScript 変数にデータを保存でき、以降の呼び出しで使用することができます。</span><span class="sxs-lookup"><span data-stu-id="6031e-162">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="6031e-163">保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を呼び出す場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="6031e-163">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="6031e-164">たとえば、同じ Web リソースへの追加呼び出しを避けるために、呼び出しから返されたデータを Web リソースに保存することができます。</span><span class="sxs-lookup"><span data-stu-id="6031e-164">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="6031e-165">次のコード サンプルでは、状態をグローバルに保存する温度ストリーミング関数の実装を示します。</span><span class="sxs-lookup"><span data-stu-id="6031e-165">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="6031e-166">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="6031e-166">Note the following about this code:</span></span>

- <span data-ttu-id="6031e-167">`streamTemperature` 関数がセルに表示される温度の値を毎秒更新し、`savedTemperatures` 変数をデータ ソースとして使用します。</span><span class="sxs-lookup"><span data-stu-id="6031e-167">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="6031e-168">`streamTemperature` はストリーム関数であるため、その関数がキャンセルされたときに実行されるキャンセル ハンドラーを実装します。</span><span class="sxs-lookup"><span data-stu-id="6031e-168">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="6031e-169">ユーザーが `streamTemperature` 関数を Excel の複数のセルから呼び出す場合、`streamTemperature` 関数は実行のたびに、同じ `savedTemperatures` 変数からのデータを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="6031e-169">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="6031e-170">`refreshTemperature` 関数は、特定の温度計の温度を毎秒読み取り、結果を `savedTemperatures` 変数に格納します。</span><span class="sxs-lookup"><span data-stu-id="6031e-170">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="6031e-171">`refreshTemperature` 関数は、Excel でエンド ユーザーには公開されないので、JSON ファイルに登録する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="6031e-171">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="coauthoring"></a><span data-ttu-id="6031e-172">共同編集</span><span class="sxs-lookup"><span data-stu-id="6031e-172">Coauthoring</span></span>

<span data-ttu-id="6031e-173">Excel Online と Excel for Windows で Office 365 サブスクリプションを利用している場合、ドキュメントの共同編集を行うことができ、カスタム関数を使用できます。</span><span class="sxs-lookup"><span data-stu-id="6031e-173">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="6031e-174">ブックでカスタム関数を使用している場合、仕事仲間はカスタム関数のアドインを読み込むように要求されます。</span><span class="sxs-lookup"><span data-stu-id="6031e-174">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="6031e-175">双方がアドインを読み込むと、共同編集によりカスタム関数は結果を共有します。</span><span class="sxs-lookup"><span data-stu-id="6031e-175">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="6031e-176">共同編集の詳細については、「[Excel での共同編集](/office/vba/excel/concepts/about-coauthoring-in-excel)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6031e-176">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="6031e-177">データの範囲を使用する</span><span class="sxs-lookup"><span data-stu-id="6031e-177">Working with ranges of data</span></span>

<span data-ttu-id="6031e-178">カスタム関数は、データの範囲を入力パラメーターとして受け入れることができ、また、データの範囲を返すこともできます。</span><span class="sxs-lookup"><span data-stu-id="6031e-178">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="6031e-179">JavaScript では、データの範囲は 2 次元配列として表されます。</span><span class="sxs-lookup"><span data-stu-id="6031e-179">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="6031e-180">例えば、関数が Excel に保存されている数値の範囲から 2 番目に大きい値を返すとします。</span><span class="sxs-lookup"><span data-stu-id="6031e-180">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="6031e-181">次の関数は、`Excel.CustomFunctionDimensionality.matrix` 型の `values` パラメーターを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="6031e-181">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="6031e-182">なお、この関数の JSON メタデータでは、パラメーターの`type`プロパティを`matrix` と設定します。</span><span class="sxs-lookup"><span data-stu-id="6031e-182">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="determine-which-cell-invoked-your-custom-function"></a><span data-ttu-id="6031e-183">カスタム関数が呼び出したセルを特定する</span><span class="sxs-lookup"><span data-stu-id="6031e-183">Determine which cell invoked your custom function</span></span>

<span data-ttu-id="6031e-184">場合によっては、カスタム関数が呼び出したセルのアドレスを取得する必要が生じます。</span><span class="sxs-lookup"><span data-stu-id="6031e-184">In some cases you'll need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="6031e-185">これは、次の種類のシナリオで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="6031e-185">This may be useful in the following types of scenarios:</span></span>

- <span data-ttu-id="6031e-186">範囲の書式設定: [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data) で情報を格納するキーとしてセル アドレスを使用します。</span><span class="sxs-lookup"><span data-stu-id="6031e-186">Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="6031e-187">Excel で [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) を使用して`AsyncStorage` からキーを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="6031e-187">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="6031e-188">キャッシュされた値を表示させる: 関数がオフラインで使用される場合、`onCalculated` を使用して `AsyncStorage` に格納されているキャッシュされた値を表示します。</span><span class="sxs-lookup"><span data-stu-id="6031e-188">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="6031e-189">調整: セル アドレスを使用して元のセルを検出し、処理が発生している場所での調整を行えます。</span><span class="sxs-lookup"><span data-stu-id="6031e-189">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="6031e-190">セルのアドレスに関する情報は、関数の JSON メタデータ ファイルで `requiresAddress` が`true` とマークされている場合にのみ公開されます。</span><span class="sxs-lookup"><span data-stu-id="6031e-190">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="6031e-191">これの例を次のサンプルに示します。</span><span class="sxs-lookup"><span data-stu-id="6031e-191">The following sample gives an example of this:</span></span>

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

<span data-ttu-id="6031e-192">セルのアドレスを検索するために、スクリプト ファイル (**./src/functions/functions.js** または **./src/functions/functions.ts**) に `getAddress` 関数を追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6031e-192">In the script file (**./src/customfunctions.js** or **./src/customfunctions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="6031e-193">この関数は、次のサンプルで示される `parameter1` のようなパラメーターを受け取ることができます。</span><span class="sxs-lookup"><span data-stu-id="6031e-193">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="6031e-194">最後のパラメーターは常に `invocationContext` で、これはJSON メタデータ ファイルで `requiresAddress` が `true` とマークされているときに Excel が返すセルの位置が格納されているオブジェクトのことです。</span><span class="sxs-lookup"><span data-stu-id="6031e-194">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="6031e-195">既定では、`getAddress` 関数が返す値は次の形式に従います: `SheetName!CellNumber`。</span><span class="sxs-lookup"><span data-stu-id="6031e-195">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="6031e-196">たとえば、ある関数が Expenses という名前のシートのセル B2 から呼び出される場合の戻り値は `Expenses!B2` になります。</span><span class="sxs-lookup"><span data-stu-id="6031e-196">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="known-issues"></a><span data-ttu-id="6031e-197">既知の問題</span><span class="sxs-lookup"><span data-stu-id="6031e-197">Known issues</span></span>

<span data-ttu-id="6031e-198">既知の問題については、[Excel カスタム関数についての GitHub のレポート](https://github.com/OfficeDev/Excel-Custom-Functions/issues)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6031e-198">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="see-also"></a><span data-ttu-id="6031e-199">関連項目</span><span class="sxs-lookup"><span data-stu-id="6031e-199">See also</span></span>

* [<span data-ttu-id="6031e-200">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="6031e-200">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="6031e-201">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="6031e-201">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="6031e-202">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="6031e-202">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="6031e-203">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="6031e-203">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="6031e-204">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="6031e-204">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="6031e-205">カスタム関数のデバッグ</span><span class="sxs-lookup"><span data-stu-id="6031e-205">Custom functions debugging</span></span>](custom-functions-debugging.md)
