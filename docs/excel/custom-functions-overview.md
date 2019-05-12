---
ms.date: 05/03/2019
description: JavaScript を使用して Excel でカスタム関数を作成する。
title: Excel でカスタム関数を作成する
localization_priority: Priority
ms.openlocfilehash: 5a31cc8ddfe98b880ab09803c7c0b7b615ba85db
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659651"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="fabf4-103">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="fabf4-103">Create custom functions in Excel</span></span> 

<span data-ttu-id="fabf4-104">開発者は、カスタム関数を使用して関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="fabf4-105">ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="fabf4-106">この記事では、Excel でカスタム関数を作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="fabf4-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="fabf4-107">次のアニメーション画像は、JavaScript または Typescript で作成した関数を呼び出すブックを示しています。</span><span class="sxs-lookup"><span data-stu-id="fabf4-107">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="fabf4-108">この例では、カスタム関数 `=MYFUNCTION.SPHEREVOLUME` は球の体積を計算します。</span><span class="sxs-lookup"><span data-stu-id="fabf4-108">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolume.gif" />

<span data-ttu-id="fabf4-109">`=MYFUNCTION.SPHEREVOLUME` カスタム関数は次のコードにより定義されます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-109">The following code defines the `=MYFUNCTION.SPHEREVOLUME` custom function.</span></span>

```js
/**
 * Returns the volume of a sphere. 
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
CustomFunctions.associate("SPHEREVOLUME", sphereVolume)
```

> [!NOTE]
> <span data-ttu-id="fabf4-110">この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。</span><span class="sxs-lookup"><span data-stu-id="fabf4-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="fabf4-111">コードでカスタム関数を定義する方法</span><span class="sxs-lookup"><span data-stu-id="fabf4-111">How a custom function is defined in code</span></span>

<span data-ttu-id="fabf4-112">[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Excel のカスタム関数アドイン プロジェクトを作成する場合、使用する関数、作業ウィンドウ、およびアドイン全体をこのジェネレーターが作成します。</span><span class="sxs-lookup"><span data-stu-id="fabf4-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall.</span></span> <span data-ttu-id="fabf4-113">このため、カスタム関数に重要なファイルに注意を集中できます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-113">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="fabf4-114">ファイル</span><span class="sxs-lookup"><span data-stu-id="fabf4-114">File</span></span> | <span data-ttu-id="fabf4-115">ファイル形式</span><span class="sxs-lookup"><span data-stu-id="fabf4-115">File format</span></span> | <span data-ttu-id="fabf4-116">説明</span><span class="sxs-lookup"><span data-stu-id="fabf4-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="fabf4-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="fabf4-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="fabf4-118">または</span><span class="sxs-lookup"><span data-stu-id="fabf4-118">or</span></span><br/><span data-ttu-id="fabf4-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="fabf4-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="fabf4-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="fabf4-120">JavaScript</span></span><br/><span data-ttu-id="fabf4-121">または</span><span class="sxs-lookup"><span data-stu-id="fabf4-121">or</span></span><br/><span data-ttu-id="fabf4-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="fabf4-122">TypeScript</span></span> | <span data-ttu-id="fabf4-123">カスタム関数を定義するコードが含みます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="fabf4-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="fabf4-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="fabf4-125">HTML</span><span class="sxs-lookup"><span data-stu-id="fabf4-125">HTML</span></span> | <span data-ttu-id="fabf4-126">カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。</span><span class="sxs-lookup"><span data-stu-id="fabf4-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="fabf4-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="fabf4-127">**./manifest.xml**</span></span> | <span data-ttu-id="fabf4-128">XML</span><span class="sxs-lookup"><span data-stu-id="fabf4-128">XML</span></span> | <span data-ttu-id="fabf4-129">アドイン内のすべてのカスタム関数の名前空間と、この表で前述した JavaScript ファイルと HTML ファイルの位置を指定します。</span><span class="sxs-lookup"><span data-stu-id="fabf4-129">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript and HTML files that are listed previously in this table.</span></span> <span data-ttu-id="fabf4-130">また、作業ウィンドウ ファイルやコマンド ファイルなど、アドインで使用する可能性のある他のファイルの位置もリストされます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-130">It also lists the locations of other files your add-in might make use of, such as the task pane files and command files.</span></span> |

### <a name="script-file"></a><span data-ttu-id="fabf4-131">スクリプト ファイル</span><span class="sxs-lookup"><span data-stu-id="fabf4-131">Script file</span></span>

<span data-ttu-id="fabf4-132">スクリプト ファイル (**./src/customfunctions.js** または **/src/customfunctions.ts**) は、カスタム関数を定義し、どのコードがその関数を定義するかをコメントし、カスタム関数の名前を JSON メタデータ ファイルのオブジェクトに関連付けるコードを格納しています。</span><span class="sxs-lookup"><span data-stu-id="fabf4-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions, comments which define the function, and associates the names of the custom functions to objects in the JSON metadata file.</span></span>

<span data-ttu-id="fabf4-133">`add` カスタム関数は次のコードにより定義されます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-133">The following code defines the `add` custom function.</span></span> <span data-ttu-id="fabf4-134">コード コメントは、Excel にカスタム関数を記述する JSON メタデータ ファイルを生成するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-134">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="fabf4-135">必須の `@customfunction` コメントが最初に宣言されて、これがカスタム関数であることを示します。</span><span class="sxs-lookup"><span data-stu-id="fabf4-135">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="fabf4-136">さらに、お気付きのように `first` と `second` の 2 つのパラメーターが宣言されており、その後にそれらの `description` プロパティが記述されます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-136">Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties.</span></span> <span data-ttu-id="fabf4-137">最後に `returns` の説明が記述されます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-137">Finally, a `returns` description is given.</span></span> <span data-ttu-id="fabf4-138">カスタム関数で必要になるコメントに関する詳細については、「[カスタム関数の JSON メタデータを作成する](custom-functions-json-autogeneration.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fabf4-138">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="fabf4-139">次のコードでは、`CustomFunctions.associate("ADD", add)` も呼び出して、`add()` 関数を JSON メタデータ ファイル `ADD` の ID と関連付けます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-139">The following code also calls `CustomFunctions.associate("ADD", add)` to associate the function `add()` with its ID in the JSON metadata file `ADD`.</span></span> <span data-ttu-id="fabf4-140">関数の関連付けに関する詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md#associating-function-names-with-json-metadata)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fabf4-140">For more information on associating functions, see [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span></span>

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="fabf4-141">カスタム関数のランタイムの読み込みを制御する **functions.html** ファイルは、カスタム関数の現在の CDN にリンクしていなければならないことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="fabf4-141">Note that the **functions.html** file, which governs the loading of the custom functions runtime, must link to the current CDN for custom functions.</span></span> <span data-ttu-id="fabf4-142">最新バージョンの Yo Office ジェネレーターを使用して作成されたプロジェクトは、正しい CDN を参照します。</span><span class="sxs-lookup"><span data-stu-id="fabf4-142">Projects prepared with the current version of the Yo Office generator reference the correct CDN.</span></span> <span data-ttu-id="fabf4-143">2019 年 3 月以前の古いカスタム関数のプロジェクトを改良する場合は、以下のコードを **functions.html** ページにコピーする必要があります。</span><span class="sxs-lookup"><span data-stu-id="fabf4-143">If you are retrofitting a previous custom function project from March 2019 or earlier, you need to copy in the code below to the **functions.html** page.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### <a name="manifest-file"></a><span data-ttu-id="fabf4-144">マニフェスト ファイル</span><span class="sxs-lookup"><span data-stu-id="fabf4-144">Manifest file</span></span>

<span data-ttu-id="fabf4-145">カスタム関数 (Yo Office ジェネレーターが作成するプロジェクトでは **./manifest.xml**) を定義するアドインの XML マニフェスト ファイルは、アドイン内のすべてのカスタム関数の名前空間と、 JavaScript、JSON、および HTML の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="fabf4-145">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> 

<span data-ttu-id="fabf4-146">次の基本的な XML マークアップは、カスタム関数を有効にするアドインのマニフェストに含める必要がある要素`<ExtensionPoint>` と `<Resources>` の例を示しています。</span><span class="sxs-lookup"><span data-stu-id="fabf4-146">The following basic XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span> <span data-ttu-id="fabf4-147">Yo Office ジェネレーターを使用する場合、生成されたカスタム関数ファイルには、さらに複雑なマニフェスト ファイルが格納されます。こちらの[Github リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml)で比較できます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-147">If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span></span>

> [!NOTE] 
> <span data-ttu-id="fabf4-148">カスタム関数のJavaScript、JSON、HTML ファイルのマニフェスト ファイルで指定した URL はだれでもアクセスでき、同じサブドメインを持つ必要があります。</span><span class="sxs-lookup"><span data-stu-id="fabf4-148">The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.</span></span>

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
> <span data-ttu-id="fabf4-149">Excel の関数は、XML マニフェスト ファイルで指定された名前空間が接頭辞として付加されます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-149">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="fabf4-150">関数の名前空間は、関数名の前に付けられ、ピリオドで区切られます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-150">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="fabf4-151">例えば、Excel ワークシートのセル内で、`ADD42` 関数を呼び出すためには、`=CONTOSO.ADD42` と入力します。これは、`CONTOSO` が名前空間で、`ADD42` が JSON ファイルで指定された関数の名前だからです。</span><span class="sxs-lookup"><span data-stu-id="fabf4-151">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="fabf4-152">名前空間は、会社またはアドインの識別子としての使用を目的としています。</span><span class="sxs-lookup"><span data-stu-id="fabf4-152">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="fabf4-153">名前空間にはアルファベットとピリオドのみを含めることが出来ます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-153">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="fabf4-154">共同編集</span><span class="sxs-lookup"><span data-stu-id="fabf4-154">Coauthoring</span></span>

<span data-ttu-id="fabf4-155">Excel Online と Excel for Windows で Office 365 サブスクリプションを利用している場合、ドキュメントの共同編集を行うことができ、カスタム関数を使用できます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-155">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="fabf4-156">ブックでカスタム関数を使用している場合、仕事仲間はカスタム関数のアドインを読み込むように要求されます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-156">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="fabf4-157">双方がアドインを読み込むと、共同編集によりカスタム関数は結果を共有します。</span><span class="sxs-lookup"><span data-stu-id="fabf4-157">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="fabf4-158">共同編集の詳細については、「[Excel での共同編集](/office/vba/excel/concepts/about-coauthoring-in-excel)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fabf4-158">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="fabf4-159">既知の問題</span><span class="sxs-lookup"><span data-stu-id="fabf4-159">Known issues</span></span>

<span data-ttu-id="fabf4-160">既知の問題については、[Excel カスタム関数についての GitHub のレポート](https://github.com/OfficeDev/Excel-Custom-Functions/issues)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fabf4-160">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="fabf4-161">次の手順</span><span class="sxs-lookup"><span data-stu-id="fabf4-161">Next steps</span></span>

<span data-ttu-id="fabf4-162">カスタム関数を試してみましょう。</span><span class="sxs-lookup"><span data-stu-id="fabf4-162">Want to try out custom functions?</span></span> <span data-ttu-id="fabf4-163">もしまだであれば、簡単な[カスタム関数クイックスタート](../quickstarts/excel-custom-functions-quickstart.md)または、詳細な[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="fabf4-163">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span> 

<span data-ttu-id="fabf4-164">独自のカスタム関数を試すもう 1 つの簡単な方法は[スクリプト ラボ](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)を使用し、アドインで Excel のカスタム関数を試してみることができます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-164">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="fabf4-165">独自のカスタム関数を作成したり、提供されたサンプルを再生してみることができます。</span><span class="sxs-lookup"><span data-stu-id="fabf4-165">You can try out creating your own custom function or play with the provided samples.</span></span>

<span data-ttu-id="fabf4-166">カスタム関数の機能の詳細について読む準備はできましたか?</span><span class="sxs-lookup"><span data-stu-id="fabf4-166">Ready to read more about the capabilities custom functions?</span></span> <span data-ttu-id="fabf4-167">[カスタム関数のアーキテクチャ](custom-functions-architecture.md)の概要をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="fabf4-167">Learn about an overview of [the custom functions architecture](custom-functions-architecture.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="fabf4-168">関連項目</span><span class="sxs-lookup"><span data-stu-id="fabf4-168">See also</span></span> 
* [<span data-ttu-id="fabf4-169">カスタム関数の要件</span><span class="sxs-lookup"><span data-stu-id="fabf4-169">Custom functions requirements</span></span>](custom-functions-requirements.md)
* [<span data-ttu-id="fabf4-170">名前付けのガイドライン</span><span class="sxs-lookup"><span data-stu-id="fabf4-170">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="fabf4-171">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="fabf4-171">Best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="fabf4-172">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="fabf4-172">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
