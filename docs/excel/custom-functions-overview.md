---
ms.date: 08/13/2020
description: Office アドインの Excel カスタム関数を作成します。
title: Excel でカスタム関数を作成する
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 731e8d99a36cfef7d125838c67efcdd7a77b4bb1
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819561"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="ee953-103">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="ee953-103">Create custom functions in Excel</span></span>

<span data-ttu-id="ee953-104">開発者は、カスタム関数を使用して関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="ee953-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="ee953-105">ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="ee953-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="ee953-106">次のアニメーション画像は、JavaScript または Typescript で作成した関数を呼び出すブックを示しています。</span><span class="sxs-lookup"><span data-stu-id="ee953-106">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="ee953-107">この例では、カスタム関数 `=MYFUNCTION.SPHEREVOLUME` は球の体積を計算します。</span><span class="sxs-lookup"><span data-stu-id="ee953-107">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="ee953-108">`=MYFUNCTION.SPHEREVOLUME` カスタム関数は次のコードにより定義されます。</span><span class="sxs-lookup"><span data-stu-id="ee953-108">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

```js
/**
 * Returns the volume of a sphere.
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
```

> [!NOTE]
> <span data-ttu-id="ee953-109">この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。</span><span class="sxs-lookup"><span data-stu-id="ee953-109">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="ee953-110">コードでカスタム関数を定義する方法</span><span class="sxs-lookup"><span data-stu-id="ee953-110">How a custom function is defined in code</span></span>

<span data-ttu-id="ee953-111">[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Excel のカスタム関数アドイン プロジェクトを作成する場合、使用する関数および作業ウィンドウを制御するファイルが作成されます。</span><span class="sxs-lookup"><span data-stu-id="ee953-111">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, it creates files which control your functions and task pane.</span></span> <span data-ttu-id="ee953-112">このため、カスタム関数に重要なファイルに注意を集中できます。</span><span class="sxs-lookup"><span data-stu-id="ee953-112">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="ee953-113">ファイル</span><span class="sxs-lookup"><span data-stu-id="ee953-113">File</span></span> | <span data-ttu-id="ee953-114">ファイル形式</span><span class="sxs-lookup"><span data-stu-id="ee953-114">File format</span></span> | <span data-ttu-id="ee953-115">説明</span><span class="sxs-lookup"><span data-stu-id="ee953-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="ee953-116">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="ee953-116">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="ee953-117">または</span><span class="sxs-lookup"><span data-stu-id="ee953-117">or</span></span><br/><span data-ttu-id="ee953-118">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="ee953-118">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="ee953-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="ee953-119">JavaScript</span></span><br/><span data-ttu-id="ee953-120">または</span><span class="sxs-lookup"><span data-stu-id="ee953-120">or</span></span><br/><span data-ttu-id="ee953-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="ee953-121">TypeScript</span></span> | <span data-ttu-id="ee953-122">カスタム関数を定義するコードが含みます。</span><span class="sxs-lookup"><span data-stu-id="ee953-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="ee953-123">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="ee953-123">**./src/functions/functions.html**</span></span> | <span data-ttu-id="ee953-124">HTML</span><span class="sxs-lookup"><span data-stu-id="ee953-124">HTML</span></span> | <span data-ttu-id="ee953-125">カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。</span><span class="sxs-lookup"><span data-stu-id="ee953-125">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="ee953-126">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="ee953-126">**./manifest.xml**</span></span> | <span data-ttu-id="ee953-127">XML</span><span class="sxs-lookup"><span data-stu-id="ee953-127">XML</span></span> | <span data-ttu-id="ee953-128">カスタム関数 JavaScript、JSON、HTML ファイルなど、カスタム関数が使用する複数のファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="ee953-128">Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="ee953-129">また、作業ウィンドウ ファイルおよびコマンド ファイルの場所を表示すると共に、カスタム関数が使用するランタイムも指定します。</span><span class="sxs-lookup"><span data-stu-id="ee953-129">It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use.</span></span> |

### <a name="script-file"></a><span data-ttu-id="ee953-130">スクリプト ファイル</span><span class="sxs-lookup"><span data-stu-id="ee953-130">Script file</span></span>

<span data-ttu-id="ee953-131">スクリプト ファイル (**./src/customfunctions.js** または **/src/customfunctions.ts**) には、カスタム関数を定義するコードと関数を定義するコメントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="ee953-131">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="ee953-132">`add` カスタム関数は次のコードにより定義されます。</span><span class="sxs-lookup"><span data-stu-id="ee953-132">The following code defines the custom function `add`.</span></span> <span data-ttu-id="ee953-133">コード コメントは、Excel にカスタム関数を記述する JSON メタデータ ファイルを生成するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="ee953-133">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="ee953-134">必須の `@customfunction` コメントが最初に宣言されて、これがカスタム関数であることを示します。</span><span class="sxs-lookup"><span data-stu-id="ee953-134">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="ee953-135">次に、`description` プロパティに続いて、`first` および `second` の 2 つのパラメーターが宣言されます。</span><span class="sxs-lookup"><span data-stu-id="ee953-135">Next, two parameters are declared, `first` and `second`, followed by their `description` properties.</span></span> <span data-ttu-id="ee953-136">最後に `returns` の説明が記述されます。</span><span class="sxs-lookup"><span data-stu-id="ee953-136">Finally, a `returns` description is given.</span></span> <span data-ttu-id="ee953-137">カスタム関数で必要になるコメントに関する詳細については、「[カスタム関数の JSON メタデータを作成する](custom-functions-json-autogeneration.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ee953-137">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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
```

### <a name="manifest-file"></a><span data-ttu-id="ee953-138">マニフェスト ファイル</span><span class="sxs-lookup"><span data-stu-id="ee953-138">Manifest file</span></span>

<span data-ttu-id="ee953-139">カスタム関数 (Yo Office ジェネレーターによって作成されたプロジェクトの **./manifest.xml**) を定義するアドイン用 XML マニフェスト ファイルには、以下のような複数の機能があります。</span><span class="sxs-lookup"><span data-stu-id="ee953-139">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) does several things:</span></span>

- <span data-ttu-id="ee953-140">カスタム関数の名前空間を定義します。</span><span class="sxs-lookup"><span data-stu-id="ee953-140">Defines the namespace for your custom functions.</span></span> <span data-ttu-id="ee953-141">ユーザーがアドインの一部として関数を特定するのに役立つように、名前空間がカスタム関数の前に付加されます。</span><span class="sxs-lookup"><span data-stu-id="ee953-141">A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.</span></span>
- <span data-ttu-id="ee953-142">カスタム関数マニフェストに固有の `<ExtensionPoint>` および `<Resources>` 要素を使用します。</span><span class="sxs-lookup"><span data-stu-id="ee953-142">Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest.</span></span> <span data-ttu-id="ee953-143">これらの要素には、JavaScript、JSON、および HTML ファイルの場所に関する情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ee953-143">These elements contain the information about the locations of the JavaScript, JSON, and HTML files.</span></span>
- <span data-ttu-id="ee953-144">カスタム関数に使用するランタイムを指定します。</span><span class="sxs-lookup"><span data-stu-id="ee953-144">Specifies which runtime to use for your custom function.</span></span> <span data-ttu-id="ee953-145">別のランタイムを特段必要とする場合を除いて、共有ランタイムは関数と作業ウィンドウの間でデータを共有できるため、共有ランタイムを常に使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ee953-145">We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane.</span></span> <span data-ttu-id="ee953-146">共有ランタイムを使うことは、アドインが Microsoft Edge ではなく Internet Explorer 11 の使用を意味することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="ee953-146">Note that using a shared runtime means your add-in will use Internet Explorer 11, not Microsoft Edge.</span></span>

<span data-ttu-id="ee953-147">Yo Office ジェネレーターを使用してファイルを作成する場合、共有ランタイムはこのようなファイルの既定ではないため、それを使用するようにマニフェストを調整することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ee953-147">If you are using the Yo Office generator to create files, we recommend adjusting your manifest to use a shared runtime, as this is not the default for these files.</span></span> <span data-ttu-id="ee953-148">マニフェストを変更するには、「[Excel アドインを構成して、共有されている JavaScript ランタイムを使用する](configure-your-add-in-to-use-a-shared-runtime.md)」の手順に従います。</span><span class="sxs-lookup"><span data-stu-id="ee953-148">To change your manifest, follow the instructions in [Configure your Excel add-in to use a shared JavaScript runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="ee953-149">サンプル アドインからフル機能マニフェストを確認する方法については、「[この Github リポジトリ](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ee953-149">To see a full working manifest from a sample add-in, see [this Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a><span data-ttu-id="ee953-150">共同編集</span><span class="sxs-lookup"><span data-stu-id="ee953-150">Coauthoring</span></span>

<span data-ttu-id="ee953-151">Excel on the web および Microsoft 365 サブスクリプションに接続されている Windows では、Excel で共同編集することができます。</span><span class="sxs-lookup"><span data-stu-id="ee953-151">Excel on the web and Windows connected to a Microsoft 365 subscription allow you to coauthor in Excel.</span></span> <span data-ttu-id="ee953-152">ブックでカスタム関数を使用している場合、共同編集中の仕事仲間はカスタム関数のアドインを読み込むように要求されます。</span><span class="sxs-lookup"><span data-stu-id="ee953-152">If your workbook uses a custom function, your coauthoring colleague is prompted to load the custom function's add-in.</span></span> <span data-ttu-id="ee953-153">双方がアドインを読み込むと、共同編集によりカスタム関数は結果を共有します。</span><span class="sxs-lookup"><span data-stu-id="ee953-153">Once you both have loaded the add-in, the custom function shares results through coauthoring.</span></span>

<span data-ttu-id="ee953-154">共同編集の詳細については、「[Excel での共同編集](/office/vba/excel/concepts/about-coauthoring-in-excel)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ee953-154">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="ee953-155">既知の問題</span><span class="sxs-lookup"><span data-stu-id="ee953-155">Known issues</span></span>

<span data-ttu-id="ee953-156">既知の問題については、[Excel カスタム関数についての GitHub のレポート](https://github.com/OfficeDev/Excel-Custom-Functions/issues)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ee953-156">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="ee953-157">次の手順</span><span class="sxs-lookup"><span data-stu-id="ee953-157">Next steps</span></span>

<span data-ttu-id="ee953-158">カスタム関数を試してみましょう。</span><span class="sxs-lookup"><span data-stu-id="ee953-158">Want to try out custom functions?</span></span> <span data-ttu-id="ee953-159">もしまだであれば、簡単な[カスタム関数クイックスタート](../quickstarts/excel-custom-functions-quickstart.md)または、詳細な[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="ee953-159">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="ee953-160">独自のカスタム関数を試すもう 1 つの簡単な方法は[スクリプト ラボ](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)を使用し、アドインで Excel のカスタム関数を試してみることができます。</span><span class="sxs-lookup"><span data-stu-id="ee953-160">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="ee953-161">独自のカスタム関数を作成したり、提供されたサンプルを再生してみることができます。</span><span class="sxs-lookup"><span data-stu-id="ee953-161">You can try out creating your own custom function or play with the provided samples.</span></span>

## <a name="see-also"></a><span data-ttu-id="ee953-162">関連項目</span><span class="sxs-lookup"><span data-stu-id="ee953-162">See also</span></span> 
* [<span data-ttu-id="ee953-163">カスタム関数の要件</span><span class="sxs-lookup"><span data-stu-id="ee953-163">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="ee953-164">名前付けのガイドライン</span><span class="sxs-lookup"><span data-stu-id="ee953-164">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="ee953-165">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="ee953-165">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
