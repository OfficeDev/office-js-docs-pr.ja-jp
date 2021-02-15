---
ms.date: 01/08/2020
description: Office アドインの Excel カスタム関数を作成します。
title: Excel でカスタム関数を作成する
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 804895f3e10cac849dc20b67625e4f30164eb41d
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237673"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="66a9a-103">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="66a9a-103">Create custom functions in Excel</span></span>

<span data-ttu-id="66a9a-104">開発者は、カスタム関数を使用して関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="66a9a-105">ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="66a9a-106">次のアニメーション画像は、JavaScript または Typescript で作成した関数を呼び出すブックを示しています。</span><span class="sxs-lookup"><span data-stu-id="66a9a-106">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="66a9a-107">この例では、カスタム関数 `=MYFUNCTION.SPHEREVOLUME` は球の体積を計算します。</span><span class="sxs-lookup"><span data-stu-id="66a9a-107">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="66a9a-108">`=MYFUNCTION.SPHEREVOLUME` カスタム関数は次のコードにより定義されます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-108">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

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

> [!TIP]
> <span data-ttu-id="66a9a-109">カスタム関数アドインがカスタム関数のコードの実行に加えて作業ウィンドウまたはリボン ボタンを使用する場合、共有 JavaScript ランタイムを設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="66a9a-109">If your custom function add-in will use a task pane or a ribbon button, in addition to running custom function code, you will need to set up a shared JavaScript runtime.</span></span> <span data-ttu-id="66a9a-110">詳細については、「[Office アドインを構成して共有 JavaScript ランタイムを使用する ](../develop/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="66a9a-110">See [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) to learn more.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="66a9a-111">コードでカスタム関数を定義する方法</span><span class="sxs-lookup"><span data-stu-id="66a9a-111">How a custom function is defined in code</span></span>

<span data-ttu-id="66a9a-112">[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Excel のカスタム関数アドイン プロジェクトを作成する場合、使用する関数および作業ウィンドウを制御するファイルが作成されます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, it creates files which control your functions and task pane.</span></span> <span data-ttu-id="66a9a-113">このため、カスタム関数に重要なファイルに注意を集中できます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-113">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="66a9a-114">ファイル</span><span class="sxs-lookup"><span data-stu-id="66a9a-114">File</span></span> | <span data-ttu-id="66a9a-115">ファイル形式</span><span class="sxs-lookup"><span data-stu-id="66a9a-115">File format</span></span> | <span data-ttu-id="66a9a-116">説明</span><span class="sxs-lookup"><span data-stu-id="66a9a-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="66a9a-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="66a9a-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="66a9a-118">または</span><span class="sxs-lookup"><span data-stu-id="66a9a-118">or</span></span><br/><span data-ttu-id="66a9a-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="66a9a-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="66a9a-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="66a9a-120">JavaScript</span></span><br/><span data-ttu-id="66a9a-121">または</span><span class="sxs-lookup"><span data-stu-id="66a9a-121">or</span></span><br/><span data-ttu-id="66a9a-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="66a9a-122">TypeScript</span></span> | <span data-ttu-id="66a9a-123">カスタム関数を定義するコードが含みます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="66a9a-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="66a9a-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="66a9a-125">HTML</span><span class="sxs-lookup"><span data-stu-id="66a9a-125">HTML</span></span> | <span data-ttu-id="66a9a-126">カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。</span><span class="sxs-lookup"><span data-stu-id="66a9a-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="66a9a-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="66a9a-127">**./manifest.xml**</span></span> | <span data-ttu-id="66a9a-128">XML</span><span class="sxs-lookup"><span data-stu-id="66a9a-128">XML</span></span> | <span data-ttu-id="66a9a-129">カスタム関数 JavaScript、JSON、HTML ファイルなど、カスタム関数が使用する複数のファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="66a9a-129">Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="66a9a-130">また、作業ウィンドウ ファイルおよびコマンド ファイルの場所を表示すると共に、カスタム関数が使用するランタイムも指定します。</span><span class="sxs-lookup"><span data-stu-id="66a9a-130">It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use.</span></span> |

### <a name="script-file"></a><span data-ttu-id="66a9a-131">スクリプト ファイル</span><span class="sxs-lookup"><span data-stu-id="66a9a-131">Script file</span></span>

<span data-ttu-id="66a9a-132">スクリプト ファイル (**./src/customfunctions.js** または **/src/customfunctions.ts**) には、カスタム関数を定義するコードと関数を定義するコメントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="66a9a-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="66a9a-133">`add` カスタム関数は次のコードにより定義されます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-133">The following code defines the custom function `add`.</span></span> <span data-ttu-id="66a9a-134">コード コメントは、Excel にカスタム関数を記述する JSON メタデータ ファイルを生成するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-134">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="66a9a-135">必須の `@customfunction` コメントが最初に宣言されて、これがカスタム関数であることを示します。</span><span class="sxs-lookup"><span data-stu-id="66a9a-135">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="66a9a-136">次に、`description` プロパティに続いて、`first` および `second` の 2 つのパラメーターが宣言されます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-136">Next, two parameters are declared, `first` and `second`, followed by their `description` properties.</span></span> <span data-ttu-id="66a9a-137">最後に `returns` の説明が記述されます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-137">Finally, a `returns` description is given.</span></span> <span data-ttu-id="66a9a-138">カスタム関数で必要になるコメントに関する詳細については、「[カスタム関数の JSON メタデータを自動作成する](custom-functions-json-autogeneration.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="66a9a-138">For more information about what comments are required for your custom function, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

### <a name="manifest-file"></a><span data-ttu-id="66a9a-139">マニフェスト ファイル</span><span class="sxs-lookup"><span data-stu-id="66a9a-139">Manifest file</span></span>

<span data-ttu-id="66a9a-140">カスタム関数 (Yo Office ジェネレーターによって作成されたプロジェクトの **./manifest.xml**) を定義するアドイン用 XML マニフェスト ファイルには、以下のような複数の機能があります。</span><span class="sxs-lookup"><span data-stu-id="66a9a-140">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) does several things:</span></span>

- <span data-ttu-id="66a9a-141">カスタム関数の名前空間を定義します。</span><span class="sxs-lookup"><span data-stu-id="66a9a-141">Defines the namespace for your custom functions.</span></span> <span data-ttu-id="66a9a-142">ユーザーがアドインの一部として関数を特定するのに役立つように、名前空間がカスタム関数の前に付加されます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-142">A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.</span></span>
- <span data-ttu-id="66a9a-143">カスタム関数マニフェストに固有の `<ExtensionPoint>` および `<Resources>` 要素を使用します。</span><span class="sxs-lookup"><span data-stu-id="66a9a-143">Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest.</span></span> <span data-ttu-id="66a9a-144">これらの要素には、JavaScript、JSON、および HTML ファイルの場所に関する情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="66a9a-144">These elements contain the information about the locations of the JavaScript, JSON, and HTML files.</span></span>
- <span data-ttu-id="66a9a-145">カスタム関数に使用するランタイムを指定します。</span><span class="sxs-lookup"><span data-stu-id="66a9a-145">Specifies which runtime to use for your custom function.</span></span> <span data-ttu-id="66a9a-146">別のランタイムを特段必要とする場合を除いて、共有ランタイムは関数と作業ウィンドウの間でデータを共有できるため、共有ランタイムを常に使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="66a9a-146">We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane.</span></span> <span data-ttu-id="66a9a-147">共有ランタイムを使うことは、アドインが Microsoft Edge ではなく Internet Explorer 11 の使用を意味することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="66a9a-147">Note that using a shared runtime means your add-in will use Internet Explorer 11, not Microsoft Edge.</span></span>

<span data-ttu-id="66a9a-148">Yo Office ジェネレーターを使用してファイルを作成する場合、共有ランタイムはこのようなファイルの既定ではないため、それを使用するようにマニフェストを調整することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="66a9a-148">If you are using the Yo Office generator to create files, we recommend adjusting your manifest to use a shared runtime, as this is not the default for these files.</span></span> <span data-ttu-id="66a9a-149">マニフェストを変更するには、「[Excel アドインを構成して、共有されている JavaScript ランタイムを使用する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)」の手順に従います。</span><span class="sxs-lookup"><span data-stu-id="66a9a-149">To change your manifest, follow the instructions in [Configure your Excel add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="66a9a-150">サンプル アドインからフル機能マニフェストを確認する方法については、「[この Github リポジトリ](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="66a9a-150">To see a full working manifest from a sample add-in, see [this Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a><span data-ttu-id="66a9a-151">共同編集</span><span class="sxs-lookup"><span data-stu-id="66a9a-151">Coauthoring</span></span>

<span data-ttu-id="66a9a-152">Excel on the web および Microsoft 365 サブスクリプションに接続されている Windows 版の Excel では、Excel で共同編集を行うことができます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-152">Excel on the web and on Windows connected to a Microsoft 365 subscription allow you to coauthor in Excel.</span></span> <span data-ttu-id="66a9a-153">ブックでカスタム関数を使用している場合、共同編集中の仕事仲間はカスタム関数のアドインを読み込むように要求されます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-153">If your workbook uses a custom function, your coauthoring colleague is prompted to load the custom function's add-in.</span></span> <span data-ttu-id="66a9a-154">双方がアドインを読み込むと、共同編集によりカスタム関数は結果を共有します。</span><span class="sxs-lookup"><span data-stu-id="66a9a-154">Once you both have loaded the add-in, the custom function shares results through coauthoring.</span></span>

<span data-ttu-id="66a9a-155">共同編集の詳細については、「[Excel での共同編集](/office/vba/excel/concepts/about-coauthoring-in-excel)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="66a9a-155">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="next-steps"></a><span data-ttu-id="66a9a-156">次の手順</span><span class="sxs-lookup"><span data-stu-id="66a9a-156">Next steps</span></span>

<span data-ttu-id="66a9a-157">カスタム関数を試してみましょう。</span><span class="sxs-lookup"><span data-stu-id="66a9a-157">Want to try out custom functions?</span></span> <span data-ttu-id="66a9a-158">もしまだであれば、簡単な[カスタム関数クイックスタート](../quickstarts/excel-custom-functions-quickstart.md)または、詳細な[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="66a9a-158">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="66a9a-159">独自のカスタム関数を試すもう 1 つの簡単な方法は[スクリプト ラボ](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)を使用し、アドインで Excel のカスタム関数を試してみることができます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-159">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="66a9a-160">独自のカスタム関数を作成したり、提供されたサンプルを再生してみることができます。</span><span class="sxs-lookup"><span data-stu-id="66a9a-160">You can try out creating your own custom function or play with the provided samples.</span></span>

## <a name="see-also"></a><span data-ttu-id="66a9a-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="66a9a-161">See also</span></span> 
* [<span data-ttu-id="66a9a-162">Microsoft 365 開発者プログラムについて</span><span class="sxs-lookup"><span data-stu-id="66a9a-162">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
* [<span data-ttu-id="66a9a-163">カスタム関数の要件セット</span><span class="sxs-lookup"><span data-stu-id="66a9a-163">Custom functions requirement sets</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="66a9a-164">カスタム関数の名前付けのガイドライン</span><span class="sxs-lookup"><span data-stu-id="66a9a-164">Custom functions naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="66a9a-165">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="66a9a-165">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="66a9a-166">Office アドインを構成して共有 JavaScript ランタイムを使用する</span><span class="sxs-lookup"><span data-stu-id="66a9a-166">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
