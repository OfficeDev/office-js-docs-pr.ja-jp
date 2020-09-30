---
ms.date: 09/25/2020
description: JSDoc タグを使用して、カスタム関数の JSON メタデータを動的に作成します。
title: カスタム関数用の JSON メタデータの自動生成
localization_priority: Normal
ms.openlocfilehash: 151dc7c97b2a98743906b7e0a920fdc1eff62e7f
ms.sourcegitcommit: 42202d7e2ac24dffa77cf937f5697a1cd79ee790
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/30/2020
ms.locfileid: "48308538"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="0ecbf-103">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="0ecbf-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="0ecbf-104">Excel カスタム関数が JavaScript または TypeScript で記述されている場合、カスタム関数に関する追加の情報を提供するために、[JSDoc タグ](https://jsdoc.app/)が使用されます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-104">When an Excel custom function is written in JavaScript or TypeScript, [JSDoc tags](https://jsdoc.app/) are used to provide extra information about the custom function.</span></span> <span data-ttu-id="0ecbf-105">JSDoc タグはビルド時に使用して、[JSON メタデータ ファイル](custom-functions-json.md)を作成します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="0ecbf-106">JSDoc タグを使用すると、JSON メタデータ ファイルを手動で編集する手間が省けます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="0ecbf-107">JavaScript または TypeScript 関数のコード コメントに`@customfunction`タグを追加して、カスタム関数としてマークします。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="0ecbf-108">関数パラメーターの型は、JavaScript の [@param](#param) タグを使用して指定するか、TypeScript の[関数の型](https://www.typescriptlang.org/docs/handbook/functions.html)から指定できます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="0ecbf-109">詳細については、「[@param](#param) タグ」セクションと「[型](#types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-109">For more information, see the [@param](#param) tag and [Types](#types) sections.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="0ecbf-110">関数に説明を追加する</span><span class="sxs-lookup"><span data-stu-id="0ecbf-110">Adding a description to a function</span></span>

<span data-ttu-id="0ecbf-111">説明は、カスタム関数の機能を理解するためのヘルプが必要な場合に、ヘルプ テキストとしてユーザーに表示されます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="0ecbf-112">説明に特定のタグは必要ありません。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="0ecbf-113">JSDoc コメントに簡単な説明を入力するだけです。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="0ecbf-114">一般に、説明は JSDoc コメント セクションの先頭に配置されますが、配置場所に関係なく機能します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="0ecbf-115">組み込み関数の説明の例を表示するには、Excel を開き、**[数式]** タブに移動し、**[関数の​​挿入]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="0ecbf-116">すべての関数の説明を参照したり、独自のカスタム関数を一覧表示したりすることができます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="0ecbf-117">次の例では、「球の体積を計算します。」</span><span class="sxs-lookup"><span data-stu-id="0ecbf-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="0ecbf-118">が、カスタム関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="0ecbf-119">JSDoc タグ</span><span class="sxs-lookup"><span data-stu-id="0ecbf-119">JSDoc Tags</span></span>

<span data-ttu-id="0ecbf-120">Excel カスタム関数では、次の JSDoc タグがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-120">The following JSDoc tags are supported in Excel custom functions.</span></span>

* [<span data-ttu-id="0ecbf-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="0ecbf-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="0ecbf-122">[@customfunction](#customfunction) id 名</span><span class="sxs-lookup"><span data-stu-id="0ecbf-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="0ecbf-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="0ecbf-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="0ecbf-124">[@param](#param) _{type}_ 名前の説明</span><span class="sxs-lookup"><span data-stu-id="0ecbf-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="0ecbf-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="0ecbf-125">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="0ecbf-126">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="0ecbf-126">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="0ecbf-127">@streaming</span><span class="sxs-lookup"><span data-stu-id="0ecbf-127">@streaming</span></span>](#streaming)
* [<span data-ttu-id="0ecbf-128">@volatile</span><span class="sxs-lookup"><span data-stu-id="0ecbf-128">@volatile</span></span>](#volatile)

---
<a id="cancelable"></a>

### <a name="cancelable"></a><span data-ttu-id="0ecbf-129">@cancelable</span><span class="sxs-lookup"><span data-stu-id="0ecbf-129">@cancelable</span></span>

<span data-ttu-id="0ecbf-130">関数が取り消されたときにカスタム関数が処理を実行することを示します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-130">Indicates that a custom function performs an action when the function is canceled.</span></span>

<span data-ttu-id="0ecbf-131">最後の関数パラメーターは `CustomFunctions.CancelableInvocation` の型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-131">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="0ecbf-132">関数は、関数 `oncanceled` が取り消されたときの結果を示すために、関数をプロパティに割り当てることができます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-132">The function can assign a function to the `oncanceled` property to denote the result when the function is canceled.</span></span>

<span data-ttu-id="0ecbf-133">最後の関数のパラメーターが `CustomFunctions.CancelableInvocation` 型の場合、タグは表示されませんが、`@cancelable` と見なされます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-133">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag isn't present.</span></span>

<span data-ttu-id="0ecbf-134">関数には `@cancelable` と `@streaming` の両方のタグを含めることはできません。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-134">A function can't have both `@cancelable` and `@streaming` tags.</span></span>

---
<a id="customfunction"></a>

### <a name="customfunction"></a><span data-ttu-id="0ecbf-135">@customfunction</span><span class="sxs-lookup"><span data-stu-id="0ecbf-135">@customfunction</span></span>

<span data-ttu-id="0ecbf-136">構文: @customfunction _id_ _名_</span><span class="sxs-lookup"><span data-stu-id="0ecbf-136">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="0ecbf-137">このタグは、JavaScript/TypeScript 関数が Excel カスタム関数であることを示します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-137">This tag indicates that the JavaScript/TypeScript function is an Excel custom function.</span></span> <span data-ttu-id="0ecbf-138">カスタム関数のメタデータを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-138">It is required to create metadata for the custom function.</span></span>

<span data-ttu-id="0ecbf-139">このタグの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-139">The following shows an example of this tag.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="0ecbf-140">id</span><span class="sxs-lookup"><span data-stu-id="0ecbf-140">id</span></span>

<span data-ttu-id="0ecbf-141">は、 `id` カスタム関数を識別します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-141">The `id` identifies a custom function.</span></span>

* <span data-ttu-id="0ecbf-142">`id` が提供されていない場合、JavaScript または TypeScript の関数名は大文字に変換され、許可されない文字は削除されます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-142">If `id` isn't provided, the JavaScript/TypeScript function name is converted to uppercase and disallowed characters are removed.</span></span>
* <span data-ttu-id="0ecbf-143">`id` はすべてのカスタム関数で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-143">The `id` must be unique for all custom functions.</span></span>
* <span data-ttu-id="0ecbf-144">指定できる文字は、A から Z、a から z、0 から 9、アンダースコア (\_)、ピリオド (.) に制限されます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-144">The allowed characters are limited to: A-Z, a-z, 0-9, underscores (\_), and period (.).</span></span>

<span data-ttu-id="0ecbf-145">次の例では、インクリメントは関数の `id` と `name` です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-145">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="0ecbf-146">name</span><span class="sxs-lookup"><span data-stu-id="0ecbf-146">name</span></span>

<span data-ttu-id="0ecbf-147">カスタム関数の表示用の `name` を提供します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-147">Provides the display `name` for the custom function.</span></span>

* <span data-ttu-id="0ecbf-148">name が指定されていない場合、id が名前としても使用されます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-148">If name isn't provided, the id is also used as the name.</span></span>
* <span data-ttu-id="0ecbf-149">使用できる文字は、文字 [Unicode アルファベット](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、ピリオド (.)、およびアンダースコア (\_)です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-149">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="0ecbf-150">最初の文字は、アルファベット文字にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-150">Must start with a letter.</span></span>
* <span data-ttu-id="0ecbf-151">最大文字数は 128 文字です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-151">Maximum length is 128 characters.</span></span>

<span data-ttu-id="0ecbf-152">次の例では、INC は関数の`id` で、 `increment` は`name`です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-152">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="0ecbf-153">説明</span><span class="sxs-lookup"><span data-stu-id="0ecbf-153">description</span></span>

<span data-ttu-id="0ecbf-154">Excel のユーザーには、関数の入力時に説明が表示され、関数の動作を指定します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-154">A description appears to users in Excel as they are entering the function and specifies what the function does.</span></span> <span data-ttu-id="0ecbf-155">説明に特定のタグは必要ありません。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-155">A description doesn't require any specific tag.</span></span> <span data-ttu-id="0ecbf-156">JSDoc コメント内に関数の機能を説明するフレーズを入力して、カスタム関数に説明を追加します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-156">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="0ecbf-157">既定では、JSDoc コメント セクションでタグが付けられていないテキストは、関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-157">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span>

<span data-ttu-id="0ecbf-158">次の例では、「2 つの数値を加算する関数」というフレーズが、ID プロパティ `ADD` のカスタム関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-158">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
<a id="helpurl"></a>

### <a name="helpurl"></a><span data-ttu-id="0ecbf-159">@helpurl</span><span class="sxs-lookup"><span data-stu-id="0ecbf-159">@helpurl</span></span>

<span data-ttu-id="0ecbf-160">構文: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="0ecbf-160">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="0ecbf-161">指定された _url_ が Excel で表示されます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-161">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="0ecbf-162">次の例では、 `helpurl` がに `www.contoso.com/weatherhelp` なります。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-162">In the following example, the `helpurl` is `www.contoso.com/weatherhelp`.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
<a id="param"></a>

### <a name="param"></a><span data-ttu-id="0ecbf-163">@param</span><span class="sxs-lookup"><span data-stu-id="0ecbf-163">@param</span></span>

#### <a name="javascript"></a><span data-ttu-id="0ecbf-164">JavaScript</span><span class="sxs-lookup"><span data-stu-id="0ecbf-164">JavaScript</span></span>

<span data-ttu-id="0ecbf-165">JavaScript 構文: @param {type} 名_の説明_</span><span class="sxs-lookup"><span data-stu-id="0ecbf-165">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="0ecbf-166">`{type}` 中かっこで囲まれた型情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-166">`{type}` specifies the type info within curly braces.</span></span> <span data-ttu-id="0ecbf-167">使用できる型に関する詳細については、「[型](#types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-167">See the [Types](#types) section for more information about the types which may be used.</span></span> <span data-ttu-id="0ecbf-168">種類が指定されていない場合は、既定の種類が使用され `any` ます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-168">If no type is specified, the default type `any` will be used.</span></span>
* <span data-ttu-id="0ecbf-169">`name` @param タグを適用するパラメーターを指定します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-169">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="0ecbf-170">これは必須です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-170">It is required.</span></span>
* <span data-ttu-id="0ecbf-171">`description` は、Excel で表示される関数のパラメーターの説明を示します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-171">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="0ecbf-172">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-172">It is optional.</span></span>

<span data-ttu-id="0ecbf-173">カスタム関数内のパラメーターを省略可能と指定する方法:</span><span class="sxs-lookup"><span data-stu-id="0ecbf-173">To denote a custom function parameter as optional:</span></span>

* <span data-ttu-id="0ecbf-174">パラメーター名を角かっこで囲みます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-174">Put square brackets around the parameter name.</span></span> <span data-ttu-id="0ecbf-175">例: `@param {string} [text] Optional text`。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-175">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="0ecbf-176">省略可能なパラメーターの既定値は `null` です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-176">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="0ecbf-177">次の例は、2つまたは3つの数字を省略可能なパラメーターとして追加する ADD 関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-177">The following example shows a ADD function that adds two or three numbers, with the third number as an optional parameter.</span></span>

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### <a name="typescript"></a><span data-ttu-id="0ecbf-178">TypeScript</span><span class="sxs-lookup"><span data-stu-id="0ecbf-178">TypeScript</span></span>

<span data-ttu-id="0ecbf-179">TypeScript 構文: @param 名 _の説明_</span><span class="sxs-lookup"><span data-stu-id="0ecbf-179">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="0ecbf-180">`name` @param タグを適用するパラメーターを指定します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-180">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="0ecbf-181">これは必須です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-181">It is required.</span></span>
* <span data-ttu-id="0ecbf-182">`description` は、Excel で表示される関数のパラメーターの説明を示します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-182">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="0ecbf-183">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-183">It is optional.</span></span>

<span data-ttu-id="0ecbf-184">使用できる関数のパラメーターの型に関する詳細については、「[型](#types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-184">See the [Types](#types) section for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="0ecbf-185">カスタム関数のパラメーターを省略可能として示すには、以下のいずれかを実行します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-185">To denote a custom function parameter as optional, do one of the following:</span></span>

* <span data-ttu-id="0ecbf-186">省略可能なパラメーターを使用する。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-186">Use an optional parameter.</span></span> <span data-ttu-id="0ecbf-187">例: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="0ecbf-187">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="0ecbf-188">パラメーターに既定値を指定する。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-188">Give the parameter a default value.</span></span> <span data-ttu-id="0ecbf-189">例: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="0ecbf-189">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="0ecbf-190">@param の詳しい説明については、「[JSDoc](https://jsdoc.app/tags-param.html)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-190">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="0ecbf-191">省略可能なパラメーターの既定値は `null` です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-191">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="0ecbf-192">次の例は、2 つの数値を加算する `add` 関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-192">The following example shows the `add` function that adds two numbers.</span></span>

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
<a id="requiresAddress"></a>

### <a name="requiresaddress"></a><span data-ttu-id="0ecbf-193">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="0ecbf-193">@requiresAddress</span></span>

<span data-ttu-id="0ecbf-194">関数が評価されているセルのアドレスを指定する必要があることを示します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-194">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="0ecbf-195">最後の関数のパラメーターは、`CustomFunctions.Invocation` 型または派生型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-195">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="0ecbf-196">関数が呼び出されると、`address` プロパティにアドレスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-196">When the function is called, the `address` property will contain the address.</span></span>

---
<a id="returns"></a>

### <a name="returns"></a><span data-ttu-id="0ecbf-197">@returns</span><span class="sxs-lookup"><span data-stu-id="0ecbf-197">@returns</span></span>

<span data-ttu-id="0ecbf-198">構文: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="0ecbf-198">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="0ecbf-199">戻り値の型を指定します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-199">Provides the type for the return value.</span></span>

<span data-ttu-id="0ecbf-200">`{type}` を省略すると、TypeScript の型情報が使用されます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-200">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="0ecbf-201">型情報がない場合、型は `any` になります。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-201">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="0ecbf-202">次の例は、 `@returns` タグを使用する `add` 関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-202">The following example shows the `add` function that uses the `@returns` tag.</span></span>

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
<a id="streaming"></a>

### <a name="streaming"></a><span data-ttu-id="0ecbf-203">@streaming</span><span class="sxs-lookup"><span data-stu-id="0ecbf-203">@streaming</span></span>

<span data-ttu-id="0ecbf-204">カスタム関数がストリーミング関数であることを示すのに使用されます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-204">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="0ecbf-205">最後のパラメーターの型は `CustomFunctions.StreamingInvocation<ResultType>` です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-205">The last parameter is of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="0ecbf-206">関数が戻り `void` ます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-206">The function returns `void`.</span></span>

<span data-ttu-id="0ecbf-207">ストリーミング関数は、値を直接返すのではなく、 `setResult(result: ResultType)` 最後のパラメーターを使用して呼び出します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-207">Streaming functions don't return values directly, instead they call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="0ecbf-208">ストリーム関数によってスローされる例外は無視されます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-208">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="0ecbf-209">`setResult()` が、エラー結果を示すために、Error により呼び出されることがあります。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-209">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="0ecbf-210">ストリーミング関数と詳細については、「[ストリーミング関数を作成する](custom-functions-web-reqs.md#make-a-streaming-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-210">For an example of a streaming function and more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="0ecbf-211">ストリーミング関数は、[@volatile](#volatile) としてマークできません。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-211">Streaming functions can't be marked as [@volatile](#volatile).</span></span>

---
<a id="volatile"></a>

### <a name="volatile"></a><span data-ttu-id="0ecbf-212">@volatile</span><span class="sxs-lookup"><span data-stu-id="0ecbf-212">@volatile</span></span>

<span data-ttu-id="0ecbf-213">揮発性関数とは、引数を取らない場合や引数が変更されていない場合でも、ある瞬間と次の瞬間では結果が異なる関数です。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-213">A volatile function is one whose result isn't the same from one moment to the next, even if it takes no arguments or the arguments haven't changed.</span></span> <span data-ttu-id="0ecbf-214">Excel では、再計算が実行される度に、揮発性関数を含むセルはすべての参照先と共に、再評価されます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-214">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="0ecbf-215">このため、揮発性関数を多用し過ぎると再計算にかかる時間が長くなる可能性があるため、多用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-215">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="0ecbf-216">ストリーミング関数に揮発性関数は使用できません。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-216">Streaming functions can't be volatile.</span></span>

<span data-ttu-id="0ecbf-217">次の関数は揮発性で、 `@volatile` タグを使用します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-217">The following function is volatile and uses the `@volatile` tag.</span></span>

```js
/**
 * Simulates rolling a 6-sided die.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a><span data-ttu-id="0ecbf-218">型</span><span class="sxs-lookup"><span data-stu-id="0ecbf-218">Types</span></span>

<span data-ttu-id="0ecbf-219">パラメーターの型を指定すると、Excel は値を指定した型に変換してから関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-219">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="0ecbf-220">型が`any`の場合、変換は実行されません。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-220">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="0ecbf-221">値の型</span><span class="sxs-lookup"><span data-stu-id="0ecbf-221">Value types</span></span>

<span data-ttu-id="0ecbf-222">1 つの値は、`boolean`、 `number`、`string`の型のいずれかを使用して表現できます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-222">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="0ecbf-223">マトリックス型</span><span class="sxs-lookup"><span data-stu-id="0ecbf-223">Matrix type</span></span>

<span data-ttu-id="0ecbf-224">2 次元配列型を使用して、パラメーターまたは戻り値を値のマトリックスにします。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-224">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="0ecbf-225">たとえば、`number[][]`の型は数字のマトリックスを示します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-225">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="0ecbf-226">`string[][]` は、文字列のマトリックスを示します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-226">`string[][]` indicates a matrix of strings.</span></span>

### <a name="error-type"></a><span data-ttu-id="0ecbf-227">エラーの種類</span><span class="sxs-lookup"><span data-stu-id="0ecbf-227">Error type</span></span>

<span data-ttu-id="0ecbf-228">非ストリーミング関数は、エラーの種類を返すことによりエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-228">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="0ecbf-229">ストリーミング関数は、エラーの種類で `setResult()` を返してエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-229">A streaming function can indicate an error by calling `setResult()` with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="0ecbf-230">Promise</span><span class="sxs-lookup"><span data-stu-id="0ecbf-230">Promise</span></span>

<span data-ttu-id="0ecbf-231">関数は Promise を返すことができます。これは、promise が解決されたときに値を提供します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-231">A function can return a Promise, that provides the value when the promise is resolved.</span></span> <span data-ttu-id="0ecbf-232">Promise が拒否されると、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-232">If the promise is rejected, then it will throw an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="0ecbf-233">その他の型</span><span class="sxs-lookup"><span data-stu-id="0ecbf-233">Other types</span></span>

<span data-ttu-id="0ecbf-234">その他の型は、エラーとして処理されます。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-234">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="0ecbf-235">次の手順</span><span class="sxs-lookup"><span data-stu-id="0ecbf-235">Next steps</span></span>

<span data-ttu-id="0ecbf-236">[カスタム関数用の命名規則](custom-functions-naming.md)について説明します。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-236">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="0ecbf-237">または、[JSON ファイルを手で書く](custom-functions-json.md)必要のある[機能をローカライズする](custom-functions-localize.md)方法を確認してください。</span><span class="sxs-lookup"><span data-stu-id="0ecbf-237">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0ecbf-238">関連項目</span><span class="sxs-lookup"><span data-stu-id="0ecbf-238">See also</span></span>

* [<span data-ttu-id="0ecbf-239">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="0ecbf-239">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="0ecbf-240">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0ecbf-240">Create custom functions in Excel</span></span>](custom-functions-overview.md)
