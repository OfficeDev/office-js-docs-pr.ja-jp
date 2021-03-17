---
ms.date: 03/15/2021
description: JSDoc タグを使用して、カスタム関数の JSON メタデータを動的に作成します。
title: カスタム関数用の JSON メタデータの自動生成
localization_priority: Normal
ms.openlocfilehash: 344239c35e38bd88bfee5338289d1c2a929ea14c
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836866"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="d5c40-103">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="d5c40-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="d5c40-104">Excel カスタム関数が JavaScript または TypeScript で記述されている場合、カスタム関数に関する追加の情報を提供するために、[JSDoc タグ](https://jsdoc.app/)が使用されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-104">When an Excel custom function is written in JavaScript or TypeScript, [JSDoc tags](https://jsdoc.app/) are used to provide extra information about the custom function.</span></span> <span data-ttu-id="d5c40-105">JSDoc タグはビルド時に使用して、JSON メタデータ ファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-105">The JSDoc tags are then used at build time to create the JSON metadata file.</span></span> <span data-ttu-id="d5c40-106">JSDoc タグを使用すると、JSON メタデータ ファイルを手動で [編集する手間が省きます](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="d5c40-106">Using JSDoc tags saves you from the effort of [manually editing the JSON metadata file](custom-functions-json.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="d5c40-107">JavaScript または TypeScript 関数のコード コメントに`@customfunction`タグを追加して、カスタム関数としてマークします。</span><span class="sxs-lookup"><span data-stu-id="d5c40-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="d5c40-108">関数パラメーターの型は、JavaScript の [@param](#param) タグを使用して指定するか、TypeScript の[関数の型](https://www.typescriptlang.org/docs/handbook/functions.html)から指定できます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="d5c40-109">詳細については、「[@param](#param) タグ」セクションと「[型](#types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d5c40-109">For more information, see the [@param](#param) tag and [Types](#types) sections.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="d5c40-110">関数に説明を追加する</span><span class="sxs-lookup"><span data-stu-id="d5c40-110">Adding a description to a function</span></span>

<span data-ttu-id="d5c40-111">説明は、カスタム関数の機能を理解するためのヘルプが必要な場合に、ヘルプ テキストとしてユーザーに表示されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="d5c40-112">説明に特定のタグは必要ありません。</span><span class="sxs-lookup"><span data-stu-id="d5c40-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="d5c40-113">JSDoc コメントに簡単な説明を入力するだけです。</span><span class="sxs-lookup"><span data-stu-id="d5c40-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="d5c40-114">一般に、説明は JSDoc コメント セクションの先頭に配置されますが、配置場所に関係なく機能します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="d5c40-115">組み込み関数の説明の例を表示するには、Excel を開き、**[数式]** タブに移動し、**[関数の​​挿入]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="d5c40-116">すべての関数の説明を参照したり、独自のカスタム関数を一覧表示したりすることができます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="d5c40-117">次の例では、「球の体積を計算します。」</span><span class="sxs-lookup"><span data-stu-id="d5c40-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="d5c40-118">が、カスタム関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="d5c40-119">JSDoc タグ</span><span class="sxs-lookup"><span data-stu-id="d5c40-119">JSDoc Tags</span></span>

<span data-ttu-id="d5c40-120">Excel カスタム関数では、次の JSDoc タグがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="d5c40-120">The following JSDoc tags are supported in Excel custom functions.</span></span>

* [<span data-ttu-id="d5c40-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="d5c40-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="d5c40-122">[@customfunction](#customfunction) id 名</span><span class="sxs-lookup"><span data-stu-id="d5c40-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="d5c40-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="d5c40-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="d5c40-124">[@param](#param) _{type}_ 名前の説明</span><span class="sxs-lookup"><span data-stu-id="d5c40-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="d5c40-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="d5c40-125">@requiresAddress</span></span>](#requiresAddress)
* [<span data-ttu-id="d5c40-126">@requiresParameterAddresses</span><span class="sxs-lookup"><span data-stu-id="d5c40-126">@requiresParameterAddresses</span></span>](#requiresParameterAddresses)
* <span data-ttu-id="d5c40-127">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="d5c40-127">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="d5c40-128">@streaming</span><span class="sxs-lookup"><span data-stu-id="d5c40-128">@streaming</span></span>](#streaming)
* [<span data-ttu-id="d5c40-129">@volatile</span><span class="sxs-lookup"><span data-stu-id="d5c40-129">@volatile</span></span>](#volatile)

---
<a id="cancelable"></a>
### <a name="cancelable"></a><span data-ttu-id="d5c40-130">@cancelable</span><span class="sxs-lookup"><span data-stu-id="d5c40-130">@cancelable</span></span>

<span data-ttu-id="d5c40-131">関数が取り消された場合に、カスタム関数がアクションを実行します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-131">Indicates that a custom function performs an action when the function is canceled.</span></span>

<span data-ttu-id="d5c40-132">最後の関数パラメーターは `CustomFunctions.CancelableInvocation` の型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="d5c40-132">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="d5c40-133">関数はプロパティに関数を割り当て、関数が取り消された場合の結果 `oncanceled` を示します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-133">The function can assign a function to the `oncanceled` property to denote the result when the function is canceled.</span></span>

<span data-ttu-id="d5c40-134">最後の関数のパラメーターが `CustomFunctions.CancelableInvocation` 型の場合、タグは表示されませんが、`@cancelable` と見なされます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-134">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag isn't present.</span></span>

<span data-ttu-id="d5c40-135">関数には `@cancelable` と `@streaming` の両方のタグを含めることはできません。</span><span class="sxs-lookup"><span data-stu-id="d5c40-135">A function can't have both `@cancelable` and `@streaming` tags.</span></span>

<a id="customfunction"></a>

### <a name="customfunction"></a><span data-ttu-id="d5c40-136">@customfunction</span><span class="sxs-lookup"><span data-stu-id="d5c40-136">@customfunction</span></span>

<span data-ttu-id="d5c40-137">構文: @customfunction _id_ _名_</span><span class="sxs-lookup"><span data-stu-id="d5c40-137">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="d5c40-138">このタグは、JavaScript/TypeScript 関数が Excel カスタム関数を表します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-138">This tag indicates that the JavaScript/TypeScript function is an Excel custom function.</span></span> <span data-ttu-id="d5c40-139">カスタム関数のメタデータを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d5c40-139">It is required to create metadata for the custom function.</span></span>

<span data-ttu-id="d5c40-140">次に、このタグの例を示します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-140">The following shows an example of this tag.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="d5c40-141">id</span><span class="sxs-lookup"><span data-stu-id="d5c40-141">id</span></span>

<span data-ttu-id="d5c40-142">カスタム `id` 関数を識別します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-142">The `id` identifies a custom function.</span></span>

* <span data-ttu-id="d5c40-143">`id` が提供されていない場合、JavaScript または TypeScript の関数名は大文字に変換され、許可されない文字は削除されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-143">If `id` isn't provided, the JavaScript/TypeScript function name is converted to uppercase and disallowed characters are removed.</span></span>
* <span data-ttu-id="d5c40-144">`id` はすべてのカスタム関数で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="d5c40-144">The `id` must be unique for all custom functions.</span></span>
* <span data-ttu-id="d5c40-145">指定できる文字は、A から Z、a から z、0 から 9、アンダースコア (\_)、ピリオド (.) に制限されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-145">The allowed characters are limited to: A-Z, a-z, 0-9, underscores (\_), and period (.).</span></span>

<span data-ttu-id="d5c40-146">次の例では、インクリメントは関数の `id` と `name` です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-146">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="d5c40-147">name</span><span class="sxs-lookup"><span data-stu-id="d5c40-147">name</span></span>

<span data-ttu-id="d5c40-148">カスタム関数の表示用の `name` を提供します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-148">Provides the display `name` for the custom function.</span></span>

* <span data-ttu-id="d5c40-149">name が指定されていない場合、id が名前としても使用されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-149">If name isn't provided, the id is also used as the name.</span></span>
* <span data-ttu-id="d5c40-150">使用できる文字は、文字 [Unicode アルファベット](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、ピリオド (.)、およびアンダースコア (\_)です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-150">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="d5c40-151">最初の文字は、アルファベット文字にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="d5c40-151">Must start with a letter.</span></span>
* <span data-ttu-id="d5c40-152">最大文字数は 128 文字です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-152">Maximum length is 128 characters.</span></span>

<span data-ttu-id="d5c40-153">次の例では、INC は関数の`id` で、 `increment` は`name`です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-153">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="d5c40-154">説明</span><span class="sxs-lookup"><span data-stu-id="d5c40-154">description</span></span>

<span data-ttu-id="d5c40-155">関数を入力すると、Excel のユーザーに説明が表示され、関数の動作を指定します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-155">A description appears to users in Excel as they are entering the function and specifies what the function does.</span></span> <span data-ttu-id="d5c40-156">説明に特定のタグは必要ありません。</span><span class="sxs-lookup"><span data-stu-id="d5c40-156">A description doesn't require any specific tag.</span></span> <span data-ttu-id="d5c40-157">JSDoc コメント内に関数の機能を説明するフレーズを入力して、カスタム関数に説明を追加します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-157">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="d5c40-158">既定では、JSDoc コメント セクションでタグが付けられていないテキストは、関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-158">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span>

<span data-ttu-id="d5c40-159">次の例では、「2 つの数値を加算する関数」というフレーズが、ID プロパティ `ADD` のカスタム関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-159">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

<a id="helpurl"></a>
### <a name="helpurl"></a><span data-ttu-id="d5c40-160">@helpurl</span><span class="sxs-lookup"><span data-stu-id="d5c40-160">@helpurl</span></span>

<span data-ttu-id="d5c40-161">構文: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="d5c40-161">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="d5c40-162">指定された _url_ が Excel で表示されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-162">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="d5c40-163">次の例では、 `helpurl` です `www.contoso.com/weatherhelp` 。</span><span class="sxs-lookup"><span data-stu-id="d5c40-163">In the following example, the `helpurl` is `www.contoso.com/weatherhelp`.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

<a id="param"></a>
### <a name="param"></a><span data-ttu-id="d5c40-164">@param</span><span class="sxs-lookup"><span data-stu-id="d5c40-164">@param</span></span>

#### <a name="javascript"></a><span data-ttu-id="d5c40-165">JavaScript</span><span class="sxs-lookup"><span data-stu-id="d5c40-165">JavaScript</span></span>

<span data-ttu-id="d5c40-166">JavaScript 構文: @param {type} 名 _の説明_</span><span class="sxs-lookup"><span data-stu-id="d5c40-166">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="d5c40-167">`{type}` 中かっこ内の型情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-167">`{type}` specifies the type info within curly braces.</span></span> <span data-ttu-id="d5c40-168">使用できる型に関する詳細については、「[型](#types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d5c40-168">See the [Types](#types) section for more information about the types which may be used.</span></span> <span data-ttu-id="d5c40-169">型が指定されていない場合は、既定の型 `any` が使用されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-169">If no type is specified, the default type `any` will be used.</span></span>
* <span data-ttu-id="d5c40-170">`name` タグが適用されるパラメーター@param指定します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-170">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="d5c40-171">必須です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-171">It is required.</span></span>
* <span data-ttu-id="d5c40-172">`description` は、Excel で表示される関数のパラメーターの説明を示します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-172">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="d5c40-173">オプションです。</span><span class="sxs-lookup"><span data-stu-id="d5c40-173">It is optional.</span></span>

<span data-ttu-id="d5c40-174">カスタム関数内のパラメーターを省略可能と指定する方法:</span><span class="sxs-lookup"><span data-stu-id="d5c40-174">To denote a custom function parameter as optional:</span></span>

* <span data-ttu-id="d5c40-175">パラメーター名を角かっこで囲みます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-175">Put square brackets around the parameter name.</span></span> <span data-ttu-id="d5c40-176">例: `@param {string} [text] Optional text`。</span><span class="sxs-lookup"><span data-stu-id="d5c40-176">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="d5c40-177">省略可能なパラメーターの既定値は `null` です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-177">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="d5c40-178">次の例は、2 つまたは 3 つの数値を追加する ADD 関数を示しています。3 番目の数値は省略可能なパラメーターです。</span><span class="sxs-lookup"><span data-stu-id="d5c40-178">The following example shows an ADD function that adds two or three numbers, with the third number as an optional parameter.</span></span>

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

#### <a name="typescript"></a><span data-ttu-id="d5c40-179">TypeScript</span><span class="sxs-lookup"><span data-stu-id="d5c40-179">TypeScript</span></span>

<span data-ttu-id="d5c40-180">TypeScript 構文: @param 名 _の説明_</span><span class="sxs-lookup"><span data-stu-id="d5c40-180">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="d5c40-181">`name` タグが適用されるパラメーター@param指定します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-181">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="d5c40-182">必須です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-182">It is required.</span></span>
* <span data-ttu-id="d5c40-183">`description` は、Excel で表示される関数のパラメーターの説明を示します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-183">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="d5c40-184">オプションです。</span><span class="sxs-lookup"><span data-stu-id="d5c40-184">It is optional.</span></span>

<span data-ttu-id="d5c40-185">使用できる関数のパラメーターの型に関する詳細については、「[型](#types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d5c40-185">See the [Types](#types) section for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="d5c40-186">カスタム関数のパラメーターを省略可能として示すには、以下のいずれかを実行します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-186">To denote a custom function parameter as optional, do one of the following:</span></span>

* <span data-ttu-id="d5c40-187">省略可能なパラメーターを使用する。</span><span class="sxs-lookup"><span data-stu-id="d5c40-187">Use an optional parameter.</span></span> <span data-ttu-id="d5c40-188">例: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="d5c40-188">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="d5c40-189">パラメーターに既定値を指定する。</span><span class="sxs-lookup"><span data-stu-id="d5c40-189">Give the parameter a default value.</span></span> <span data-ttu-id="d5c40-190">例: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="d5c40-190">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="d5c40-191">@param の詳しい説明については、「[JSDoc](https://jsdoc.app/tags-param.html)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d5c40-191">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="d5c40-192">省略可能なパラメーターの既定値は `null` です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-192">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="d5c40-193">次の例は、2 つの数値を加算する `add` 関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="d5c40-193">The following example shows the `add` function that adds two numbers.</span></span>

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

<a id="requiresAddress"></a>

### <a name="requiresaddress"></a><span data-ttu-id="d5c40-194">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="d5c40-194">@requiresAddress</span></span>

<span data-ttu-id="d5c40-195">関数が評価されているセルのアドレスを指定する必要があることを示します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-195">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="d5c40-196">最後の関数パラメーターは、使用する型 `CustomFunctions.Invocation` または派生型である必要があります `@requiresAddress` 。</span><span class="sxs-lookup"><span data-stu-id="d5c40-196">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type to use `@requiresAddress`.</span></span> <span data-ttu-id="d5c40-197">関数が呼び出されると、`address` プロパティにアドレスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-197">When the function is called, the `address` property will contain the address.</span></span>

<span data-ttu-id="d5c40-198">次のサンプルは、パラメーターを組み合わせて使用して、カスタム関数を呼び出したセルのアドレス `invocation` `@requiresAddress` を返す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d5c40-198">The following sample shows how to use the `invocation` parameter in combination with `@requiresAddress` to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="d5c40-199">詳細については [、「呼び出しパラメーター](custom-functions-parameter-options.md#invocation-parameter) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d5c40-199">See [Invocation parameter](custom-functions-parameter-options.md#invocation-parameter) for more information.</span></span>

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<a id="requiresParameterAddresses"></a>
### <a name="requiresparameteraddresses"></a><span data-ttu-id="d5c40-200">@requiresParameterAddresses</span><span class="sxs-lookup"><span data-stu-id="d5c40-200">@requiresParameterAddresses</span></span>

<span data-ttu-id="d5c40-201">関数が入力パラメーターのアドレスを返す必要があります。</span><span class="sxs-lookup"><span data-stu-id="d5c40-201">Indicates that the function should return the addresses of input parameters.</span></span> 

<span data-ttu-id="d5c40-202">最後の関数パラメーターは、使用する型 `CustomFunctions.Invocation` または派生型である必要があります  `@requiresParameterAddresses` 。</span><span class="sxs-lookup"><span data-stu-id="d5c40-202">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type to use  `@requiresParameterAddresses`.</span></span> <span data-ttu-id="d5c40-203">JSDoc コメントには、戻り値を行列として指定するタグも含 `@returns` める `@returns {string[][]}` 必要があります `@returns {number[][]}` 。</span><span class="sxs-lookup"><span data-stu-id="d5c40-203">The JSDoc comment must also include an `@returns` tag specifying that the return value be a matrix, such as `@returns {string[][]}` or `@returns {number[][]}`.</span></span> <span data-ttu-id="d5c40-204">詳細については [、「Matrix 型](/office/dev/add-ins/excel/custom-functions-json-autogeneration#matrix-type) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d5c40-204">See [Matrix types](/office/dev/add-ins/excel/custom-functions-json-autogeneration#matrix-type) for additional information.</span></span> 

<span data-ttu-id="d5c40-205">関数が呼び出された場合、 `parameterAddresses` プロパティには入力パラメーターのアドレスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-205">When the function is called, the `parameterAddresses` property will contain the addresses of the input parameters.</span></span>

<span data-ttu-id="d5c40-206">次のサンプルは、3 つの入力パラメーターのアドレスを返す場合と組み合わせてパラメーターを使用 `invocation` `@requiresParameterAddresses` する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d5c40-206">The following sample shows how to use the `invocation` parameter in combination with `@requiresParameterAddresses` to return the addresses of three input parameters.</span></span> <span data-ttu-id="d5c40-207">詳細 [については、「パラメーターのアドレスを検出する](/office/dev/add-ins/excel/custom-functions-parameter-options#detect-the-address-of-a-parameter) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d5c40-207">See [Detect the address of a parameter](/office/dev/add-ins/excel/custom-functions-parameter-options#detect-the-address-of-a-parameter) for more information.</span></span> 

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array.
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<a id="returns"></a>
### <a name="returns"></a><span data-ttu-id="d5c40-208">@returns</span><span class="sxs-lookup"><span data-stu-id="d5c40-208">@returns</span></span>

<span data-ttu-id="d5c40-209">構文: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="d5c40-209">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="d5c40-210">戻り値の型を指定します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-210">Provides the type for the return value.</span></span>

<span data-ttu-id="d5c40-211">`{type}` を省略すると、TypeScript の型情報が使用されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-211">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="d5c40-212">型情報がない場合、型は `any` になります。</span><span class="sxs-lookup"><span data-stu-id="d5c40-212">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="d5c40-213">次の例は、 `@returns` タグを使用する `add` 関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="d5c40-213">The following example shows the `add` function that uses the `@returns` tag.</span></span>

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

<a id="streaming"></a>
### <a name="streaming"></a><span data-ttu-id="d5c40-214">@streaming</span><span class="sxs-lookup"><span data-stu-id="d5c40-214">@streaming</span></span>

<span data-ttu-id="d5c40-215">カスタム関数がストリーミング関数であることを示すのに使用されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-215">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="d5c40-216">最後のパラメーターは型です `CustomFunctions.StreamingInvocation<ResultType>` 。</span><span class="sxs-lookup"><span data-stu-id="d5c40-216">The last parameter is of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="d5c40-217">この関数は、 を返します `void` 。</span><span class="sxs-lookup"><span data-stu-id="d5c40-217">The function returns `void`.</span></span>

<span data-ttu-id="d5c40-218">ストリーミング関数は値を直接返すのではなく、最後のパラメーターを使用 `setResult(result: ResultType)` して呼び出します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-218">Streaming functions don't return values directly, instead they call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="d5c40-219">ストリーム関数によってスローされる例外は無視されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-219">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="d5c40-220">`setResult()` が、エラー結果を示すために、Error により呼び出されることがあります。</span><span class="sxs-lookup"><span data-stu-id="d5c40-220">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="d5c40-221">ストリーミング関数と詳細については、「[ストリーミング関数を作成する](custom-functions-web-reqs.md#make-a-streaming-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d5c40-221">For an example of a streaming function and more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="d5c40-222">ストリーミング関数は、[@volatile](#volatile) としてマークできません。</span><span class="sxs-lookup"><span data-stu-id="d5c40-222">Streaming functions can't be marked as [@volatile](#volatile).</span></span>

<a id="volatile"></a>
### <a name="volatile"></a><span data-ttu-id="d5c40-223">@volatile</span><span class="sxs-lookup"><span data-stu-id="d5c40-223">@volatile</span></span>

<span data-ttu-id="d5c40-224">揮発性関数とは、引数を取らない場合や引数が変更されていない場合でも、ある瞬間と次の瞬間では結果が異なる関数です。</span><span class="sxs-lookup"><span data-stu-id="d5c40-224">A volatile function is one whose result isn't the same from one moment to the next, even if it takes no arguments or the arguments haven't changed.</span></span> <span data-ttu-id="d5c40-225">Excel では、再計算が実行される度に、揮発性関数を含むセルはすべての参照先と共に、再評価されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-225">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="d5c40-226">このため、揮発性関数を多用し過ぎると再計算にかかる時間が長くなる可能性があるため、多用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="d5c40-226">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="d5c40-227">ストリーミング関数に揮発性関数は使用できません。</span><span class="sxs-lookup"><span data-stu-id="d5c40-227">Streaming functions can't be volatile.</span></span>

<span data-ttu-id="d5c40-228">次の関数は揮発性で、 `@volatile` タグを使用します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-228">The following function is volatile and uses the `@volatile` tag.</span></span>

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

## <a name="types"></a><span data-ttu-id="d5c40-229">型</span><span class="sxs-lookup"><span data-stu-id="d5c40-229">Types</span></span>

<span data-ttu-id="d5c40-230">パラメーターの型を指定すると、Excel は値を指定した型に変換してから関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-230">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="d5c40-231">型が`any`の場合、変換は実行されません。</span><span class="sxs-lookup"><span data-stu-id="d5c40-231">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="d5c40-232">値の型</span><span class="sxs-lookup"><span data-stu-id="d5c40-232">Value types</span></span>

<span data-ttu-id="d5c40-233">1 つの値は、`boolean`、 `number`、`string`の型のいずれかを使用して表現できます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-233">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="d5c40-234">マトリックス型</span><span class="sxs-lookup"><span data-stu-id="d5c40-234">Matrix type</span></span>

<span data-ttu-id="d5c40-235">2 次元配列型を使用して、パラメーターまたは戻り値を値のマトリックスにします。</span><span class="sxs-lookup"><span data-stu-id="d5c40-235">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="d5c40-236">たとえば、型は `number[][]` 数値の行列を示し、 `string[][]` 文字列の行列を示します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-236">For example, the type `number[][]` indicates a matrix of numbers and `string[][]` indicates a matrix of strings.</span></span>

### <a name="error-type"></a><span data-ttu-id="d5c40-237">エラーの種類</span><span class="sxs-lookup"><span data-stu-id="d5c40-237">Error type</span></span>

<span data-ttu-id="d5c40-238">非ストリーミング関数は、エラーの種類を返すことによりエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-238">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="d5c40-239">ストリーミング関数は、エラーの種類で `setResult()` を返してエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-239">A streaming function can indicate an error by calling `setResult()` with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="d5c40-240">Promise</span><span class="sxs-lookup"><span data-stu-id="d5c40-240">Promise</span></span>

<span data-ttu-id="d5c40-241">カスタム関数は、約束が解決された場合に値を提供する約束を返します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-241">A custom function can return a promise that provides the value when the promise is resolved.</span></span> <span data-ttu-id="d5c40-242">約束が拒否された場合、カスタム関数はエラーをスローします。</span><span class="sxs-lookup"><span data-stu-id="d5c40-242">If the promise is rejected, then the custom function will throw an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="d5c40-243">その他の型</span><span class="sxs-lookup"><span data-stu-id="d5c40-243">Other types</span></span>

<span data-ttu-id="d5c40-244">その他の型は、エラーとして処理されます。</span><span class="sxs-lookup"><span data-stu-id="d5c40-244">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="d5c40-245">次の手順</span><span class="sxs-lookup"><span data-stu-id="d5c40-245">Next steps</span></span>

<span data-ttu-id="d5c40-246">[カスタム関数用の命名規則](custom-functions-naming.md)について説明します。</span><span class="sxs-lookup"><span data-stu-id="d5c40-246">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="d5c40-247">または、[JSON ファイルを手で書く](custom-functions-json.md)必要のある[機能をローカライズする](custom-functions-localize.md)方法を確認してください。</span><span class="sxs-lookup"><span data-stu-id="d5c40-247">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d5c40-248">関連項目</span><span class="sxs-lookup"><span data-stu-id="d5c40-248">See also</span></span>

* [<span data-ttu-id="d5c40-249">カスタム関数の JSON メタデータを手動で作成する</span><span class="sxs-lookup"><span data-stu-id="d5c40-249">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d5c40-250">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="d5c40-250">Create custom functions in Excel</span></span>](custom-functions-overview.md)
