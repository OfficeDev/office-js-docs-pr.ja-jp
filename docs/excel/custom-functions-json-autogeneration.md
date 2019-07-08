---
ms.date: 06/27/2019
description: JSDoc タグを使用して、カスタム関数の JSON メタデータを動的に作成します。
title: カスタム関数用の JSON メタデータの自動生成
localization_priority: Priority
ms.openlocfilehash: 1230e1bfdeead306531a218373c2756b29fa4abe
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454658"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="86bd3-103">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="86bd3-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="86bd3-104">Excel カスタム関数が JavaScript または TypeScript で記述されている場合、カスタム関数に関する追加の情報を提供するために、JSDoc タグが使用されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="86bd3-105">JSDoc タグはビルド時に使用して、[JSON メタデータ ファイル](custom-functions-json.md)を作成します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="86bd3-106">JSDoc タグを使用すると、JSON メタデータ ファイルを手動で編集する手間が省けます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="86bd3-107">JavaScript または TypeScript 関数のコード コメントに`@customfunction`タグを追加して、カスタム関数としてマークします。</span><span class="sxs-lookup"><span data-stu-id="86bd3-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="86bd3-108">関数パラメーターの型は、JavaScript の [@param](#param) タグを使用して指定するか、TypeScript の[関数の型](https://www.typescriptlang.org/docs/handbook/functions.html)から指定できます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="86bd3-109">詳細については、「[@param](#param) タグ」セクションと「[型](#types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="86bd3-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="86bd3-110">関数に説明を追加する</span><span class="sxs-lookup"><span data-stu-id="86bd3-110">Adding a description to a function</span></span>

<span data-ttu-id="86bd3-111">説明は、カスタム関数の機能を理解するためのヘルプが必要な場合に、ヘルプ テキストとしてユーザーに表示されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="86bd3-112">説明に特定のタグは必要ありません。</span><span class="sxs-lookup"><span data-stu-id="86bd3-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="86bd3-113">JSDoc コメントに簡単な説明を入力するだけです。</span><span class="sxs-lookup"><span data-stu-id="86bd3-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="86bd3-114">一般に、説明は JSDoc コメント セクションの先頭に配置されますが、配置場所に関係なく機能します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="86bd3-115">組み込み関数の説明の例を表示するには、Excel を開き、**[数式]** タブに移動し、**[関数の​​挿入]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="86bd3-116">すべての関数の説明を参照したり、独自のカスタム関数を一覧表示したりすることができます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="86bd3-117">次の例では、「球の体積を計算します。」</span><span class="sxs-lookup"><span data-stu-id="86bd3-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="86bd3-118">が、カスタム関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="86bd3-119">JSDoc タグ</span><span class="sxs-lookup"><span data-stu-id="86bd3-119">JSDoc Tags</span></span>
<span data-ttu-id="86bd3-120">Excel カスタム関数では、次の JSDoc タグを利用できます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-120">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="86bd3-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="86bd3-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="86bd3-122">[@customfunction](#customfunction) id 名</span><span class="sxs-lookup"><span data-stu-id="86bd3-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="86bd3-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="86bd3-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="86bd3-124">[@param](#param) _{type}_ 名前の説明</span><span class="sxs-lookup"><span data-stu-id="86bd3-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="86bd3-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="86bd3-125">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="86bd3-126">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="86bd3-126">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="86bd3-127">@streaming</span><span class="sxs-lookup"><span data-stu-id="86bd3-127">@streaming</span></span>](#streaming)
* [<span data-ttu-id="86bd3-128">@volatile</span><span class="sxs-lookup"><span data-stu-id="86bd3-128">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="86bd3-129">@cancelable</span><span class="sxs-lookup"><span data-stu-id="86bd3-129">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="86bd3-130">関数がキャンセルされた場合にカスタム関数がアクションを実行することを示します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-130">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="86bd3-131">最後の関数パラメーターは `CustomFunctions.CancelableInvocation` の型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="86bd3-131">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="86bd3-132">関数は `oncanceled` プロパティに関数を割り当て、関数がキャンセルされた場合に実行するアクションを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-132">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="86bd3-133">最後の関数のパラメーターが `CustomFunctions.CancelableInvocation` 型の場合、タグは表示されませんが、`@cancelable` と見なされます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-133">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="86bd3-134">関数には `@cancelable` と `@streaming` の両方のタグを含めることはできません。</span><span class="sxs-lookup"><span data-stu-id="86bd3-134">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="86bd3-135">@customfunction</span><span class="sxs-lookup"><span data-stu-id="86bd3-135">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="86bd3-136">構文: @customfunction _id_ _名_</span><span class="sxs-lookup"><span data-stu-id="86bd3-136">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="86bd3-137">このタグを指定すると、JavaScript または TypeScript の関数を、Excel のカスタム関数として処理できます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-137">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span> 

<span data-ttu-id="86bd3-138">このタグは、カスタム関数のメタデータを作成するために必要です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-138">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="86bd3-139">次への呼び出しもあります: `CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="86bd3-139">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

<span data-ttu-id="86bd3-140">次の例は、カスタム関数を宣言する最も簡単な方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="86bd3-140">The following example shows the simplest way to declare a custom function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="86bd3-141">ID</span><span class="sxs-lookup"><span data-stu-id="86bd3-141">id</span></span>

<span data-ttu-id="86bd3-142">`id` は、カスタム関数の不変識別子です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-142">The id is used as the invariant identifier for the custom function stored in the document.</span></span>

* <span data-ttu-id="86bd3-143">`id` が提供されていない場合、JavaScript または TypeScript の関数名は大文字に変換され、許可されない文字は削除されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-143">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="86bd3-144">`id` はすべてのカスタム関数で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="86bd3-144">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="86bd3-145">指定できる文字は、A から Z、a から z、0 から 9、アンダースコア (\_)、ピリオド (.) に制限されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-145">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

<span data-ttu-id="86bd3-146">次の例では、インクリメントは関数の `id` と `name` です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-146">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="86bd3-147">name</span><span class="sxs-lookup"><span data-stu-id="86bd3-147">name</span></span>

<span data-ttu-id="86bd3-148">カスタム関数の表示用の `name` を提供します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-148">Provides the display name for the custom function.</span></span>

* <span data-ttu-id="86bd3-149">name が指定されていない場合、id が名前としても使用されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-149">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="86bd3-150">使用できる文字は、文字 [Unicode アルファベット](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、ピリオド (.)、およびアンダースコア (\_)です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-150">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="86bd3-151">最初の文字は、アルファベット文字にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="86bd3-151">Must start with a letter.</span></span>
* <span data-ttu-id="86bd3-152">最大文字数は 128 文字です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-152">Maximum length is 128 characters.</span></span>

<span data-ttu-id="86bd3-153">次の例では、INC は関数の`id` で、 `increment` は`name`です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-153">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="86bd3-154">説明</span><span class="sxs-lookup"><span data-stu-id="86bd3-154">description</span></span>

<span data-ttu-id="86bd3-155">説明に特定のタグは必要ありません。</span><span class="sxs-lookup"><span data-stu-id="86bd3-155">A description doesn't require any specific tag.</span></span> <span data-ttu-id="86bd3-156">JSDoc コメント内に関数の機能を説明するフレーズを入力して、カスタム関数に説明を追加します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-156">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="86bd3-157">既定では、JSDoc コメント セクションでタグが付けられていないテキストは、関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-157">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span> <span data-ttu-id="86bd3-158">Excel では、関数の入力時に、ユーザーにこの説明が表示されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-158">The description appears to users in Excel as they are entering the function.</span></span> <span data-ttu-id="86bd3-159">次の例では、「2 つの数値を加算する関数」というフレーズが、ID プロパティ `ADD` のカスタム関数の説明です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-159">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

<span data-ttu-id="86bd3-160">次の例では、ADD は関数の `id` と `name` で、 説明が加えられています。</span><span class="sxs-lookup"><span data-stu-id="86bd3-160">In the following example, ADD is the `id` and `name` of the function and a description is given.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
### <a name="helpurl"></a><span data-ttu-id="86bd3-161">@helpurl</span><span class="sxs-lookup"><span data-stu-id="86bd3-161">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="86bd3-162">構文: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="86bd3-162">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="86bd3-163">指定された _url_ が Excel で表示されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-163">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="86bd3-164">次の例では、 `helpurl` は www.contoso.com/weatherhelp です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-164">In the following example, the `helpurl` is www.contoso.com/weatherhelp.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
### <a name="param"></a><span data-ttu-id="86bd3-165">@param</span><span class="sxs-lookup"><span data-stu-id="86bd3-165">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="86bd3-166">JavaScript</span><span class="sxs-lookup"><span data-stu-id="86bd3-166">JavaScript</span></span>

<span data-ttu-id="86bd3-167">JavaScript 構文: @param {type} 名_の説明_</span><span class="sxs-lookup"><span data-stu-id="86bd3-167">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="86bd3-168">`{type}` は、中かっこ内の型の情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-168">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="86bd3-169">使用できる型に関する詳細については、「[型](#types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="86bd3-169">See the [Types](#types) for more information about the types which may be used.</span></span> <span data-ttu-id="86bd3-170">省略可能: 指定しない場合、`any` 型が使用されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-170">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="86bd3-171">`name` は、@param タグを適用するパラメーターを指定します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-171">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="86bd3-172">必須です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-172">Required.</span></span>
* <span data-ttu-id="86bd3-173">`description` は、Excel で表示される関数のパラメーターの説明を示します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-173">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="86bd3-174">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-174">Optional.</span></span>

<span data-ttu-id="86bd3-175">カスタム関数内のパラメーターを省略可能と指定する方法:</span><span class="sxs-lookup"><span data-stu-id="86bd3-175">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="86bd3-176">パラメーター名を角かっこで囲みます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-176">Put square brackets around the parameter name.</span></span> <span data-ttu-id="86bd3-177">例: `@param {string} [text] Optional text`。</span><span class="sxs-lookup"><span data-stu-id="86bd3-177">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="86bd3-178">省略可能なパラメーターの既定値は `null` です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-178">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="86bd3-179">次の例では、ADD 関数を使用して 2 つまたは 3 つの数値を加算します。3 番目の数値を省略可能なパラメーターとして使用します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-179">The following example shows a ADD function which adds two or three numbers, with the third number as an optional parameter.</span></span>

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

#### <a name="typescript"></a><span data-ttu-id="86bd3-180">TypeScript</span><span class="sxs-lookup"><span data-stu-id="86bd3-180">TypeScript</span></span>

<span data-ttu-id="86bd3-181">TypeScript 構文: @param 名 _の説明_</span><span class="sxs-lookup"><span data-stu-id="86bd3-181">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="86bd3-182">`name` は、@param タグを適用するパラメーターを指定します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-182">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="86bd3-183">必須です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-183">Required.</span></span>
* <span data-ttu-id="86bd3-184">`description` は、Excel で表示される関数のパラメーターの説明を示します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-184">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="86bd3-185">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-185">Optional.</span></span>

<span data-ttu-id="86bd3-186">使用できる関数のパラメーターの型に関する詳細については、「[型](#types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="86bd3-186">See the [Types](#types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="86bd3-187">カスタム関数のパラメーターを省略可能として示すには、以下のいずれかを実行します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-187">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="86bd3-188">省略可能なパラメーターを使用する。</span><span class="sxs-lookup"><span data-stu-id="86bd3-188">Use an optional parameter.</span></span> <span data-ttu-id="86bd3-189">例: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="86bd3-189">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="86bd3-190">パラメーターに既定値を指定する。</span><span class="sxs-lookup"><span data-stu-id="86bd3-190">Give the parameter a default value.</span></span> <span data-ttu-id="86bd3-191">例: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="86bd3-191">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="86bd3-192">@param の詳しい説明については、「[JSDoc](https://jsdoc.app/tags-param.html)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="86bd3-192">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="86bd3-193">省略可能なパラメーターの既定値は `null` です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-193">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="86bd3-194">次の例は、2 つの数値を加算する `add` 関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="86bd3-194">The following example shows the `add` function that adds two numbers.</span></span>

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
### <a name="requiresaddress"></a><span data-ttu-id="86bd3-195">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="86bd3-195">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="86bd3-196">関数が評価されているセルのアドレスを指定する必要があることを示します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-196">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="86bd3-197">最後の関数のパラメーターは、`CustomFunctions.Invocation` 型または派生型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="86bd3-197">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="86bd3-198">関数が呼び出されると、`address` プロパティにアドレスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-198">When the function is called, the `address` property will contain the address.</span></span> <span data-ttu-id="86bd3-199">`@requiresAddress` タグを使用する関数の例については、 「[セルのコンテキストパラメーターを指定する](./custom-functions-parameter-options.md#addressing-cells-context-parameter)」 を参照してください。</span><span class="sxs-lookup"><span data-stu-id="86bd3-199">For an example of a function that uses the `@requiresAddress` tag, see [Addressing cell's context parameter](./custom-functions-parameter-options.md#addressing-cells-context-parameter).</span></span>

---
### <a name="returns"></a><span data-ttu-id="86bd3-200">@returns</span><span class="sxs-lookup"><span data-stu-id="86bd3-200">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="86bd3-201">構文: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="86bd3-201">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="86bd3-202">戻り値の型を指定します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-202">Provides the type for the return value.</span></span>

<span data-ttu-id="86bd3-203">`{type}` を省略すると、TypeScript の型情報が使用されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-203">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="86bd3-204">型情報がない場合、型は `any` になります。</span><span class="sxs-lookup"><span data-stu-id="86bd3-204">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="86bd3-205">次の例は、 `@returns` タグを使用する `add` 関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="86bd3-205">The following example shows the `add` function that uses the `@returns` tag.</span></span>

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
### <a name="streaming"></a><span data-ttu-id="86bd3-206">@streaming</span><span class="sxs-lookup"><span data-stu-id="86bd3-206">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="86bd3-207">カスタム関数がストリーミング関数であることを示すのに使用されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-207">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="86bd3-208">最後のパラメーターは、`CustomFunctions.StreamingInvocation<ResultType>` 型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="86bd3-208">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="86bd3-209">関数は `void` を返します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-209">The function should return `void`.</span></span>

<span data-ttu-id="86bd3-210">ストリーミング関数は値を直接返さず、代わりに、最後のパラメーターを使用して `setResult(result: ResultType)` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-210">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="86bd3-211">ストリーム関数によってスローされる例外は無視されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-211">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="86bd3-212">`setResult()` が、エラー結果を示すために、Error により呼び出されることがあります。</span><span class="sxs-lookup"><span data-stu-id="86bd3-212">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="86bd3-213">ストリーミング関数と詳細については、「[ストリーミング関数を作成する](./custom-functions-web-reqs.md#make-a-streaming-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="86bd3-213">For an example of a streaming function and more information, see [Make a streaming function](./custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="86bd3-214">ストリーミング関数は、[@volatile](#volatile) としてマークできません。</span><span class="sxs-lookup"><span data-stu-id="86bd3-214">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="86bd3-215">@volatile</span><span class="sxs-lookup"><span data-stu-id="86bd3-215">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="86bd3-216">揮発性関数とは、引数を取らない場合や引数が変更されていない場合でも、ある瞬間と次の瞬間では結果が異なる関数です。</span><span class="sxs-lookup"><span data-stu-id="86bd3-216">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="86bd3-217">Excel では、再計算が実行される度に、揮発性関数を含むセルはすべての参照先と共に、再評価されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-217">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="86bd3-218">このため、揮発性関数を多用し過ぎると再計算にかかる時間が長くなる可能性があるため、多用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="86bd3-218">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="86bd3-219">ストリーミング関数に揮発性関数は使用できません。</span><span class="sxs-lookup"><span data-stu-id="86bd3-219">Streaming functions cannot be volatile.</span></span>

<span data-ttu-id="86bd3-220">次の関数は揮発性で、 `@volatile` タグを使用します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-220">The following function is volatile and uses the `@volatile` tag.</span></span>

```js
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a><span data-ttu-id="86bd3-221">型</span><span class="sxs-lookup"><span data-stu-id="86bd3-221">Types</span></span>

<span data-ttu-id="86bd3-222">パラメーターの型を指定すると、Excel は値を指定した型に変換してから関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-222">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="86bd3-223">型が`any`の場合、変換は実行されません。</span><span class="sxs-lookup"><span data-stu-id="86bd3-223">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="86bd3-224">値の型</span><span class="sxs-lookup"><span data-stu-id="86bd3-224">Value types</span></span>

<span data-ttu-id="86bd3-225">1 つの値は、`boolean`、 `number`、`string`の型のいずれかを使用して表現できます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-225">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="86bd3-226">マトリックス型</span><span class="sxs-lookup"><span data-stu-id="86bd3-226">Matrix type</span></span>

<span data-ttu-id="86bd3-227">2 次元配列型を使用して、パラメーターまたは戻り値を値のマトリックスにします。</span><span class="sxs-lookup"><span data-stu-id="86bd3-227">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="86bd3-228">たとえば、`number[][]`の型は数字のマトリックスを示します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-228">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="86bd3-229">`string[][]` は、文字列のマトリックスを示します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-229">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="86bd3-230">エラーの種類</span><span class="sxs-lookup"><span data-stu-id="86bd3-230">Error type</span></span>

<span data-ttu-id="86bd3-231">非ストリーミング関数は、エラーの種類を返すことによりエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-231">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="86bd3-232">ストリーミング関数は、エラーの種類で `setResult()` を返してエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-232">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="86bd3-233">Promise</span><span class="sxs-lookup"><span data-stu-id="86bd3-233">Promise</span></span>

<span data-ttu-id="86bd3-234">関数は Promise を返すことができ、Promise が解決されたときに値を提供します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-234">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="86bd3-235">Promise が拒否された場合は、エラーになります。</span><span class="sxs-lookup"><span data-stu-id="86bd3-235">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="86bd3-236">その他の型</span><span class="sxs-lookup"><span data-stu-id="86bd3-236">Other types</span></span>

<span data-ttu-id="86bd3-237">その他の型は、エラーとして処理されます。</span><span class="sxs-lookup"><span data-stu-id="86bd3-237">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="86bd3-238">次の手順</span><span class="sxs-lookup"><span data-stu-id="86bd3-238">Next steps</span></span>
<span data-ttu-id="86bd3-239">[カスタム関数用の命名規則](custom-functions-naming.md)について説明します。</span><span class="sxs-lookup"><span data-stu-id="86bd3-239">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="86bd3-240">または、[JSON ファイルを手で書く](custom-functions-json.md)必要のある[機能をローカライズする](custom-functions-localize.md)方法を確認してください。</span><span class="sxs-lookup"><span data-stu-id="86bd3-240">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="86bd3-241">関連項目</span><span class="sxs-lookup"><span data-stu-id="86bd3-241">See also</span></span>

* [<span data-ttu-id="86bd3-242">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="86bd3-242">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="86bd3-243">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="86bd3-243">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="86bd3-244">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="86bd3-244">Create custom functions in Excel</span></span>](custom-functions-overview.md)
