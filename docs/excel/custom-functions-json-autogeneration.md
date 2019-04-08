---
ms.date: 04/03/2019
description: JSDOC タグを使用して、カスタム関数の JSON メタデータを動的に作成します。
title: カスタム関数の JSON メタデータを作成する (プレビュー)
localization_priority: Priority
ms.openlocfilehash: c6d89684da2d0773ccfb1763e5e3e426e647523b
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/04/2019
ms.locfileid: "31478963"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a><span data-ttu-id="bd7c9-103">カスタム関数の JSON メタデータを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="bd7c9-103">Create JSON metadata for custom functions (preview)</span></span>

<span data-ttu-id="bd7c9-104">Excel カスタム関数が JavaScript または TypeScript で書き込まれている場合、JSDoc タグを使用するとカスタム関数に関する追加の情報が得られます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="bd7c9-105">JSDoc タグはビルド時に使用して、[JSON メタデータ ファイル](custom-functions-json.md)を作成します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="bd7c9-106">JSDoc タグを使用して、JSON メタデータ ファイルを手動で編集してから作業内容を保存します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

<span data-ttu-id="bd7c9-107">JavaScript または TypeScript 関数のコード コメントに`@customfunction`タグを追加して、カスタム関数としてマークします。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="bd7c9-108">関数パラメーターの型は、JavaScript の[@param](#param)タグ、または TypeScript の[関数の型](http://www.typescriptlang.org/docs/handbook/functions.html)から指定されることがあります。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](http://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="bd7c9-109">詳細については、「[@param](#param)タグと[型](#Types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-109">For more information, see the [@param](#param) tag and [Types](#Types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="bd7c9-110">JSDoc タグ</span><span class="sxs-lookup"><span data-stu-id="bd7c9-110">JSDoc Tags</span></span>
<span data-ttu-id="bd7c9-111">Excel カスタム関数では、次の JSDoc タグがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [@cancelable](#cancelable)
* <span data-ttu-id="bd7c9-112">[@customfunction](#customfunction) id 名</span><span class="sxs-lookup"><span data-stu-id="bd7c9-112">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="bd7c9-113">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="bd7c9-113">URL</span></span>
* <span data-ttu-id="bd7c9-114">[@param](#param) _{type}_ 名前の説明</span><span class="sxs-lookup"><span data-stu-id="bd7c9-114">[@param](#param) _{type}_ name description</span></span>
* [@requiresAddress](#requiresAddress)
* <span data-ttu-id="bd7c9-115">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="bd7c9-115">Type</span></span>
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@cancelable
<a id="cancelable"/>

<span data-ttu-id="bd7c9-116">関数がキャンセルされた場合は、カスタム関数がアクションを実行する必要があるということです。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-116">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="bd7c9-117">最後の関数パラメーターは`CustomFunctions.CancelableInvocation`の型にしてください。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-117">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="bd7c9-118">関数がキャンセルされた場合、関数は`oncanceled`プロパティに関数を割り当て、実行するアクションを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-118">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="bd7c9-119">最後の関数のパラメーターが`CustomFunctions.CancelableInvocation`の型である場合、タグが存在しない場合でも`@cancelable`と見なされます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-119">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="bd7c9-120">関数には`@cancelable`と`@streaming`のタグの両方を含めることはできません。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-120">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

<span data-ttu-id="bd7c9-121">構文: @customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="bd7c9-121">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="bd7c9-122">このタグを指定すると、Excel のカスタム関数として JavaScript または TypeScript の関数を処理できます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-122">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="bd7c9-123">このタグには、カスタム関数のメタデータを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-123">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="bd7c9-124">次への呼び出しもあります</span><span class="sxs-lookup"><span data-stu-id="bd7c9-124">There should also be a call to</span></span> `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a><span data-ttu-id="bd7c9-125">id</span><span class="sxs-lookup"><span data-stu-id="bd7c9-125">id</span></span> 

<span data-ttu-id="bd7c9-126">この id は、文書に保存されているカスタム関数の不変の識別子として使用されます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-126">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="bd7c9-127">変更する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-127">It should not change.</span></span>

* <span data-ttu-id="bd7c9-128">id が提供されていない場合、JavaScript または TypeScript の関数名は大文字に変換され、許可されていない文字は削除されます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-128">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="bd7c9-129">id はすべてのカスタム関数に一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-129">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="bd7c9-130">使用できる文字は、A～Z、a～z、0～9、ピリオド (.) に制限されています。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-130">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="bd7c9-131">name</span><span class="sxs-lookup"><span data-stu-id="bd7c9-131">name</span></span>

<span data-ttu-id="bd7c9-132">カスタム関数の表示名を提供します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-132">Provides the display name of a custom category for the property.</span></span> 

* <span data-ttu-id="bd7c9-133">name が指定されていない場合、id も名前として使用します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-133">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="bd7c9-134">使用できる文字は、文字[Unicode アルファベット](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、ピリオド (.)、アンダースコア (\_)です。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-134">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="bd7c9-135">最初は文字にしてください。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-135">Must begin with a letter.</span></span>
* <span data-ttu-id="bd7c9-136">最大文字数は 128 文字です。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-136">Maximum length is 255 characters.</span></span>

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

<span data-ttu-id="bd7c9-137">構文: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="bd7c9-137">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="bd7c9-138">指定された _url_ が Excel で表示されます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-138">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="bd7c9-139">JavaScript</span><span class="sxs-lookup"><span data-stu-id="bd7c9-139">JavaScript</span></span>

<span data-ttu-id="bd7c9-140">JavaScript 構文: @param{type} 名_の説明_</span><span class="sxs-lookup"><span data-stu-id="bd7c9-140">JavaScript Syntax: @param {type} name _description_</span></span>

* `{type}` <span data-ttu-id="bd7c9-141">中かっこ内の型の情報を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-141">should specify the type info within curly braces.</span></span> <span data-ttu-id="bd7c9-142">使用する可能性のある型に関する詳細については、「[型](##types)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-142">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="bd7c9-143">オプション: 指定しない場合、`any`の型を使用します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-143">Optional: if not specified, the type `any` will be used.</span></span>
* `name` <span data-ttu-id="bd7c9-144">@paramのタグを適用するパラメーターを指定します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-144">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="bd7c9-145">必須です。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-145">Required.</span></span>
* `description` <span data-ttu-id="bd7c9-146">関数のパラメーターの Excel で表示される説明を取得できます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-146">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="bd7c9-147">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-147">Optional.</span></span>

<span data-ttu-id="bd7c9-148">オプションとしてカスタム関数のパラメーターを示すには:</span><span class="sxs-lookup"><span data-stu-id="bd7c9-148">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="bd7c9-149">パラメーター名を角かっこで囲みます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-149">Put square brackets around the parameter name.</span></span> <span data-ttu-id="bd7c9-150">例: `@param {string} [text] Optional text`。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-150">For example: `@param {string} [text] Optional text`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="bd7c9-151">TypeScript</span><span class="sxs-lookup"><span data-stu-id="bd7c9-151">TypeScript</span></span>

<span data-ttu-id="bd7c9-152">TypeScript 構文: @param名前_の説明_</span><span class="sxs-lookup"><span data-stu-id="bd7c9-152">TypeScript Syntax: @param name _description_</span></span>

* `name` <span data-ttu-id="bd7c9-153">@paramのタグを適用するパラメーターを指定します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-153">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="bd7c9-154">必須です。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-154">Required.</span></span>
* `description` <span data-ttu-id="bd7c9-155">関数のパラメーターの Excel で表示される説明を取得できます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-155">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="bd7c9-156">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-156">Optional.</span></span>

<span data-ttu-id="bd7c9-157">使用する可能性のある関数のパラメーターの型に関する詳細については、「[型](##types)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-157">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="bd7c9-158">オプションとしてカスタム関数のパラメーターを示すには、以下のいずれかを実行します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-158">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="bd7c9-159">オプションのパラメーターを使用します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-159">Use an optional parameter.</span></span> <span data-ttu-id="bd7c9-160">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-160">For example:</span></span> `function f(text?: string)`
* <span data-ttu-id="bd7c9-161">パラメーターに既定値を指定します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-161">Give the parameter a default value.</span></span> <span data-ttu-id="bd7c9-162">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-162">For example:</span></span> `function f(text: string = "abc")`

<span data-ttu-id="bd7c9-163">@paramの詳しい説明については、「[JSDoc](http://usejsdoc.org/tags-param.html)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-163">For a detailed description of the code, see "HelloData Details."</span></span>

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

<span data-ttu-id="bd7c9-164">機能が評価されているセルのアドレスを指定する必要があることを示しています。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-164">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="bd7c9-165">最後の関数のパラメーターは、`CustomFunctions.Invocation`の型または派生型にしてください。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-165">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="bd7c9-166">関数が呼び出される場合、`address`プロパティにはアドレスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-166">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a>@returns
<a id="returns"/>

<span data-ttu-id="bd7c9-167">構文: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="bd7c9-167">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="bd7c9-168">戻り値の型を指定します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-168">Provides the type for the return value.</span></span>

<span data-ttu-id="bd7c9-169">`{type}`を省略すると、TypeScript 型の情報が使用されます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-169">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="bd7c9-170">型の情報がない場合、`any`になります。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-170">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

<span data-ttu-id="bd7c9-171">カスタム関数がストリーミング関数であることを示すのに使用されます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-171">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="bd7c9-172">最後のパラメーターは`CustomFunctions.StreamingInvocation<ResultType>`型のいずれかにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-172">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="bd7c9-173">関数は`void`を返す必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-173">The function should return `void`.</span></span>

<span data-ttu-id="bd7c9-174">ストリーミング関数は直接値を返しませんが、代わりに最後のパラメーターを使用して`setResult(result: ResultType)`を呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-174">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="bd7c9-175">ストリーム関数によってスローされる例外は無視されます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-175">Exceptions thrown by a streaming function are ignored.</span></span> `setResult()` <span data-ttu-id="bd7c9-176">エラー結果を示すためにエラーで呼び出されることがあります。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-176">may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="bd7c9-177">ストリーミング関数は[@volatile](#volatile)としてマークされます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-177">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

<span data-ttu-id="bd7c9-178">揮発性関数とは、引数を取らない場合や引数が変更されていない場合でも、結果が頻繁に同じになるとは想定できないものです。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-178">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="bd7c9-179">Excel は再計算が実行される度に、揮発性関数が含まれているセルをすべての参照先と共に再評価します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-179">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="bd7c9-180">このため、揮発性関数に依存しすぎていると、再計算にかかる時間が長くなる可能性があるため、使いすぎないようにしてください。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-180">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="bd7c9-181">ストリーミング関数を揮発性関数にはできません。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-181">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="bd7c9-182">型</span><span class="sxs-lookup"><span data-stu-id="bd7c9-182">Types</span></span>

<span data-ttu-id="bd7c9-183">パラメーターの型を指定すると、Excel は値を指定した型に変換してから関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-183">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="bd7c9-184">型が`any`の場合、変換は実行されません。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-184">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="bd7c9-185">値の型</span><span class="sxs-lookup"><span data-stu-id="bd7c9-185">Value types</span></span>

<span data-ttu-id="bd7c9-186">1 つの値は、`boolean`、 `number`、`string`の型のいずれかを使用して表現できる場合があります。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-186">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="bd7c9-187">マトリックスの型</span><span class="sxs-lookup"><span data-stu-id="bd7c9-187">Matrix type</span></span>

<span data-ttu-id="bd7c9-188">2 次元配列型を使用して、パラメーターまたは戻り値をマトリックス値にします。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-188">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="bd7c9-189">たとえば、`number[][]`の型は数字のマトリックスを示します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-189">For example, the type `number[][]` indicates a matrix of numbers.</span></span> `string[][]` <span data-ttu-id="bd7c9-190">文字列のマトリックスを示します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-190">indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="bd7c9-191">エラーの種類</span><span class="sxs-lookup"><span data-stu-id="bd7c9-191">Error Type</span></span>

<span data-ttu-id="bd7c9-192">非ストリーミング関数は、エラーの種類を返してエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-192">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="bd7c9-193">ストリーミング関数は、エラーの種類で setResult() を返してエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-193">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="bd7c9-194">Promise</span><span class="sxs-lookup"><span data-stu-id="bd7c9-194">Promise object.</span></span>

<span data-ttu-id="bd7c9-195">関数は Promise を返すことができ、Promise が解決される場合に値を指定します。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-195">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="bd7c9-196">Promise が拒否される場合は、エラーになっています。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-196">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="bd7c9-197">その他の型</span><span class="sxs-lookup"><span data-stu-id="bd7c9-197">Other solution types</span></span>

<span data-ttu-id="bd7c9-198">その他の型はエラーとして処理されます。</span><span class="sxs-lookup"><span data-stu-id="bd7c9-198">Any other type will be treated as an error.</span></span>

## <a name="see-also"></a><span data-ttu-id="bd7c9-199">関連項目</span><span class="sxs-lookup"><span data-stu-id="bd7c9-199">See also</span></span>

* [<span data-ttu-id="bd7c9-200">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="bd7c9-200">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="bd7c9-201">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="bd7c9-201">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="bd7c9-202">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="bd7c9-202">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="bd7c9-203">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="bd7c9-203">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="bd7c9-204">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="bd7c9-204">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="bd7c9-205">カスタム関数のデバッグ</span><span class="sxs-lookup"><span data-stu-id="bd7c9-205">Custom functions debugging</span></span>](custom-functions-debugging.md)
