---
ms.date: 04/03/2019
description: JSDOC タグを使用して、カスタム関数の JSON メタデータを動的に作成する。
title: カスタム関数の JSON メタデータを作成する (プレビュー)
localization_priority: Priority
ms.openlocfilehash: c6d89684da2d0773ccfb1763e5e3e426e647523b
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/04/2019
ms.locfileid: "31478963"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a><span data-ttu-id="0a720-103">カスタム関数の JSON メタデータを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="0a720-103">Create JSON metadata for custom functions (preview)</span></span>

<span data-ttu-id="0a720-104">Excel カスタム関数が JavaScript または TypeScript で記述されている場合、カスタム関数に関する追加の情報を提供するために、JSDoc タグが使用されます。</span><span class="sxs-lookup"><span data-stu-id="0a720-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="0a720-105">JSDoc タグはビルド時に使用して、[JSON メタデータ ファイル](custom-functions-json.md)を作成します。</span><span class="sxs-lookup"><span data-stu-id="0a720-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="0a720-106">JSDoc タグを使用すると、JSON メタデータ ファイルを手動で編集する手間が省けます。</span><span class="sxs-lookup"><span data-stu-id="0a720-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

<span data-ttu-id="0a720-107">JavaScript または TypeScript 関数のコード コメントに`@customfunction`タグを追加して、カスタム関数としてマークします。</span><span class="sxs-lookup"><span data-stu-id="0a720-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="0a720-108">関数パラメーターの型は、JavaScript の [@param](#param) タグを使用して指定するか、TypeScript の[関数の型](http://www.typescriptlang.org/docs/handbook/functions.html)から指定できます。</span><span class="sxs-lookup"><span data-stu-id="0a720-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](http://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="0a720-109">詳細については、「[@param](#param) タグ」セクションと「[型](#Types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0a720-109">For more information, see the [@param](#param) tag and [Types](#Types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="0a720-110">JSDoc タグ</span><span class="sxs-lookup"><span data-stu-id="0a720-110">JSDoc Tags</span></span>
<span data-ttu-id="0a720-111">Excel カスタム関数では、次の JSDoc タグを利用できます。</span><span class="sxs-lookup"><span data-stu-id="0a720-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [@cancelable](#cancelable)
* <span data-ttu-id="0a720-112">[@customfunction](#customfunction) id name</span><span class="sxs-lookup"><span data-stu-id="0a720-112">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="0a720-113">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="0a720-113">URL</span></span>
* <span data-ttu-id="0a720-114">[@param](#param) _{type}_ 名前の説明</span><span class="sxs-lookup"><span data-stu-id="0a720-114">[@param](#param) _{type}_ name description</span></span>
* [@requiresAddress](#requiresAddress)
* <span data-ttu-id="0a720-115">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="0a720-115">Type</span></span>
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@cancelable
<a id="cancelable"/>

<span data-ttu-id="0a720-116">関数がキャンセルされた場合にカスタム関数がアクションを実行することを示します。</span><span class="sxs-lookup"><span data-stu-id="0a720-116">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="0a720-117">最後の関数パラメーターは `CustomFunctions.CancelableInvocation` の型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0a720-117">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="0a720-118">関数は `oncanceled` プロパティに関数を割り当て、関数がキャンセルされた場合に実行するアクションを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="0a720-118">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="0a720-119">最後の関数のパラメーターが `CustomFunctions.CancelableInvocation` 型の場合、タグは表示されませんが、`@cancelable` と見なされます。</span><span class="sxs-lookup"><span data-stu-id="0a720-119">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="0a720-120">関数には `@cancelable` と `@streaming` の両方のタグを含めることはできません。</span><span class="sxs-lookup"><span data-stu-id="0a720-120">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

<span data-ttu-id="0a720-121">構文: @customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="0a720-121">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="0a720-122">このタグを指定すると、JavaScript または TypeScript の関数を、Excel のカスタム関数として処理できます。</span><span class="sxs-lookup"><span data-stu-id="0a720-122">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="0a720-123">このタグは、カスタム関数のメタデータを作成するために必要です。</span><span class="sxs-lookup"><span data-stu-id="0a720-123">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="0a720-124">次への呼び出しもあります: </span><span class="sxs-lookup"><span data-stu-id="0a720-124">There should also be a call to</span></span> `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a><span data-ttu-id="0a720-125">id</span><span class="sxs-lookup"><span data-stu-id="0a720-125">id</span></span> 

<span data-ttu-id="0a720-126">id は、文書に格納されているカスタム関数の不変の識別子として使用されます。</span><span class="sxs-lookup"><span data-stu-id="0a720-126">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="0a720-127">変更する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="0a720-127">It should not change.</span></span>

* <span data-ttu-id="0a720-128">id が提供されていない場合、JavaScript または TypeScript の関数名は大文字に変換され、許可されない文字は削除されます。</span><span class="sxs-lookup"><span data-stu-id="0a720-128">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="0a720-129">id はすべてのカスタム関数で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="0a720-129">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="0a720-130">使用できる文字は、A～Z、a～z、0～9、ピリオド (.) に制限されています。</span><span class="sxs-lookup"><span data-stu-id="0a720-130">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="0a720-131">name</span><span class="sxs-lookup"><span data-stu-id="0a720-131">name</span></span>

<span data-ttu-id="0a720-132">カスタム関数の表示名を提供します。</span><span class="sxs-lookup"><span data-stu-id="0a720-132">Provides the display name of a custom category for the property.</span></span> 

* <span data-ttu-id="0a720-133">name が指定されていない場合、id が名前としても使用されます。</span><span class="sxs-lookup"><span data-stu-id="0a720-133">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="0a720-134">使用できる文字は、文字 [Unicode アルファベット](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、ピリオド (.)、およびアンダースコア (\_)です。</span><span class="sxs-lookup"><span data-stu-id="0a720-134">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="0a720-135">最初の文字は、アルファベット文字にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0a720-135">Must begin with a letter.</span></span>
* <span data-ttu-id="0a720-136">最大文字数は 128 文字です。</span><span class="sxs-lookup"><span data-stu-id="0a720-136">Maximum length is 255 characters.</span></span>

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

<span data-ttu-id="0a720-137">構文: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="0a720-137">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="0a720-138">指定された _url_ が Excel で表示されます。</span><span class="sxs-lookup"><span data-stu-id="0a720-138">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="0a720-139">JavaScript</span><span class="sxs-lookup"><span data-stu-id="0a720-139">JavaScript</span></span>

<span data-ttu-id="0a720-140">JavaScript 構文: @param{type} 名_の説明_</span><span class="sxs-lookup"><span data-stu-id="0a720-140">JavaScript Syntax: @param {type} name _description_</span></span>

* `{type}` <span data-ttu-id="0a720-141">は、中かっこ内の型の情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="0a720-141">should specify the type info within curly braces.</span></span> <span data-ttu-id="0a720-142">使用できる型に関する詳細については、「[型](##types)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0a720-142">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="0a720-143">省略可能: 指定しない場合、`any` 型が使用されます。</span><span class="sxs-lookup"><span data-stu-id="0a720-143">Optional: if not specified, the type `any` will be used.</span></span>
* `name` <span data-ttu-id="0a720-144">は、タグ @param を適用するパラメーターを指定します。</span><span class="sxs-lookup"><span data-stu-id="0a720-144">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="0a720-145">必須です。</span><span class="sxs-lookup"><span data-stu-id="0a720-145">Required.</span></span>
* `description` <span data-ttu-id="0a720-146">は、Excel で表示される関数のパラメーターの説明を示します。</span><span class="sxs-lookup"><span data-stu-id="0a720-146">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="0a720-147">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="0a720-147">Optional.</span></span>

<span data-ttu-id="0a720-148">カスタム関数内のパラメーターを省略可能と指定する方法:</span><span class="sxs-lookup"><span data-stu-id="0a720-148">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="0a720-149">パラメーター名を角かっこで囲みます。</span><span class="sxs-lookup"><span data-stu-id="0a720-149">Put square brackets around the parameter name.</span></span> <span data-ttu-id="0a720-150">例: `@param {string} [text] Optional text`。</span><span class="sxs-lookup"><span data-stu-id="0a720-150">For example: `@param {string} [text] Optional text`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="0a720-151">TypeScript</span><span class="sxs-lookup"><span data-stu-id="0a720-151">TypeScript</span></span>

<span data-ttu-id="0a720-152">TypeScript 構文: @paramname_description_</span><span class="sxs-lookup"><span data-stu-id="0a720-152">TypeScript Syntax: @param name _description_</span></span>

* `name` <span data-ttu-id="0a720-153">は、タグ @param を適用するパラメーターを指定します。</span><span class="sxs-lookup"><span data-stu-id="0a720-153">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="0a720-154">必須です。</span><span class="sxs-lookup"><span data-stu-id="0a720-154">Required.</span></span>
* `description` <span data-ttu-id="0a720-155">は、Excel で表示される関数のパラメーターの説明を示します。</span><span class="sxs-lookup"><span data-stu-id="0a720-155">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="0a720-156">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="0a720-156">Optional.</span></span>

<span data-ttu-id="0a720-157">使用できる関数のパラメーターの型に関する詳細については、「[型](##types)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0a720-157">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="0a720-158">カスタム関数のパラメーターを省略可能として示すには、以下のいずれかを実行します。</span><span class="sxs-lookup"><span data-stu-id="0a720-158">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="0a720-159">省略可能なパラメーターを使用する。</span><span class="sxs-lookup"><span data-stu-id="0a720-159">Use an optional parameter.</span></span> <span data-ttu-id="0a720-160">例: </span><span class="sxs-lookup"><span data-stu-id="0a720-160">For example:</span></span> `function f(text?: string)`
* <span data-ttu-id="0a720-161">パラメーターに既定値を指定する。</span><span class="sxs-lookup"><span data-stu-id="0a720-161">Give the parameter a default value.</span></span> <span data-ttu-id="0a720-162">例: </span><span class="sxs-lookup"><span data-stu-id="0a720-162">For example:</span></span> `function f(text: string = "abc")`

<span data-ttu-id="0a720-163">@paramの詳しい説明については、「[JSDoc](http://usejsdoc.org/tags-param.html)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0a720-163">For a detailed description of the code, see "HelloData Details."</span></span>

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

<span data-ttu-id="0a720-164">関数が評価されているセルのアドレスを指定する必要があることを示します。</span><span class="sxs-lookup"><span data-stu-id="0a720-164">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="0a720-165">最後の関数のパラメーターは、`CustomFunctions.Invocation` 型または派生型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0a720-165">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="0a720-166">関数が呼び出されると、`address` プロパティにアドレスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0a720-166">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a>@returns
<a id="returns"/>

<span data-ttu-id="0a720-167">構文: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="0a720-167">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="0a720-168">戻り値の型を指定します。</span><span class="sxs-lookup"><span data-stu-id="0a720-168">Provides the type for the return value.</span></span>

<span data-ttu-id="0a720-169">`{type}` を省略すると、TypeScript の型情報が使用されます。</span><span class="sxs-lookup"><span data-stu-id="0a720-169">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="0a720-170">型情報がない場合、型は `any` になります。</span><span class="sxs-lookup"><span data-stu-id="0a720-170">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

<span data-ttu-id="0a720-171">カスタム関数がストリーミング関数であることを示すのに使用されます。</span><span class="sxs-lookup"><span data-stu-id="0a720-171">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="0a720-172">最後のパラメーターは、`CustomFunctions.StreamingInvocation<ResultType>` 型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0a720-172">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="0a720-173">関数は `void` を返します。</span><span class="sxs-lookup"><span data-stu-id="0a720-173">The function should return `void`.</span></span>

<span data-ttu-id="0a720-174">ストリーミング関数は値を直接返さず、代わりに、最後のパラメーターを使用して `setResult(result: ResultType)` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="0a720-174">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="0a720-175">ストリーム関数によってスローされる例外は無視されます。</span><span class="sxs-lookup"><span data-stu-id="0a720-175">Exceptions thrown by a streaming function are ignored.</span></span> `setResult()` <span data-ttu-id="0a720-176">が、エラー結果を示すために、Error により呼び出されることがあります。</span><span class="sxs-lookup"><span data-stu-id="0a720-176">may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="0a720-177">ストリーミング関数は、[@volatile](#volatile) としてマークできません。</span><span class="sxs-lookup"><span data-stu-id="0a720-177">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

<span data-ttu-id="0a720-178">揮発性関数とは、引数を取らない場合や引数が変更されていない場合でも、ある瞬間と次の瞬間では結果が異なる可能性があると見なされる関数です。</span><span class="sxs-lookup"><span data-stu-id="0a720-178">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="0a720-179">Excel では、再計算が実行される度に、揮発性関数を含むセルはすべての参照先と共に、再評価されます。</span><span class="sxs-lookup"><span data-stu-id="0a720-179">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="0a720-180">このため、揮発性関数を多用し過ぎると再計算にかかる時間が長くなる可能性があるため、多用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="0a720-180">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="0a720-181">ストリーミング関数に揮発性関数は使用できません。</span><span class="sxs-lookup"><span data-stu-id="0a720-181">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="0a720-182">型</span><span class="sxs-lookup"><span data-stu-id="0a720-182">Types</span></span>

<span data-ttu-id="0a720-183">パラメーターの型を指定すると、Excel は値を指定した型に変換してから関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="0a720-183">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="0a720-184">型が`any`の場合、変換は実行されません。</span><span class="sxs-lookup"><span data-stu-id="0a720-184">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="0a720-185">値の型</span><span class="sxs-lookup"><span data-stu-id="0a720-185">Value types</span></span>

<span data-ttu-id="0a720-186">1 つの値は、`boolean`、 `number`、`string`の型のいずれかを使用して表現できます。</span><span class="sxs-lookup"><span data-stu-id="0a720-186">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="0a720-187">マトリックス型</span><span class="sxs-lookup"><span data-stu-id="0a720-187">Matrix type</span></span>

<span data-ttu-id="0a720-188">2 次元配列型を使用して、パラメーターまたは戻り値を値のマトリックスにします。</span><span class="sxs-lookup"><span data-stu-id="0a720-188">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="0a720-189">たとえば、`number[][]`の型は数字のマトリックスを示します。</span><span class="sxs-lookup"><span data-stu-id="0a720-189">For example, the type `number[][]` indicates a matrix of numbers.</span></span> `string[][]` <span data-ttu-id="0a720-190">は、文字列のマトリックスを示します。</span><span class="sxs-lookup"><span data-stu-id="0a720-190">indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="0a720-191">エラーの種類</span><span class="sxs-lookup"><span data-stu-id="0a720-191">Error Type</span></span>

<span data-ttu-id="0a720-192">非ストリーミング関数は、エラーの種類を返すことによりエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="0a720-192">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="0a720-193">ストリーミング関数は、エラーの種類で setResult() を返してエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="0a720-193">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="0a720-194">Promise</span><span class="sxs-lookup"><span data-stu-id="0a720-194">Promise object.</span></span>

<span data-ttu-id="0a720-195">関数は Promise を返すことができ、Promise が解決されたときに値を提供します。</span><span class="sxs-lookup"><span data-stu-id="0a720-195">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="0a720-196">Promise が拒否された場合は、エラーになります。</span><span class="sxs-lookup"><span data-stu-id="0a720-196">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="0a720-197">その他の型</span><span class="sxs-lookup"><span data-stu-id="0a720-197">Other solution types</span></span>

<span data-ttu-id="0a720-198">その他の型は、エラーとして処理されます。</span><span class="sxs-lookup"><span data-stu-id="0a720-198">Any other type will be treated as an error.</span></span>

## <a name="see-also"></a><span data-ttu-id="0a720-199">関連項目</span><span class="sxs-lookup"><span data-stu-id="0a720-199">See also</span></span>

* [<span data-ttu-id="0a720-200">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="0a720-200">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="0a720-201">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="0a720-201">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="0a720-202">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="0a720-202">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="0a720-203">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="0a720-203">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="0a720-204">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="0a720-204">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="0a720-205">カスタム関数のデバッグ</span><span class="sxs-lookup"><span data-stu-id="0a720-205">Custom functions debugging</span></span>](custom-functions-debugging.md)
