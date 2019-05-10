---
ms.date: 05/03/2019
description: JSDOC タグを使用して、カスタム関数の JSON メタデータを動的に作成する。
title: カスタム関数用の JSON メタデータの自動生成
localization_priority: Priority
ms.openlocfilehash: df1c0114597e2aa98a15db48c515469fb9db6cd9
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628089"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="833fe-103">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="833fe-103">Create JSON metadata for custom functions</span></span>

<span data-ttu-id="833fe-104">Excel カスタム関数が JavaScript または TypeScript で記述されている場合、カスタム関数に関する追加の情報を提供するために、JSDoc タグが使用されます。</span><span class="sxs-lookup"><span data-stu-id="833fe-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="833fe-105">JSDoc タグはビルド時に使用して、[JSON メタデータ ファイル](custom-functions-json.md)を作成します。</span><span class="sxs-lookup"><span data-stu-id="833fe-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="833fe-106">JSDoc タグを使用すると、JSON メタデータ ファイルを手動で編集する手間が省けます。</span><span class="sxs-lookup"><span data-stu-id="833fe-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="833fe-107">JavaScript または TypeScript 関数のコード コメントに`@customfunction`タグを追加して、カスタム関数としてマークします。</span><span class="sxs-lookup"><span data-stu-id="833fe-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="833fe-108">関数パラメーターの型は、JavaScript の [@param](#param) タグを使用して指定するか、TypeScript の[関数の型](https://www.typescriptlang.org/docs/handbook/functions.html)から指定できます。</span><span class="sxs-lookup"><span data-stu-id="833fe-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="833fe-109">詳細については、「[@param](#param) タグ」セクションと「[型](#types)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="833fe-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="833fe-110">JSDoc タグ</span><span class="sxs-lookup"><span data-stu-id="833fe-110">JSDoc Tags</span></span>
<span data-ttu-id="833fe-111">Excel カスタム関数では、次の JSDoc タグを利用できます。</span><span class="sxs-lookup"><span data-stu-id="833fe-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="833fe-112">@cancelable</span><span class="sxs-lookup"><span data-stu-id="833fe-112">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="833fe-113">[@customfunction](#customfunction) id 名</span><span class="sxs-lookup"><span data-stu-id="833fe-113">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="833fe-114">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="833fe-114">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="833fe-115">[@param](#param) _{type}_ 名前の説明</span><span class="sxs-lookup"><span data-stu-id="833fe-115">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="833fe-116">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="833fe-116">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="833fe-117">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="833fe-117">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="833fe-118">@streaming</span><span class="sxs-lookup"><span data-stu-id="833fe-118">@streaming</span></span>](#streaming)
* [<span data-ttu-id="833fe-119">@volatile</span><span class="sxs-lookup"><span data-stu-id="833fe-119">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="833fe-120">@cancelable</span><span class="sxs-lookup"><span data-stu-id="833fe-120">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="833fe-121">関数がキャンセルされた場合にカスタム関数がアクションを実行することを示します。</span><span class="sxs-lookup"><span data-stu-id="833fe-121">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="833fe-122">最後の関数パラメーターは `CustomFunctions.CancelableInvocation` の型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="833fe-122">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="833fe-123">関数は `oncanceled` プロパティに関数を割り当て、関数がキャンセルされた場合に実行するアクションを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="833fe-123">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="833fe-124">最後の関数のパラメーターが `CustomFunctions.CancelableInvocation` 型の場合、タグは表示されませんが、`@cancelable` と見なされます。</span><span class="sxs-lookup"><span data-stu-id="833fe-124">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="833fe-125">関数には `@cancelable` と `@streaming` の両方のタグを含めることはできません。</span><span class="sxs-lookup"><span data-stu-id="833fe-125">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="833fe-126">@customfunction</span><span class="sxs-lookup"><span data-stu-id="833fe-126">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="833fe-127">構文: @customfunction _id_ _名_</span><span class="sxs-lookup"><span data-stu-id="833fe-127">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="833fe-128">このタグを指定すると、JavaScript または TypeScript の関数を、Excel のカスタム関数として処理できます。</span><span class="sxs-lookup"><span data-stu-id="833fe-128">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="833fe-129">このタグは、カスタム関数のメタデータを作成するために必要です。</span><span class="sxs-lookup"><span data-stu-id="833fe-129">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="833fe-130">次への呼び出しもあります: `CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="833fe-130">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

#### <a name="id"></a><span data-ttu-id="833fe-131">id</span><span class="sxs-lookup"><span data-stu-id="833fe-131">id</span></span>

<span data-ttu-id="833fe-132">id は、文書に格納されているカスタム関数の不変の識別子として使用されます。</span><span class="sxs-lookup"><span data-stu-id="833fe-132">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="833fe-133">変更する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="833fe-133">It should not change.</span></span>

* <span data-ttu-id="833fe-134">id が提供されていない場合、JavaScript または TypeScript の関数名は大文字に変換され、許可されない文字は削除されます。</span><span class="sxs-lookup"><span data-stu-id="833fe-134">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="833fe-135">id はすべてのカスタム関数で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="833fe-135">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="833fe-136">使用できる文字は、A～Z、a～z、0～9、ピリオド (.) に制限されています。</span><span class="sxs-lookup"><span data-stu-id="833fe-136">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="833fe-137">name</span><span class="sxs-lookup"><span data-stu-id="833fe-137">name</span></span>

<span data-ttu-id="833fe-138">カスタム関数の表示名を提供します。</span><span class="sxs-lookup"><span data-stu-id="833fe-138">Provides the display name for the custom function.</span></span>

* <span data-ttu-id="833fe-139">name が指定されていない場合、id が名前としても使用されます。</span><span class="sxs-lookup"><span data-stu-id="833fe-139">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="833fe-140">使用できる文字は、文字 [Unicode アルファベット](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、ピリオド (.)、およびアンダースコア (\_)です。</span><span class="sxs-lookup"><span data-stu-id="833fe-140">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="833fe-141">最初の文字は、アルファベット文字にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="833fe-141">Must start with a letter.</span></span>
* <span data-ttu-id="833fe-142">最大文字数は 128 文字です。</span><span class="sxs-lookup"><span data-stu-id="833fe-142">Maximum length is 128 characters.</span></span>

---
### <a name="helpurl"></a><span data-ttu-id="833fe-143">@helpurl</span><span class="sxs-lookup"><span data-stu-id="833fe-143">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="833fe-144">構文: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="833fe-144">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="833fe-145">指定された _url_ が Excel で表示されます。</span><span class="sxs-lookup"><span data-stu-id="833fe-145">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a><span data-ttu-id="833fe-146">@param</span><span class="sxs-lookup"><span data-stu-id="833fe-146">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="833fe-147">JavaScript</span><span class="sxs-lookup"><span data-stu-id="833fe-147">JavaScript</span></span>

<span data-ttu-id="833fe-148">JavaScript 構文: @param {type} 名_の説明_</span><span class="sxs-lookup"><span data-stu-id="833fe-148">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="833fe-149">`{type}` は、中かっこ内の型の情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="833fe-149">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="833fe-150">使用できる型に関する詳細については、「[型](##types)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="833fe-150">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="833fe-151">省略可能: 指定しない場合、`any` 型が使用されます。</span><span class="sxs-lookup"><span data-stu-id="833fe-151">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="833fe-152">`name` は、@param タグを適用するパラメーターを指定します。</span><span class="sxs-lookup"><span data-stu-id="833fe-152">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="833fe-153">必須です。</span><span class="sxs-lookup"><span data-stu-id="833fe-153">Required.</span></span>
* <span data-ttu-id="833fe-154">`description` は、Excel で表示される関数のパラメーターの説明を示します。</span><span class="sxs-lookup"><span data-stu-id="833fe-154">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="833fe-155">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="833fe-155">Optional.</span></span>

<span data-ttu-id="833fe-156">カスタム関数内のパラメーターを省略可能と指定する方法:</span><span class="sxs-lookup"><span data-stu-id="833fe-156">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="833fe-157">パラメーター名を角かっこで囲みます。</span><span class="sxs-lookup"><span data-stu-id="833fe-157">Put square brackets around the parameter name.</span></span> <span data-ttu-id="833fe-158">例: `@param {string} [text] Optional text`。</span><span class="sxs-lookup"><span data-stu-id="833fe-158">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="833fe-159">省略可能なパラメーターの既定値は `null` です。</span><span class="sxs-lookup"><span data-stu-id="833fe-159">The default value for optional parameters is `null`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="833fe-160">TypeScript</span><span class="sxs-lookup"><span data-stu-id="833fe-160">TypeScript</span></span>

<span data-ttu-id="833fe-161">TypeScript 構文: @param 名 _の説明_</span><span class="sxs-lookup"><span data-stu-id="833fe-161">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="833fe-162">`name` は、@param タグを適用するパラメーターを指定します。</span><span class="sxs-lookup"><span data-stu-id="833fe-162">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="833fe-163">必須です。</span><span class="sxs-lookup"><span data-stu-id="833fe-163">Required.</span></span>
* <span data-ttu-id="833fe-164">`description` は、Excel で表示される関数のパラメーターの説明を示します。</span><span class="sxs-lookup"><span data-stu-id="833fe-164">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="833fe-165">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="833fe-165">Optional.</span></span>

<span data-ttu-id="833fe-166">使用できる関数のパラメーターの型に関する詳細については、「[型](##types)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="833fe-166">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="833fe-167">カスタム関数のパラメーターを省略可能として示すには、以下のいずれかを実行します。</span><span class="sxs-lookup"><span data-stu-id="833fe-167">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="833fe-168">省略可能なパラメーターを使用する。</span><span class="sxs-lookup"><span data-stu-id="833fe-168">Use an optional parameter.</span></span> <span data-ttu-id="833fe-169">例: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="833fe-169">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="833fe-170">パラメーターに既定値を指定する。</span><span class="sxs-lookup"><span data-stu-id="833fe-170">Give the parameter a default value.</span></span> <span data-ttu-id="833fe-171">例: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="833fe-171">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="833fe-172">@param の詳しい説明については、「[JSDoc](https://usejsdoc.org/tags-param.html)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="833fe-172">For detailed description of the @param see: [JSDoc](https://usejsdoc.org/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="833fe-173">省略可能なパラメーターの既定値は `null` です。</span><span class="sxs-lookup"><span data-stu-id="833fe-173">The default value for optional parameters is `null`.</span></span>

---
### <a name="requiresaddress"></a><span data-ttu-id="833fe-174">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="833fe-174">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="833fe-175">関数が評価されているセルのアドレスを指定する必要があることを示します。</span><span class="sxs-lookup"><span data-stu-id="833fe-175">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="833fe-176">最後の関数のパラメーターは、`CustomFunctions.Invocation` 型または派生型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="833fe-176">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="833fe-177">関数が呼び出されると、`address` プロパティにアドレスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="833fe-177">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="833fe-178">@returns</span><span class="sxs-lookup"><span data-stu-id="833fe-178">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="833fe-179">構文: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="833fe-179">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="833fe-180">戻り値の型を指定します。</span><span class="sxs-lookup"><span data-stu-id="833fe-180">Provides the type for the return value.</span></span>

<span data-ttu-id="833fe-181">`{type}` を省略すると、TypeScript の型情報が使用されます。</span><span class="sxs-lookup"><span data-stu-id="833fe-181">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="833fe-182">型情報がない場合、型は `any` になります。</span><span class="sxs-lookup"><span data-stu-id="833fe-182">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a><span data-ttu-id="833fe-183">@streaming</span><span class="sxs-lookup"><span data-stu-id="833fe-183">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="833fe-184">カスタム関数がストリーミング関数であることを示すのに使用されます。</span><span class="sxs-lookup"><span data-stu-id="833fe-184">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="833fe-185">最後のパラメーターは、`CustomFunctions.StreamingInvocation<ResultType>` 型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="833fe-185">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="833fe-186">関数は `void` を返します。</span><span class="sxs-lookup"><span data-stu-id="833fe-186">The function should return `void`.</span></span>

<span data-ttu-id="833fe-187">ストリーミング関数は値を直接返さず、代わりに、最後のパラメーターを使用して `setResult(result: ResultType)` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="833fe-187">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="833fe-188">ストリーム関数によってスローされる例外は無視されます。</span><span class="sxs-lookup"><span data-stu-id="833fe-188">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="833fe-189">`setResult()` が、エラー結果を示すために、Error により呼び出されることがあります。</span><span class="sxs-lookup"><span data-stu-id="833fe-189">`setResult()` may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="833fe-190">ストリーミング関数は、[@volatile](#volatile) としてマークできません。</span><span class="sxs-lookup"><span data-stu-id="833fe-190">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="833fe-191">@volatile</span><span class="sxs-lookup"><span data-stu-id="833fe-191">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="833fe-192">揮発性関数とは、引数を取らない場合や引数が変更されていない場合でも、ある瞬間と次の瞬間では結果が異なる可能性があると見なされる関数です。</span><span class="sxs-lookup"><span data-stu-id="833fe-192">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="833fe-193">Excel では、再計算が実行される度に、揮発性関数を含むセルはすべての参照先と共に、再評価されます。</span><span class="sxs-lookup"><span data-stu-id="833fe-193">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="833fe-194">このため、揮発性関数を多用し過ぎると再計算にかかる時間が長くなる可能性があるため、多用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="833fe-194">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="833fe-195">ストリーミング関数に揮発性関数は使用できません。</span><span class="sxs-lookup"><span data-stu-id="833fe-195">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="833fe-196">型</span><span class="sxs-lookup"><span data-stu-id="833fe-196">Types</span></span>

<span data-ttu-id="833fe-197">パラメーターの型を指定すると、Excel は値を指定した型に変換してから関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="833fe-197">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="833fe-198">型が`any`の場合、変換は実行されません。</span><span class="sxs-lookup"><span data-stu-id="833fe-198">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="833fe-199">値の型</span><span class="sxs-lookup"><span data-stu-id="833fe-199">Value types</span></span>

<span data-ttu-id="833fe-200">1 つの値は、`boolean`、 `number`、`string`の型のいずれかを使用して表現できます。</span><span class="sxs-lookup"><span data-stu-id="833fe-200">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="833fe-201">マトリックス型</span><span class="sxs-lookup"><span data-stu-id="833fe-201">Matrix type</span></span>

<span data-ttu-id="833fe-202">2 次元配列型を使用して、パラメーターまたは戻り値を値のマトリックスにします。</span><span class="sxs-lookup"><span data-stu-id="833fe-202">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="833fe-203">たとえば、`number[][]`の型は数字のマトリックスを示します。</span><span class="sxs-lookup"><span data-stu-id="833fe-203">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="833fe-204">`string[][]` は、文字列のマトリックスを示します。</span><span class="sxs-lookup"><span data-stu-id="833fe-204">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="833fe-205">エラーの種類</span><span class="sxs-lookup"><span data-stu-id="833fe-205">Error type</span></span>

<span data-ttu-id="833fe-206">非ストリーミング関数は、エラーの種類を返すことによりエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="833fe-206">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="833fe-207">ストリーミング関数は、エラーの種類で setResult() を返してエラーを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="833fe-207">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="833fe-208">Promise</span><span class="sxs-lookup"><span data-stu-id="833fe-208">Promise</span></span>

<span data-ttu-id="833fe-209">関数は Promise を返すことができ、Promise が解決されたときに値を提供します。</span><span class="sxs-lookup"><span data-stu-id="833fe-209">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="833fe-210">Promise が拒否された場合は、エラーになります。</span><span class="sxs-lookup"><span data-stu-id="833fe-210">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="833fe-211">その他の型</span><span class="sxs-lookup"><span data-stu-id="833fe-211">Other types</span></span>

<span data-ttu-id="833fe-212">その他の型は、エラーとして処理されます。</span><span class="sxs-lookup"><span data-stu-id="833fe-212">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="833fe-213">次の手順</span><span class="sxs-lookup"><span data-stu-id="833fe-213">Next steps</span></span>
<span data-ttu-id="833fe-214">[カスタム関数用の命名規則](custom-functions-naming.md)について説明します。</span><span class="sxs-lookup"><span data-stu-id="833fe-214">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="833fe-215">または、[JSON ファイルを手で書く](custom-functions-json.md)必要のある[機能をローカライズする](custom-functions-localize.md)方法を確認してください。</span><span class="sxs-lookup"><span data-stu-id="833fe-215">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="833fe-216">関連項目</span><span class="sxs-lookup"><span data-stu-id="833fe-216">See also</span></span>

* [<span data-ttu-id="833fe-217">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="833fe-217">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="833fe-218">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="833fe-218">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="833fe-219">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="833fe-219">Create custom functions in Excel</span></span>](custom-functions-overview.md)
