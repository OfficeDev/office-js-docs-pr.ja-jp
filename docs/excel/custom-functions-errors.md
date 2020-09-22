---
ms.date: 09/21/2020
description: '#NULL! のようなエラーを処理して返す カスタム関数から。'
title: カスタム関数を処理し、エラーを返します。
localization_priority: Normal
ms.openlocfilehash: 58c2ab432a4525f660e2d89735fd3add6e76fa7f
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175529"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a><span data-ttu-id="33c90-104">カスタム関数を処理し、エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="33c90-104">Handle and return errors from your custom function</span></span>

<span data-ttu-id="33c90-105">カスタム関数の実行中に何らかの問題が発生した場合は、ユーザーに通知するエラーを返します。</span><span class="sxs-lookup"><span data-stu-id="33c90-105">If something goes wrong while your custom function runs, return an error to inform the user.</span></span> <span data-ttu-id="33c90-106">正の数だけなど、特定のパラメーター要件がある場合は、パラメーターをテストし、正しくない場合はエラーをスローします。</span><span class="sxs-lookup"><span data-stu-id="33c90-106">If you have specific parameter requirements, such as only positive numbers, test the parameters and throw an error if they aren't correct.</span></span> <span data-ttu-id="33c90-107">`try` - `catch` ブロックを使用して、カスタム関数の実行中に発生したエラーを検出することもできます。</span><span class="sxs-lookup"><span data-stu-id="33c90-107">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="33c90-108">エラーを検出してスローする</span><span class="sxs-lookup"><span data-stu-id="33c90-108">Detect and throw an error</span></span>

<span data-ttu-id="33c90-109">カスタム関数が動作するために zip コードパラメーターが正しい形式であることを確認する必要があるケースを見てみましょう。</span><span class="sxs-lookup"><span data-stu-id="33c90-109">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="33c90-110">次のカスタム関数は、正規表現を使用して郵便番号を確認します。</span><span class="sxs-lookup"><span data-stu-id="33c90-110">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="33c90-111">郵便番号の形式が正しい場合は、別の関数を使用して都市を検索し、その値を返します。</span><span class="sxs-lookup"><span data-stu-id="33c90-111">If the zip code format is correct, then it will look up the city using another function and return the value.</span></span> <span data-ttu-id="33c90-112">形式が有効でない場合、この関数は `#VALUE!` セルにエラーを返します。</span><span class="sxs-lookup"><span data-stu-id="33c90-112">If the format isn't valid, the function returns a `#VALUE!` error to the cell.</span></span>

```typescript
/**
* Gets a city name for the given U.S. zip code.
* @customfunction
* @param {string} zipCode
* @returns The city of the zip code.
*/
function getCity(zipCode: string): string {
  let isValidZip = /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(zipCode);
  if (isValidZip) return cityLookup(zipCode);
  let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Please provide a valid U.S. zip code.");
  throw error;
}
```

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="33c90-113">The CustomFunctions.Error object</span><span class="sxs-lookup"><span data-stu-id="33c90-113">The CustomFunctions.Error object</span></span>

<span data-ttu-id="33c90-114">Error オブジェクトは、セルにエラーを返すために使用されます[。](/javascript/api/custom-functions-runtime/customfunctions.error)</span><span class="sxs-lookup"><span data-stu-id="33c90-114">The [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) object is used to return an error back to the cell.</span></span> <span data-ttu-id="33c90-115">オブジェクトを作成するときに、次の列挙値のいずれかを選択して、使用するエラーを指定し `ErrorCode` ます。</span><span class="sxs-lookup"><span data-stu-id="33c90-115">When you create the object, specify which error you want to use by choosing one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="33c90-116">ErrorCode enum value</span><span class="sxs-lookup"><span data-stu-id="33c90-116">ErrorCode enum value</span></span>  |<span data-ttu-id="33c90-117">Excel のセル値</span><span class="sxs-lookup"><span data-stu-id="33c90-117">Excel cell value</span></span>  |<span data-ttu-id="33c90-118">意味</span><span class="sxs-lookup"><span data-stu-id="33c90-118">Meaning</span></span>  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="33c90-119">JavaScript ではゼロ除算が許可されるため、この状態を検出するには、慎重にエラーハンドラをに記述する必要があります。</span><span class="sxs-lookup"><span data-stu-id="33c90-119">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidName`    | `#NAME?`  | <span data-ttu-id="33c90-120">関数名に入力ミスがあります。</span><span class="sxs-lookup"><span data-stu-id="33c90-120">There is a typo in the function name.</span></span> <span data-ttu-id="33c90-121">このエラーは、カスタム関数の入力エラーとしてサポートされますが、カスタム関数の出力エラーとしてはサポートされていないことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="33c90-121">Note that this error is supported as a custom function input error, but not as a custom function output error.</span></span> | 
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="33c90-122">数式の数値に問題があります。</span><span class="sxs-lookup"><span data-stu-id="33c90-122">There is a problem with a number in the formula.</span></span> |
|`invalidReference` | `#REF!` | <span data-ttu-id="33c90-123">関数が無効なセルを参照しています。</span><span class="sxs-lookup"><span data-stu-id="33c90-123">The function refers to an invalid cell.</span></span> <span data-ttu-id="33c90-124">このエラーは、カスタム関数の入力エラーとしてサポートされますが、カスタム関数の出力エラーとしてはサポートされていないことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="33c90-124">Note that this error is supported as a custom function input error, but not as a custom function output error.</span></span>|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="33c90-125">数式の値の種類が正しくありません。</span><span class="sxs-lookup"><span data-stu-id="33c90-125">A value in the formula is of the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="33c90-126">関数またはサービスを使用できません。</span><span class="sxs-lookup"><span data-stu-id="33c90-126">The function or service isn't available.</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="33c90-127">数式の範囲は交差しません。</span><span class="sxs-lookup"><span data-stu-id="33c90-127">The ranges in the formula don't intersect.</span></span> |

<span data-ttu-id="33c90-128">次のコードサンプルは、無効な番号 (`#NUM!`) に対してエラーを作成して返す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="33c90-128">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="33c90-129">およびエラーでは、 `#VALUE!` `#N/A` カスタムエラーメッセージもサポートされます。</span><span class="sxs-lookup"><span data-stu-id="33c90-129">The `#VALUE!` and `#N/A` errors also support custom error messages.</span></span> <span data-ttu-id="33c90-130">エラーインジケーターメニューにカスタムエラーメッセージが表示されます。このメニューでは、エラーが発生した各セルのエラーフラグの上にカーソルがアクセスします。</span><span class="sxs-lookup"><span data-stu-id="33c90-130">Custom error messages are displayed in the error indicator menu, which is accessed by hovering over the error flag on each cell with an error.</span></span> <span data-ttu-id="33c90-131">次の例は、エラーが発生したカスタムエラーメッセージを返す方法を示して `#VALUE!` います。</span><span class="sxs-lookup"><span data-stu-id="33c90-131">The following example shows how to return a custom error message with the `#VALUE!` error.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="33c90-132">Use try-catch blocks</span><span class="sxs-lookup"><span data-stu-id="33c90-132">Use try-catch blocks</span></span>

<span data-ttu-id="33c90-133">一般的には、 `try` - `catch` カスタム関数のブロックを使用して発生する可能性のあるエラーを検出します。</span><span class="sxs-lookup"><span data-stu-id="33c90-133">In general, use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="33c90-134">コードで例外を処理しない場合は、Excel に返されます。</span><span class="sxs-lookup"><span data-stu-id="33c90-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="33c90-135">既定で `#VALUE!` は、処理されないエラーまたは例外に対して Excel が返します。</span><span class="sxs-lookup"><span data-stu-id="33c90-135">By default, Excel returns `#VALUE!` for unhandled errors or exceptions.</span></span>

<span data-ttu-id="33c90-136">次のコードサンプルでは、カスタム関数を使用して REST サービスの呼び出しを行ないます。</span><span class="sxs-lookup"><span data-stu-id="33c90-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="33c90-137">たとえば REST サービスがエラーを返したり、ネットワークがダウンした場合には、呼び出しが失敗することもあります。</span><span class="sxs-lookup"><span data-stu-id="33c90-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="33c90-138">このような場合、カスタム関数は、 `#N/A` web 呼び出しが失敗したことを示すためにを返します。</span><span class="sxs-lookup"><span data-stu-id="33c90-138">If this happens, the custom function will return `#N/A` to indicate that the web call failed.</span></span>


```typescript
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + commentID;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    })
}
```

## <a name="next-steps"></a><span data-ttu-id="33c90-139">次の手順</span><span class="sxs-lookup"><span data-stu-id="33c90-139">Next steps</span></span>

<span data-ttu-id="33c90-140">[自分のカスタム関数で問題をトラブルシューティングを行う](custom-functions-troubleshooting.md)方法についての詳細を確認する。</span><span class="sxs-lookup"><span data-stu-id="33c90-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="33c90-141">関連項目</span><span class="sxs-lookup"><span data-stu-id="33c90-141">See also</span></span>

* [<span data-ttu-id="33c90-142">カスタム関数のデバッグ</span><span class="sxs-lookup"><span data-stu-id="33c90-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="33c90-143">カスタム関数の要件</span><span class="sxs-lookup"><span data-stu-id="33c90-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="33c90-144">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="33c90-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
