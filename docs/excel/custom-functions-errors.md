---
ms.date: 03/11/2020
description: '#NULL! のようなエラーを処理して返す カスタム関数で'
title: カスタム関数でエラーを処理して返す (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 10bb7ca6ff612ef38b26b88fed5ce9ce81ed7edb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717048"
---
# <a name="handle-and-return-errors-from-your-custom-function-preview"></a><span data-ttu-id="95c17-104">カスタム関数でエラーを処理して返す (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="95c17-104">Handle and return errors from your custom function (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="95c17-105">この記事で説明する機能は現在プレビュー中であり、変更される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="95c17-105">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="95c17-106">これらを運用環境で使用することは現在サポートされていません。</span><span class="sxs-lookup"><span data-stu-id="95c17-106">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="95c17-107">プレビュー機能を試すには、 [Office Insider](https://insider.office.com/join)プログラムに参加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="95c17-107">You will need to join the [Office Insider](https://insider.office.com/join) program to try the preview features.</span></span>  <span data-ttu-id="95c17-108">プレビュー機能を試す良い方法は、Office 365 サブスクリプションを使用することです。</span><span class="sxs-lookup"><span data-stu-id="95c17-108">A good way to try out preview features is by using an Office 365 subscription.</span></span> <span data-ttu-id="95c17-109">Office 365 サブスクリプションをまだお持ちでない場合は、[Office 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加することで 90 日間の更新可能な無料の Office 365 サブスクリプションを入手できます。</span><span class="sxs-lookup"><span data-stu-id="95c17-109">If you don't already have an Office 365 subscription, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="95c17-110">カスタム関数の実行中に問題が発生した場合、エラーを返してユーザーに通知する必要があります。</span><span class="sxs-lookup"><span data-stu-id="95c17-110">If something goes wrong while your custom function runs, you will need to return an error to inform the user.</span></span> <span data-ttu-id="95c17-111">正数のみなど、特定のパラメーター要件がある場合は、パラメーターをテストし、正しくない場合はエラーをスローする必要があります。</span><span class="sxs-lookup"><span data-stu-id="95c17-111">If you have specific parameter requirements, such as only positive numbers, you will need to test the parameters and throw an error if they are not correct.</span></span> <span data-ttu-id="95c17-112">`try` - `catch` ブロックを使用して、カスタム関数の実行中に発生したエラーを検出することもできます。</span><span class="sxs-lookup"><span data-stu-id="95c17-112">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="95c17-113">エラーを検出してスローする</span><span class="sxs-lookup"><span data-stu-id="95c17-113">Detect and throw an error</span></span>

<span data-ttu-id="95c17-114">カスタム関数が動作するために zip コードパラメーターが正しい形式であることを確認する必要があるケースを見てみましょう。</span><span class="sxs-lookup"><span data-stu-id="95c17-114">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="95c17-115">次のカスタム関数は、正規表現を使用して郵便番号を確認します。</span><span class="sxs-lookup"><span data-stu-id="95c17-115">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="95c17-116">正しい場合は、(別の関数で) 都市を検索し、その値を返します。</span><span class="sxs-lookup"><span data-stu-id="95c17-116">If it is correct, then it will look up the city (in another function) and return the value.</span></span> <span data-ttu-id="95c17-117">正しくない場合は、セルに `#VALUE!` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="95c17-117">If it is not correct, it returns a `#VALUE!` error to the cell.</span></span>

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

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="95c17-118">The CustomFunctions.Error object</span><span class="sxs-lookup"><span data-stu-id="95c17-118">The CustomFunctions.Error object</span></span>

<span data-ttu-id="95c17-119">`CustomFunctions.Error` オブジェクトは、セルにエラーを返すために使用されます。</span><span class="sxs-lookup"><span data-stu-id="95c17-119">The `CustomFunctions.Error` object is used to return an error back to the cell.</span></span> <span data-ttu-id="95c17-120">オブジェクトを作成するときに、次の `ErrorCode` 列挙値のいずれかを使用して、使用するエラーを指定します。</span><span class="sxs-lookup"><span data-stu-id="95c17-120">When you create the object, specify which error you want to use by using one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="95c17-121">ErrorCode enum value</span><span class="sxs-lookup"><span data-stu-id="95c17-121">ErrorCode enum value</span></span>  |<span data-ttu-id="95c17-122">Excel のセル値</span><span class="sxs-lookup"><span data-stu-id="95c17-122">Excel cell value</span></span>  |<span data-ttu-id="95c17-123">意味</span><span class="sxs-lookup"><span data-stu-id="95c17-123">Meaning</span></span>  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="95c17-124">数式で使用されている値の型が間違っている。</span><span class="sxs-lookup"><span data-stu-id="95c17-124">A value used in the formula is the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="95c17-125">機能またはサービスが利用できない。</span><span class="sxs-lookup"><span data-stu-id="95c17-125">The function or service is not available.</span></span> |
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="95c17-126">JavaScript ではゼロ除算が許可されるため、この状態を検出するには、慎重にエラーハンドラをに記述する必要があります。</span><span class="sxs-lookup"><span data-stu-id="95c17-126">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="95c17-127">数式で使用されている番号に問題がある。</span><span class="sxs-lookup"><span data-stu-id="95c17-127">There is a problem with the number used in the formula</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="95c17-128">数式の範囲が交わることはありません。</span><span class="sxs-lookup"><span data-stu-id="95c17-128">The ranges in the formula do not intersect.</span></span> |

<span data-ttu-id="95c17-129">次のコードサンプルは、無効な番号 (`#NUM!`) に対してエラーを作成して返す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="95c17-129">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="95c17-130">`#VALUE!` エラーを返す場合、ユーザーがセルにカーソルを合わせたときにポップアップに表示されるカスタムメッセージを含めることもできます。</span><span class="sxs-lookup"><span data-stu-id="95c17-130">When you return a `#VALUE!` error you can also include a custom message that will be shown in a popup when the user hovers over the cell.</span></span> <span data-ttu-id="95c17-131">次の例は、カスタムのエラーメッセージを返す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="95c17-131">The following example shows how to return a custom error message.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="95c17-132">Use try-catch blocks</span><span class="sxs-lookup"><span data-stu-id="95c17-132">Use try-catch blocks</span></span>

<span data-ttu-id="95c17-133">通常、発生する可能性があるエラーをキャッチするには、カスタム関数で `try` - `catch` ブロックを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="95c17-133">In general, you should use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="95c17-134">コードで例外を処理しない場合は、Excel に返されます。</span><span class="sxs-lookup"><span data-stu-id="95c17-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="95c17-135">既定では、Excel は未処理の例外に対して `#VALUE!` を返します。</span><span class="sxs-lookup"><span data-stu-id="95c17-135">By default, Excel returns `#VALUE!` for an unhandled exception.</span></span>

<span data-ttu-id="95c17-136">次のコードサンプルでは、カスタム関数を使用して REST サービスの呼び出しを行ないます。</span><span class="sxs-lookup"><span data-stu-id="95c17-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="95c17-137">たとえば REST サービスがエラーを返したり、ネットワークがダウンした場合には、呼び出しが失敗することもあります。</span><span class="sxs-lookup"><span data-stu-id="95c17-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="95c17-138">この場合、カスタム関数は Web 呼び出しが失敗したことを示す `#N/A` を返します。</span><span class="sxs-lookup"><span data-stu-id="95c17-138">If this happens, the custom function will return `#N/A` to indicate the web call failed.</span></span>


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

## <a name="next-steps"></a><span data-ttu-id="95c17-139">次の手順</span><span class="sxs-lookup"><span data-stu-id="95c17-139">Next steps</span></span>

<span data-ttu-id="95c17-140">[自分のカスタム関数で問題をトラブルシューティングを行う](custom-functions-troubleshooting.md)方法についての詳細を確認する。</span><span class="sxs-lookup"><span data-stu-id="95c17-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="95c17-141">関連項目</span><span class="sxs-lookup"><span data-stu-id="95c17-141">See also</span></span>

* [<span data-ttu-id="95c17-142">カスタム関数のデバッグ</span><span class="sxs-lookup"><span data-stu-id="95c17-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="95c17-143">カスタム関数の要件</span><span class="sxs-lookup"><span data-stu-id="95c17-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="95c17-144">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="95c17-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
