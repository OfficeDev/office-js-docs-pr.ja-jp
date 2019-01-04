---
title: エラー処理
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: d9545831c27bdf11d1567f4c012ff678b1edb026
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457643"
---
# <a name="error-handling"></a><span data-ttu-id="68087-102">エラー処理</span><span class="sxs-lookup"><span data-stu-id="68087-102">Error handling</span></span>

<span data-ttu-id="68087-103">Excel JavaScript API を使用してアドインを作成する場合は、実行時エラーを考慮するために、エラー処理ロジックを含めます。</span><span class="sxs-lookup"><span data-stu-id="68087-103">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="68087-104">これは、API の非同期の性質のために重要です。</span><span class="sxs-lookup"><span data-stu-id="68087-104">Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="68087-105">**sync()** メソッドと Excel JavaScript API の非同期性の詳細については、「[Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68087-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="68087-106">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="68087-106">Best practices</span></span>

<span data-ttu-id="68087-107">このドキュメントのコード サンプルでは、`Excel.run` へのすべての呼び出しに、`Excel.run` 内で発生したエラーを検出するための `catch` ステートメントが付いていることがわかります。</span><span class="sxs-lookup"><span data-stu-id="68087-107">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`.</span></span> <span data-ttu-id="68087-108">Excel JavaScript Api を使用してアドインを構築するときには、同じパターンを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="68087-108">We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

```js
Excel.run(function (context) {
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);
```

## <a name="api-errors"></a><span data-ttu-id="68087-109">API エラー</span><span class="sxs-lookup"><span data-stu-id="68087-109">API errors</span></span>

<span data-ttu-id="68087-110">Excel JavaScript API 要求が正常に実行されない場合、API は次のプロパティを含むエラー オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="68087-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="68087-111">**code**:エラー メッセージの `code` プロパティには、`OfficeExtension.ErrorCodes` または `Excel.ErrorCodes` リストの一部である文字列が含まれます。</span><span class="sxs-lookup"><span data-stu-id="68087-111">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span></span> <span data-ttu-id="68087-112">たとえば、エラー コード "InvalidReference" は、参照が指定された操作に対して有効でないことを示します。</span><span class="sxs-lookup"><span data-stu-id="68087-112">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span></span> <span data-ttu-id="68087-113">エラー コードはローカライズされません。</span><span class="sxs-lookup"><span data-stu-id="68087-113">Error codes are not localized.</span></span>

- <span data-ttu-id="68087-114">**message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれています。</span><span class="sxs-lookup"><span data-stu-id="68087-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="68087-115">このエラー メッセージは、エンド ユーザーが使用するためのものではありません。アドインによってエンド ユーザーに表示されるエラー メッセージは、エラー コードと適切なビジネス ロジックを使用して、判断する必要があります。</span><span class="sxs-lookup"><span data-stu-id="68087-115">The error message is not intended for end-user consumption; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end-users.</span></span>

- <span data-ttu-id="68087-116">**debugInfo**:存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="68087-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="68087-117">`console.log()` を使用してエラー メッセージをコンソールに出力すると、それらのメッセージはサーバー上でのみ表示されます。</span><span class="sxs-lookup"><span data-stu-id="68087-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="68087-118">これらのエラー メッセージが、アドインの作業ウィンドウやホスト アプリケーション内のいずれかの場所で、エンド ユーザーに対して表示されることはありません。</span><span class="sxs-lookup"><span data-stu-id="68087-118">End-users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="68087-119">エラー メッセージ</span><span class="sxs-lookup"><span data-stu-id="68087-119">Error Messages</span></span>

<span data-ttu-id="68087-120">次の表は、API から返される可能性のあるエラー一覧です。</span><span class="sxs-lookup"><span data-stu-id="68087-120">The following table defines a list of errors that the API may return.</span></span>

|<span data-ttu-id="68087-121">error.code</span><span class="sxs-lookup"><span data-stu-id="68087-121">error.code</span></span> | <span data-ttu-id="68087-122">error.message</span><span class="sxs-lookup"><span data-stu-id="68087-122">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="68087-123">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="68087-123">InvalidArgument</span></span> |<span data-ttu-id="68087-124">引数が無効であるか、存在しません。または形式が正しくありません。</span><span class="sxs-lookup"><span data-stu-id="68087-124">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="68087-125">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="68087-125">InvalidRequest</span></span>  |<span data-ttu-id="68087-126">要求を処理できません。</span><span class="sxs-lookup"><span data-stu-id="68087-126">Cannot process the request.</span></span>|
|<span data-ttu-id="68087-127">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="68087-127">InvalidReference</span></span>|<span data-ttu-id="68087-128">この参照は、現在の操作に対して無効です。</span><span class="sxs-lookup"><span data-stu-id="68087-128">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="68087-129">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="68087-129">InvalidBinding</span></span>  |<span data-ttu-id="68087-130">このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。</span><span class="sxs-lookup"><span data-stu-id="68087-130">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="68087-131">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="68087-131">InvalidSelection</span></span>|<span data-ttu-id="68087-132">現在の選択内容は、この操作では無効です。</span><span class="sxs-lookup"><span data-stu-id="68087-132">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="68087-133">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="68087-133">Unauthenticated</span></span> |<span data-ttu-id="68087-134">必要な認証情報が見つからないか、無効です。</span><span class="sxs-lookup"><span data-stu-id="68087-134">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="68087-135">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="68087-135">AccessDenied</span></span> |<span data-ttu-id="68087-136">要求された操作を実行できません。</span><span class="sxs-lookup"><span data-stu-id="68087-136">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="68087-137">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="68087-137">ItemNotFound</span></span> |<span data-ttu-id="68087-138">要求されたリソースは存在しません。</span><span class="sxs-lookup"><span data-stu-id="68087-138">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="68087-139">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="68087-139">ActivityLimitReached</span></span>|<span data-ttu-id="68087-140">アクティビティの制限に達しました。</span><span class="sxs-lookup"><span data-stu-id="68087-140">Activity limit has been reached.</span></span>|
|<span data-ttu-id="68087-141">GeneralException</span><span class="sxs-lookup"><span data-stu-id="68087-141">GeneralException</span></span>|<span data-ttu-id="68087-142">要求の処理中に内部エラーが発生しました。</span><span class="sxs-lookup"><span data-stu-id="68087-142">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="68087-143">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="68087-143">NotImplemented</span></span>  |<span data-ttu-id="68087-144">要求された機能は実装されていません。</span><span class="sxs-lookup"><span data-stu-id="68087-144">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="68087-145">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="68087-145">ServiceNotAvailable</span></span>|<span data-ttu-id="68087-146">サービスを利用できません。</span><span class="sxs-lookup"><span data-stu-id="68087-146">The service is unavailable.</span></span>|
|<span data-ttu-id="68087-147">Conflict</span><span class="sxs-lookup"><span data-stu-id="68087-147">Conflict</span></span>|<span data-ttu-id="68087-148">競合のため、要求を処理できませんでした。</span><span class="sxs-lookup"><span data-stu-id="68087-148">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="68087-149">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="68087-149">ItemAlreadyExists</span></span>|<span data-ttu-id="68087-150">作成中のリソースはすでに存在しています。</span><span class="sxs-lookup"><span data-stu-id="68087-150">The resource being created already exists.</span></span>|
|<span data-ttu-id="68087-151">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="68087-151">UnsupportedOperation</span></span>|<span data-ttu-id="68087-152">試行中の操作はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="68087-152">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="68087-153">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="68087-153">RequestAborted</span></span>|<span data-ttu-id="68087-154">実行時に要求が中止されました。</span><span class="sxs-lookup"><span data-stu-id="68087-154">The request was aborted during run time.</span></span>|
|<span data-ttu-id="68087-155">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="68087-155">ApiNotAvailable</span></span>|<span data-ttu-id="68087-156">要求された API は使用できません。</span><span class="sxs-lookup"><span data-stu-id="68087-156">The requested API is not available.</span></span>|
|<span data-ttu-id="68087-157">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="68087-157">InsertDeleteConflict</span></span>|<span data-ttu-id="68087-158">試行された挿入操作または削除操作で競合が発生しました。</span><span class="sxs-lookup"><span data-stu-id="68087-158">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="68087-159">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="68087-159">InvalidOperation</span></span>|<span data-ttu-id="68087-160">試行された操作は、このオブジェクトでは無効です。</span><span class="sxs-lookup"><span data-stu-id="68087-160">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="68087-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="68087-161">See also</span></span>

- [<span data-ttu-id="68087-162">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="68087-162">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="68087-163">OfficeExtension.Error オブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="68087-163">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error)
