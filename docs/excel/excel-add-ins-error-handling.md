---
title: エラー処理
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: e3732af26aeaa6129a4b98d6cbb8e3caf501141f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325109"
---
# <a name="error-handling"></a><span data-ttu-id="94e84-102">エラー処理</span><span class="sxs-lookup"><span data-stu-id="94e84-102">Error handling</span></span>

<span data-ttu-id="94e84-p101">Excel JavaScript API を使用してアドインを作成する場合は、実行時エラーを考慮するために、エラー処理ロジックを含めます。 これは、API の非同期性のために重要になります。</span><span class="sxs-lookup"><span data-stu-id="94e84-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="94e84-105">`sync()`メソッドと EXCEL javascript api の非同期性の詳細については、「 [excel javascript api を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="94e84-105">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="94e84-106">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="94e84-106">Best practices</span></span>

<span data-ttu-id="94e84-p102">このドキュメントのコード サンプルでは、`Excel.run` へのすべての呼び出しに、`catch` 内で発生したエラーを検出するための `Excel.run` ステートメントが付いていることがわかります。 Excel JavaScript Api を使用してアドインを構築するときには、同じパターンを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="94e84-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="94e84-109">API エラー</span><span class="sxs-lookup"><span data-stu-id="94e84-109">API errors</span></span>

<span data-ttu-id="94e84-110">Excel JavaScript API 要求が正常に実行されない場合、API は次のプロパティを含むエラー オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="94e84-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="94e84-p103">**code**:エラー メッセージの `code` プロパティには、`OfficeExtension.ErrorCodes` または `Excel.ErrorCodes` リストの一部である文字列が含まれます。 たとえば、エラー コード "InvalidReference" は、参照が指定された操作に対して有効でないことを示します。 エラー コードはローカライズされません。</span><span class="sxs-lookup"><span data-stu-id="94e84-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="94e84-114">**message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれています。</span><span class="sxs-lookup"><span data-stu-id="94e84-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="94e84-115">このエラー メッセージは、エンド ユーザーが使用するためのものではありません。アドインによってエンド ユーザーに表示されるエラー メッセージは、エラー コードと適切なビジネス ロジックを使用して、判断する必要があります。</span><span class="sxs-lookup"><span data-stu-id="94e84-115">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="94e84-116">**debugInfo**:存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="94e84-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="94e84-117">`console.log()` を使用してエラー メッセージをコンソールに出力すると、それらのメッセージはサーバー上でのみ表示されます。</span><span class="sxs-lookup"><span data-stu-id="94e84-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="94e84-118">これらのエラー メッセージが、アドインの作業ウィンドウやホスト アプリケーション内のいずれかの場所で、エンド ユーザーに対して表示されることはありません。</span><span class="sxs-lookup"><span data-stu-id="94e84-118">End users will not see those error messages in the add-in task pane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="94e84-119">エラー メッセージ</span><span class="sxs-lookup"><span data-stu-id="94e84-119">Error Messages</span></span>

<span data-ttu-id="94e84-120">次の表は、API から返される可能性のあるエラー一覧です。</span><span class="sxs-lookup"><span data-stu-id="94e84-120">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="94e84-121">error.code</span><span class="sxs-lookup"><span data-stu-id="94e84-121">error.code</span></span> | <span data-ttu-id="94e84-122">error.message</span><span class="sxs-lookup"><span data-stu-id="94e84-122">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="94e84-123">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="94e84-123">InvalidArgument</span></span> |<span data-ttu-id="94e84-124">引数が無効であるか、存在しません。または形式が正しくありません。</span><span class="sxs-lookup"><span data-stu-id="94e84-124">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="94e84-125">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="94e84-125">InvalidRequest</span></span>  |<span data-ttu-id="94e84-126">要求を処理できません。</span><span class="sxs-lookup"><span data-stu-id="94e84-126">Cannot process the request.</span></span>|
|<span data-ttu-id="94e84-127">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="94e84-127">InvalidReference</span></span>|<span data-ttu-id="94e84-128">この参照は、現在の操作に対して無効です。</span><span class="sxs-lookup"><span data-stu-id="94e84-128">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="94e84-129">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="94e84-129">InvalidBinding</span></span>  |<span data-ttu-id="94e84-130">このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。</span><span class="sxs-lookup"><span data-stu-id="94e84-130">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="94e84-131">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="94e84-131">InvalidSelection</span></span>|<span data-ttu-id="94e84-132">現在の選択内容は、この操作では無効です。</span><span class="sxs-lookup"><span data-stu-id="94e84-132">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="94e84-133">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="94e84-133">Unauthenticated</span></span> |<span data-ttu-id="94e84-134">必要な認証情報が見つからないか、無効です。</span><span class="sxs-lookup"><span data-stu-id="94e84-134">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="94e84-135">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="94e84-135">AccessDenied</span></span> |<span data-ttu-id="94e84-136">要求された操作を実行できません。</span><span class="sxs-lookup"><span data-stu-id="94e84-136">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="94e84-137">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="94e84-137">ItemNotFound</span></span> |<span data-ttu-id="94e84-138">要求されたリソースは存在しません。</span><span class="sxs-lookup"><span data-stu-id="94e84-138">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="94e84-139">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="94e84-139">ActivityLimitReached</span></span>|<span data-ttu-id="94e84-140">アクティビティの制限に達しました。</span><span class="sxs-lookup"><span data-stu-id="94e84-140">Activity limit has been reached.</span></span>|
|<span data-ttu-id="94e84-141">GeneralException</span><span class="sxs-lookup"><span data-stu-id="94e84-141">GeneralException</span></span>|<span data-ttu-id="94e84-142">要求の処理中に内部エラーが発生しました。</span><span class="sxs-lookup"><span data-stu-id="94e84-142">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="94e84-143">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="94e84-143">NotImplemented</span></span>  |<span data-ttu-id="94e84-144">要求された機能は実装されていません。</span><span class="sxs-lookup"><span data-stu-id="94e84-144">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="94e84-145">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="94e84-145">ServiceNotAvailable</span></span>|<span data-ttu-id="94e84-146">サービスを利用できません。</span><span class="sxs-lookup"><span data-stu-id="94e84-146">The service is unavailable.</span></span>|
|<span data-ttu-id="94e84-147">Conflict</span><span class="sxs-lookup"><span data-stu-id="94e84-147">Conflict</span></span>|<span data-ttu-id="94e84-148">競合のため、要求を処理できませんでした。</span><span class="sxs-lookup"><span data-stu-id="94e84-148">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="94e84-149">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="94e84-149">ItemAlreadyExists</span></span>|<span data-ttu-id="94e84-150">作成中のリソースはすでに存在しています。</span><span class="sxs-lookup"><span data-stu-id="94e84-150">The resource being created already exists.</span></span>|
|<span data-ttu-id="94e84-151">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="94e84-151">UnsupportedOperation</span></span>|<span data-ttu-id="94e84-152">試行中の操作はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="94e84-152">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="94e84-153">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="94e84-153">RequestAborted</span></span>|<span data-ttu-id="94e84-154">実行時に要求が中止されました。</span><span class="sxs-lookup"><span data-stu-id="94e84-154">The request was aborted during run time.</span></span>|
|<span data-ttu-id="94e84-155">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="94e84-155">ApiNotAvailable</span></span>|<span data-ttu-id="94e84-156">要求された API は使用できません。</span><span class="sxs-lookup"><span data-stu-id="94e84-156">The requested API is not available.</span></span>|
|<span data-ttu-id="94e84-157">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="94e84-157">InsertDeleteConflict</span></span>|<span data-ttu-id="94e84-158">試行された挿入操作または削除操作で競合が発生しました。</span><span class="sxs-lookup"><span data-stu-id="94e84-158">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="94e84-159">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="94e84-159">InvalidOperation</span></span>|<span data-ttu-id="94e84-160">試行された操作は、このオブジェクトでは無効です。</span><span class="sxs-lookup"><span data-stu-id="94e84-160">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="94e84-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="94e84-161">See also</span></span>

- [<span data-ttu-id="94e84-162">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="94e84-162">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="94e84-163">OfficeExtension.Error オブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="94e84-163">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error)
