---
title: JavaScript API のExcel処理
description: ランタイム エラーをExcel JavaScript API エラー処理ロジックについて説明します。
ms.date: 01/15/2021
localization_priority: Normal
ms.openlocfilehash: 42ef52b5d20a2c2d1284f57c7b4026ff2c71ebdd
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349512"
---
# <a name="error-handling-with-the-excel-javascript-api"></a><span data-ttu-id="27b59-103">JavaScript API のExcel処理</span><span class="sxs-lookup"><span data-stu-id="27b59-103">Error handling with the Excel JavaScript API</span></span>

<span data-ttu-id="27b59-p101">Excel JavaScript API を使用してアドインを作成する場合は、実行時エラーを考慮するために、エラー処理ロジックを含めます。 これは、API の非同期性のために重要になります。</span><span class="sxs-lookup"><span data-stu-id="27b59-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="27b59-106">JavaScript API のメソッドと非同期の性質の詳細については、「Excel アドイン」の `sync()` [「Excel JavaScript](excel-add-ins-core-concepts.md)オブジェクト モデルOffice参照してください。</span><span class="sxs-lookup"><span data-stu-id="27b59-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="27b59-107">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="27b59-107">Best practices</span></span>

<span data-ttu-id="27b59-p102">このドキュメントのコード サンプルでは、`Excel.run` へのすべての呼び出しに、`catch` 内で発生したエラーを検出するための `Excel.run` ステートメントが付いていることがわかります。 Excel JavaScript Api を使用してアドインを構築するときには、同じパターンを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="27b59-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="27b59-110">API エラー</span><span class="sxs-lookup"><span data-stu-id="27b59-110">API errors</span></span>

<span data-ttu-id="27b59-111">JavaScript API Excelが正常に実行できない場合、API は次のプロパティを含むエラー オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="27b59-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties.</span></span>

- <span data-ttu-id="27b59-p103">**code**:エラー メッセージの `code` プロパティには、`OfficeExtension.ErrorCodes` または `Excel.ErrorCodes` リストの一部である文字列が含まれます。 たとえば、エラー コード "InvalidReference" は、参照が指定された操作に対して有効でないことを示します。 エラー コードはローカライズされません。</span><span class="sxs-lookup"><span data-stu-id="27b59-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="27b59-115">**message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれています。</span><span class="sxs-lookup"><span data-stu-id="27b59-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="27b59-116">このエラー メッセージは、エンド ユーザーが使用するためのものではありません。アドインによってエンド ユーザーに表示されるエラー メッセージは、エラー コードと適切なビジネス ロジックを使用して、判断する必要があります。</span><span class="sxs-lookup"><span data-stu-id="27b59-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="27b59-117">**debugInfo**:存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="27b59-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="27b59-118">`console.log()` を使用してエラー メッセージをコンソールに出力すると、それらのメッセージはサーバー上でのみ表示されます。</span><span class="sxs-lookup"><span data-stu-id="27b59-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="27b59-119">エンド ユーザーは、アドインの作業ウィンドウまたはアプリケーション内の任意の場所にこれらのエラー メッセージOfficeしません。</span><span class="sxs-lookup"><span data-stu-id="27b59-119">End users will not see those error messages in the add-in task pane or anywhere in the Office application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="27b59-120">エラー メッセージ</span><span class="sxs-lookup"><span data-stu-id="27b59-120">Error Messages</span></span>

<span data-ttu-id="27b59-121">次の表は、API から返される可能性のあるエラー一覧です。</span><span class="sxs-lookup"><span data-stu-id="27b59-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="27b59-122">エラー コード</span><span class="sxs-lookup"><span data-stu-id="27b59-122">Error code</span></span> | <span data-ttu-id="27b59-123">エラー メッセージ</span><span class="sxs-lookup"><span data-stu-id="27b59-123">Error message</span></span> |
|:----------|:--------------|
|`AccessDenied` |<span data-ttu-id="27b59-124">要求された操作を実行できません。</span><span class="sxs-lookup"><span data-stu-id="27b59-124">You cannot perform the requested operation.</span></span>|
|`ActivityLimitReached`|<span data-ttu-id="27b59-125">アクティビティの制限に達しました。</span><span class="sxs-lookup"><span data-stu-id="27b59-125">Activity limit has been reached.</span></span>|
|`ApiNotAvailable`|<span data-ttu-id="27b59-126">要求された API は使用できません。</span><span class="sxs-lookup"><span data-stu-id="27b59-126">The requested API is not available.</span></span>|
|`ApiNotFound`|<span data-ttu-id="27b59-127">使用しようとしている API が見つかりませんでした。</span><span class="sxs-lookup"><span data-stu-id="27b59-127">The API you are trying to use could not be found.</span></span> <span data-ttu-id="27b59-128">このバージョンは、新しいバージョンの Excel。</span><span class="sxs-lookup"><span data-stu-id="27b59-128">It may be available in a newer version of Excel.</span></span> <span data-ttu-id="27b59-129">詳細については[Excel JavaScript API 要件セットの記事](../reference/requirement-sets/excel-api-requirement-sets.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="27b59-129">See the [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md) article for more information.</span></span>|
|`BadPassword`|<span data-ttu-id="27b59-130">指定したパスワードが正しくありません。</span><span class="sxs-lookup"><span data-stu-id="27b59-130">The password you supplied is incorrect.</span></span>|
|`Conflict`|<span data-ttu-id="27b59-131">競合のため、要求を処理できませんでした。</span><span class="sxs-lookup"><span data-stu-id="27b59-131">Request could not be processed because of a conflict.</span></span>|
|`ContentLengthRequired`|<span data-ttu-id="27b59-132">`Content-length`HTTP ヘッダーが見つからない。</span><span class="sxs-lookup"><span data-stu-id="27b59-132">A `Content-length` HTTP header is missing.</span></span>|
|`GeneralException`|<span data-ttu-id="27b59-133">要求の処理中に内部エラーが発生しました。</span><span class="sxs-lookup"><span data-stu-id="27b59-133">There was an internal error while processing the request.</span></span>|
|`InactiveWorkbook`|<span data-ttu-id="27b59-134">複数のブックが開き、この API によって呼び出されるブックがフォーカスを失ったため、操作に失敗しました。</span><span class="sxs-lookup"><span data-stu-id="27b59-134">The operation failed because multiple workbooks are open and the workbook being called by this API has lost focus.</span></span>|
|`InsertDeleteConflict`|<span data-ttu-id="27b59-135">試行された挿入操作または削除操作で競合が発生しました。</span><span class="sxs-lookup"><span data-stu-id="27b59-135">The insert or delete operation attempted resulted in a conflict.</span></span>|
|`InvalidArgument` |<span data-ttu-id="27b59-136">引数が無効であるか、存在しません。または形式が正しくありません。</span><span class="sxs-lookup"><span data-stu-id="27b59-136">The argument is invalid or missing or has an incorrect format.</span></span>|
|`InvalidBinding`  |<span data-ttu-id="27b59-137">このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。</span><span class="sxs-lookup"><span data-stu-id="27b59-137">This object binding is no longer valid due to previous updates.</span></span>|
|`InvalidOperation`|<span data-ttu-id="27b59-138">試行された操作は、このオブジェクトでは無効です。</span><span class="sxs-lookup"><span data-stu-id="27b59-138">The operation attempted is invalid on the object.</span></span>|
|`InvalidReference`|<span data-ttu-id="27b59-139">この参照は、現在の操作に対して無効です。</span><span class="sxs-lookup"><span data-stu-id="27b59-139">This reference is not valid for the current operation.</span></span>|
|`InvalidRequest`  |<span data-ttu-id="27b59-140">要求を処理できません。</span><span class="sxs-lookup"><span data-stu-id="27b59-140">Cannot process the request.</span></span>|
|`InvalidSelection`|<span data-ttu-id="27b59-141">現在の選択内容は、この操作では無効です。</span><span class="sxs-lookup"><span data-stu-id="27b59-141">The current selection is invalid for this operation.</span></span>|
|`ItemAlreadyExists`|<span data-ttu-id="27b59-142">作成中のリソースはすでに存在しています。</span><span class="sxs-lookup"><span data-stu-id="27b59-142">The resource being created already exists.</span></span>|
|`ItemNotFound` |<span data-ttu-id="27b59-143">要求されたリソースは存在しません。</span><span class="sxs-lookup"><span data-stu-id="27b59-143">The requested resource doesn't exist.</span></span>|
|`NonBlankCellOffSheet`|<span data-ttu-id="27b59-144">Microsoft Excelセルを挿入できないのは、空でないセルをワークシートの最後から押し出すためです。</span><span class="sxs-lookup"><span data-stu-id="27b59-144">Microsoft Excel can't insert new cells because it would push non-empty cells off the end of the worksheet.</span></span> <span data-ttu-id="27b59-145">空でないセルは空に表示されますが、空白の値、書式設定、または数式が含まれます。</span><span class="sxs-lookup"><span data-stu-id="27b59-145">These non-empty cells might appear empty but have blank values, some formatting, or a formula.</span></span> <span data-ttu-id="27b59-146">挿入する行または列を十分に削除してから、もう一度やり直してください。</span><span class="sxs-lookup"><span data-stu-id="27b59-146">Delete enough rows or columns to make room for what you want to insert and then try again.</span></span>|
|`NotImplemented`|<span data-ttu-id="27b59-147">要求された機能は実装されていません。</span><span class="sxs-lookup"><span data-stu-id="27b59-147">The requested feature isn't implemented.</span></span>|
|`RangeExceedsLimit`|<span data-ttu-id="27b59-148">範囲内のセル数がサポートされている最大数を超えました。</span><span class="sxs-lookup"><span data-stu-id="27b59-148">The cell count in the range has exceeded the maximum supported number.</span></span> <span data-ttu-id="27b59-149">詳細については[、「リソースの制限](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)とパフォーマンスの最適化」Officeを参照してください。</span><span class="sxs-lookup"><span data-stu-id="27b59-149">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span>|
|`RequestAborted`|<span data-ttu-id="27b59-150">実行時に要求が中止されました。</span><span class="sxs-lookup"><span data-stu-id="27b59-150">The request was aborted during run time.</span></span>|
|`RequestPayloadSizeLimitExceeded`|<span data-ttu-id="27b59-151">要求ペイロードのサイズが制限を超えています。</span><span class="sxs-lookup"><span data-stu-id="27b59-151">The request payload size has exceeded the limit.</span></span> <span data-ttu-id="27b59-152">詳細については[、「リソースの制限](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)とパフォーマンスの最適化」Officeを参照してください。</span><span class="sxs-lookup"><span data-stu-id="27b59-152">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span> <br><br><span data-ttu-id="27b59-153">このエラーは、このエラーが発生Excel on the web。</span><span class="sxs-lookup"><span data-stu-id="27b59-153">This error only occurs in Excel on the web.</span></span>|
|`ResponsePayloadSizeLimitExceeded`|<span data-ttu-id="27b59-154">応答ペイロードのサイズが制限を超えています。</span><span class="sxs-lookup"><span data-stu-id="27b59-154">The response payload size has exceeded the limit.</span></span> <span data-ttu-id="27b59-155">詳細については[、「リソースの制限](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)とパフォーマンスの最適化」Officeを参照してください。</span><span class="sxs-lookup"><span data-stu-id="27b59-155">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span>  <br><br><span data-ttu-id="27b59-156">このエラーは、このエラーが発生Excel on the web。</span><span class="sxs-lookup"><span data-stu-id="27b59-156">This error only occurs in Excel on the web.</span></span>|
|`ServiceNotAvailable`|<span data-ttu-id="27b59-157">サービスを利用できません。</span><span class="sxs-lookup"><span data-stu-id="27b59-157">The service is unavailable.</span></span>|
|`Unauthenticated` |<span data-ttu-id="27b59-158">必要な認証情報が見つからないか、無効です。</span><span class="sxs-lookup"><span data-stu-id="27b59-158">Required authentication information is either missing or invalid.</span></span>|
|`UnsupportedOperation`|<span data-ttu-id="27b59-159">試行中の操作はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="27b59-159">The operation being attempted is not supported.</span></span>|
|`UnsupportedSheet`|<span data-ttu-id="27b59-160">このシートの種類はマクロ シートまたはグラフ シートで、この操作はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="27b59-160">This sheet type does not support this operation, since it is a Macro or Chart sheet.</span></span>|

> [!NOTE]
> <span data-ttu-id="27b59-161">前の表に、JavaScript API の使用中に発生する可能性Excel示します。</span><span class="sxs-lookup"><span data-stu-id="27b59-161">The preceding table lists error messages you may encounter while using the Excel JavaScript API.</span></span> <span data-ttu-id="27b59-162">アプリケーション固有の Excel JavaScript API の代わりに共通 API を使用している場合は、「Office 一般的な[API](../reference/javascript-api-for-office-error-codes.md)エラー コード」を参照して、関連するエラー メッセージについて説明します。</span><span class="sxs-lookup"><span data-stu-id="27b59-162">If you are working with the Common API instead of the application-specific Excel JavaScript API, see [Office Common API error codes](../reference/javascript-api-for-office-error-codes.md) to learn about relevant error messages.</span></span>

## <a name="see-also"></a><span data-ttu-id="27b59-163">関連項目</span><span class="sxs-lookup"><span data-stu-id="27b59-163">See also</span></span>

- [<span data-ttu-id="27b59-164">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="27b59-164">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="27b59-165">OfficeExtension.Error オブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="27b59-165">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [<span data-ttu-id="27b59-166">Office の一般的な API エラー コード</span><span class="sxs-lookup"><span data-stu-id="27b59-166">Office Common API error codes</span></span>](../reference/javascript-api-for-office-error-codes.md)
