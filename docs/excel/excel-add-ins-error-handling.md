---
title: エラー処理
description: ランタイムエラーを考慮した Excel JavaScript API のエラー処理ロジックについて説明します。
ms.date: 05/13/2020
localization_priority: Normal
ms.openlocfilehash: ff6336e4d76e84ddc708d65eda70f5f2e172fde7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609609"
---
# <a name="error-handling"></a><span data-ttu-id="3e5d4-103">エラー処理</span><span class="sxs-lookup"><span data-stu-id="3e5d4-103">Error handling</span></span>

<span data-ttu-id="3e5d4-p101">Excel JavaScript API を使用してアドインを作成する場合は、実行時エラーを考慮するために、エラー処理ロジックを含めます。 これは、API の非同期性のために重要になります。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="3e5d4-106">`sync()`メソッドと Excel JAVASCRIPT api の非同期性の詳細については、「 [EXCEL javascript api を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="3e5d4-107">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="3e5d4-107">Best practices</span></span>

<span data-ttu-id="3e5d4-p102">このドキュメントのコード サンプルでは、`Excel.run` へのすべての呼び出しに、`catch` 内で発生したエラーを検出するための `Excel.run` ステートメントが付いていることがわかります。 Excel JavaScript Api を使用してアドインを構築するときには、同じパターンを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="3e5d4-110">API エラー</span><span class="sxs-lookup"><span data-stu-id="3e5d4-110">API errors</span></span>

<span data-ttu-id="3e5d4-111">Excel JavaScript API 要求が正常に実行されない場合、API は次のプロパティを含むエラー オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="3e5d4-p103">**code**:エラー メッセージの `code` プロパティには、`OfficeExtension.ErrorCodes` または `Excel.ErrorCodes` リストの一部である文字列が含まれます。 たとえば、エラー コード "InvalidReference" は、参照が指定された操作に対して有効でないことを示します。 エラー コードはローカライズされません。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="3e5d4-115">**message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれています。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="3e5d4-116">このエラー メッセージは、エンド ユーザーが使用するためのものではありません。アドインによってエンド ユーザーに表示されるエラー メッセージは、エラー コードと適切なビジネス ロジックを使用して、判断する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="3e5d4-117">**debugInfo**:存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="3e5d4-118">`console.log()` を使用してエラー メッセージをコンソールに出力すると、それらのメッセージはサーバー上でのみ表示されます。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="3e5d4-119">これらのエラー メッセージが、アドインの作業ウィンドウやホスト アプリケーション内のいずれかの場所で、エンド ユーザーに対して表示されることはありません。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-119">End users will not see those error messages in the add-in task pane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="3e5d4-120">エラー メッセージ</span><span class="sxs-lookup"><span data-stu-id="3e5d4-120">Error Messages</span></span>

<span data-ttu-id="3e5d4-121">次の表は、API から返される可能性のあるエラー一覧です。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="3e5d4-122">error.code</span><span class="sxs-lookup"><span data-stu-id="3e5d4-122">error.code</span></span> | <span data-ttu-id="3e5d4-123">error.message</span><span class="sxs-lookup"><span data-stu-id="3e5d4-123">error.message</span></span> |
|:----------|:--------------|
|`AccessDenied` |<span data-ttu-id="3e5d4-124">要求された操作を実行できません。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-124">You cannot perform the requested operation.</span></span>|
|`ActivityLimitReached`|<span data-ttu-id="3e5d4-125">アクティビティの制限に達しました。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-125">Activity limit has been reached.</span></span>|
|`ApiNotAvailable`|<span data-ttu-id="3e5d4-126">要求された API は使用できません。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-126">The requested API is not available.</span></span>|
|`Conflict`|<span data-ttu-id="3e5d4-127">競合のため、要求を処理できませんでした。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-127">Request could not be processed because of a conflict.</span></span>|
|`GeneralException`|<span data-ttu-id="3e5d4-128">要求の処理中に内部エラーが発生しました。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-128">There was an internal error while processing the request.</span></span>|
|`InsertDeleteConflict`|<span data-ttu-id="3e5d4-129">試行された挿入操作または削除操作で競合が発生しました。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-129">The insert or delete operation attempted resulted in a conflict.</span></span>|
|`InvalidArgument` |<span data-ttu-id="3e5d4-130">引数が無効であるか、存在しません。または形式が正しくありません。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-130">The argument is invalid or missing or has an incorrect format.</span></span>|
|`InvalidBinding`  |<span data-ttu-id="3e5d4-131">このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-131">This object binding is no longer valid due to previous updates.</span></span>|
|`InvalidOperation`|<span data-ttu-id="3e5d4-132">試行された操作は、このオブジェクトでは無効です。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-132">The operation attempted is invalid on the object.</span></span>|
|`InvalidReference`|<span data-ttu-id="3e5d4-133">この参照は、現在の操作に対して無効です。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-133">This reference is not valid for the current operation.</span></span>|
|`InvalidRequest`  |<span data-ttu-id="3e5d4-134">要求を処理できません。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-134">Cannot process the request.</span></span>|
|`InvalidSelection`|<span data-ttu-id="3e5d4-135">現在の選択内容は、この操作では無効です。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-135">The current selection is invalid for this operation.</span></span>|
|`ItemAlreadyExists`|<span data-ttu-id="3e5d4-136">作成中のリソースはすでに存在しています。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-136">The resource being created already exists.</span></span>|
|`ItemNotFound` |<span data-ttu-id="3e5d4-137">要求されたリソースは存在しません。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-137">The requested resource doesn't exist.</span></span>|
|`NotImplemented`  |<span data-ttu-id="3e5d4-138">要求された機能は実装されていません。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-138">The requested feature isn't implemented.</span></span>|
|`RequestAborted`|<span data-ttu-id="3e5d4-139">実行時に要求が中止されました。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-139">The request was aborted during run time.</span></span>|
|`ServiceNotAvailable`|<span data-ttu-id="3e5d4-140">サービスを利用できません。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-140">The service is unavailable.</span></span>|
|`Unauthenticated` |<span data-ttu-id="3e5d4-141">必要な認証情報が見つからないか、無効です。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-141">Required authentication information is either missing or invalid.</span></span>|
|`UnsupportedOperation`|<span data-ttu-id="3e5d4-142">試行中の操作はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3e5d4-142">The operation being attempted is not supported.</span></span>|

## <a name="see-also"></a><span data-ttu-id="3e5d4-143">関連項目</span><span class="sxs-lookup"><span data-stu-id="3e5d4-143">See also</span></span>

- [<span data-ttu-id="3e5d4-144">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="3e5d4-144">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="3e5d4-145">OfficeExtension.Error オブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="3e5d4-145">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error?view=excel-js-preview)
