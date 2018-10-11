---
title: エラー処理
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b07012516cbe15374d0707c157738117a9c8fe96
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459232"
---
# <a name="error-handling"></a><span data-ttu-id="3370c-102">エラー処理</span><span class="sxs-lookup"><span data-stu-id="3370c-102">Error handling</span></span>

<span data-ttu-id="3370c-p101">Excel の JavaScript API を使用してアドインをビルドする場合は、ランタイム エラーを考慮するためのエラー処理ロジックを含めるようにしてください。これは、API の非同期の性質のために重要です。</span><span class="sxs-lookup"><span data-stu-id="3370c-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="3370c-105"> *\*Sync()** メソッドと非同期であるため Excel の JavaScript API の詳細については、 [Excel の JavaScript API を使用して基本的なプログラミングの概念](excel-add-ins-core-concepts.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3370c-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="3370c-106">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="3370c-106">Best practices</span></span>

<span data-ttu-id="3370c-p102">このドキュメントのコード サンプル全体にわたり、`Excel.run`へのすべての呼び出しが`Excel.run`内で発生するエラーをキャッチする`catch`ステートメントに付属していることが分かります。Excel の JavaScript Api を使用してアドインをビルドするときは、同じパターンを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="3370c-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="3370c-109">API エラー</span><span class="sxs-lookup"><span data-stu-id="3370c-109">API errors</span></span> 

<span data-ttu-id="3370c-110">Excel JavaScript API 要求が正常に実行されない場合、API は次のプロパティを含むエラー オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3370c-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span> 

- <span data-ttu-id="3370c-p103">**コード**: `code` エラー メッセージのプロパティが含まれている文字列を含む、 `OfficeExtension.ErrorCodes` または `Excel.ErrorCodes` リストです。たとえば、エラー コード"InvalidReference"では、参照が指定された操作に対して有効ではないことを示します。エラー コードはローカライズされません。</span><span class="sxs-lookup"><span data-stu-id="3370c-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span> 

- <span data-ttu-id="3370c-p104">**メッセージ**: `message` エラー メッセージのプロパティには、ローカライズされた文字列のエラーの概要が含まれています。エラー メッセージは、エンド ユーザーの消費対象ではありません。アドインがエンド ユーザーに示すエラー メッセージを確認するには、エラー コードと適切なビジネス ロジックを使用してください。</span><span class="sxs-lookup"><span data-stu-id="3370c-p104">**message**: The `message` property of an error message contains a summary of the error in the localized string. The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="3370c-116">**debugInfo**: 存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="3370c-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span> 

> [!NOTE]
> <span data-ttu-id="3370c-p105">`console.log()`を使用してコンソールにエラー メッセージを印刷する場合は、サーバー上にそれらのメッセージのみが表示されます。エンド ユーザーには、アドイン作業ウィンドウでまたはホスト アプリケーションの任意の場所にこれらのエラー メッセージは表示されません。</span><span class="sxs-lookup"><span data-stu-id="3370c-p105">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server. End users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="see-also"></a><span data-ttu-id="3370c-119">関連項目</span><span class="sxs-lookup"><span data-stu-id="3370c-119">See also</span></span>

- [<span data-ttu-id="3370c-120">Excel の JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="3370c-120">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="3370c-121">OfficeExtension.Error オブジェクト (Excel JavaScript API)</span><span class="sxs-lookup"><span data-stu-id="3370c-121">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
