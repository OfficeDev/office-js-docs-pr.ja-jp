---
title: エラー処理
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 23a70b1d66befb971c3c1394eb9162c19f2ee176
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348087"
---
# <a name="error-handling"></a><span data-ttu-id="4faca-102">エラー処理</span><span class="sxs-lookup"><span data-stu-id="4faca-102">Error handling</span></span>

<span data-ttu-id="4faca-103">Excel JavaScript API を使用してアドインを作成する場合は、実行時エラーを考慮するために、エラー処理ロジックを含めます。</span><span class="sxs-lookup"><span data-stu-id="4faca-103">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="4faca-104">これは、API の非同期性のために重要になります。</span><span class="sxs-lookup"><span data-stu-id="4faca-104">Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="4faca-105">**sync()** メソッドと Excel JavaScript API の非同期性の詳細については、「[Excel JavaScript API の中心概念](excel-add-ins-core-concepts.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4faca-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="4faca-106">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="4faca-106">Best practices</span></span>

<span data-ttu-id="4faca-107">このドキュメントのコード サンプルでは、`Excel.run` へのすべての呼び出しが、`Excel.run` 内で発生したエラーを検出するために `catch` 文に付随して起こっていることがわかります。</span><span class="sxs-lookup"><span data-stu-id="4faca-107">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`.</span></span> <span data-ttu-id="4faca-108">Excel JavaScript Api を使用してアドインを作成する場合は、同じパターンを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="4faca-108">We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="4faca-109">API エラー</span><span class="sxs-lookup"><span data-stu-id="4faca-109">API errors</span></span> 

<span data-ttu-id="4faca-110">Excel JavaScript API 要求が正常に実行されない場合、API は次のプロパティを含むエラー オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4faca-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span> 

- <span data-ttu-id="4faca-111">**code**: エラー メッセージの `code` プロパティには、`OfficeExtension.ErrorCodes` または `Excel.ErrorCodes` リストの一部である文字列が含まれます。</span><span class="sxs-lookup"><span data-stu-id="4faca-111">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span></span> <span data-ttu-id="4faca-112">たとえば、エラー コード "InvalidReference" は、参照が指定された操作に対して有効でないことを示します。</span><span class="sxs-lookup"><span data-stu-id="4faca-112">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span></span> <span data-ttu-id="4faca-113">エラー コードはローカライズされません。</span><span class="sxs-lookup"><span data-stu-id="4faca-113">Error codes are not localized.</span></span> 

- <span data-ttu-id="4faca-114">**message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれます。</span><span class="sxs-lookup"><span data-stu-id="4faca-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="4faca-115">このエラー メッセージは、エンド ユーザーが使用するためのものではありません。エラー コードと適切なビジネス ロジックを使用して、アドインがエンド ユーザーに表示するエラー メッセージを判断する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4faca-115">The error message is not intended for end-user consumption; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end-users.</span></span>

- <span data-ttu-id="4faca-116">**debugInfo**: 存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="4faca-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span> 

> [!NOTE]
> <span data-ttu-id="4faca-117">`console.log()` を使用してエラー メッセージをコンソールに出力すると、それらのメッセージはサーバー上でのみ表示されます。</span><span class="sxs-lookup"><span data-stu-id="4faca-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="4faca-118">これらのエラー メッセージが、アドインの作業ウィンドウやホスト アプリケーション内のいずれかの場所で、エンド ユーザーに対して表示されることはありません。</span><span class="sxs-lookup"><span data-stu-id="4faca-118">End-users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="see-also"></a><span data-ttu-id="4faca-119">関連項目</span><span class="sxs-lookup"><span data-stu-id="4faca-119">See also</span></span>

- [<span data-ttu-id="4faca-120">Excel JavaScript API の中心概念</span><span class="sxs-lookup"><span data-stu-id="4faca-120">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="4faca-121">OfficeExtension.Error オブジェクト (Excel の JavaScript API)</span><span class="sxs-lookup"><span data-stu-id="4faca-121">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
