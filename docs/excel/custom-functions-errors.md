---
ms.date: 02/08/2019
description: Excel のカスタム関数でエラーを処理します。
title: Excel のカスタム関数でのエラー処理 (プレビュー)
localization_priority: Priority
ms.openlocfilehash: 170da03331663d6779bed7bf0bf5a9b75b908b3f
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/14/2019
ms.locfileid: "30632696"
---
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="0f76d-103">カスタム関数内でのエラー処理</span><span class="sxs-lookup"><span data-stu-id="0f76d-103">Error handling within custom functions</span></span>

<span data-ttu-id="0f76d-104">カスタム関数を定義するアドインをビルドする場合は、実行時エラーを考慮して、エラー処理ロジックを含めるようにします。</span><span class="sxs-lookup"><span data-stu-id="0f76d-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="0f76d-105">カスタム関数のエラー処理は、[全体的な Excel の JavaScript API のエラー処理](excel-add-ins-error-handling.md)と同じです。</span><span class="sxs-lookup"><span data-stu-id="0f76d-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="0f76d-106">次のコード サンプルでは、`.catch` がコード内で以前に発生したエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="0f76d-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="see-also"></a><span data-ttu-id="0f76d-107">関連項目</span><span class="sxs-lookup"><span data-stu-id="0f76d-107">See also</span></span>

* [<span data-ttu-id="0f76d-108">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="0f76d-108">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="0f76d-109">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="0f76d-109">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="0f76d-110">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="0f76d-110">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="0f76d-111">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="0f76d-111">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="0f76d-112">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="0f76d-112">Custom functions changelog</span></span>](custom-functions-changelog.md)
