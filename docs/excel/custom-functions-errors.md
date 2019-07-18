---
ms.date: 06/18/2019
description: Excel のカスタム関数でエラーを処理します。
title: 'Excel のカスタム関数でのエラー処理 '
localization_priority: Priority
ms.openlocfilehash: 30c83ea930b16e717b48b9c02ffa0e278eb78b36
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771576"
---
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="1ff3a-103">カスタム関数内でのエラー処理</span><span class="sxs-lookup"><span data-stu-id="1ff3a-103">Error handling within custom functions</span></span>

<span data-ttu-id="1ff3a-104">カスタム関数を定義するアドインをビルドする場合は、実行時エラーを考慮して、エラー処理ロジックを含めるようにします。</span><span class="sxs-lookup"><span data-stu-id="1ff3a-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="1ff3a-105">カスタム関数のエラー処理は、[全体的な Excel の JavaScript API のエラー処理](excel-add-ins-error-handling.md)と同じです。</span><span class="sxs-lookup"><span data-stu-id="1ff3a-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="1ff3a-106">次のコード サンプルでは、`.catch` がコード内で以前に発生したエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="1ff3a-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="next-steps"></a><span data-ttu-id="1ff3a-107">次の手順</span><span class="sxs-lookup"><span data-stu-id="1ff3a-107">Next steps</span></span>
<span data-ttu-id="1ff3a-108">[自分のカスタム関数で問題をトラブルシューティングを行う](custom-functions-troubleshooting.md)方法についての詳細を確認する。</span><span class="sxs-lookup"><span data-stu-id="1ff3a-108">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="1ff3a-109">関連項目</span><span class="sxs-lookup"><span data-stu-id="1ff3a-109">See also</span></span>

* [<span data-ttu-id="1ff3a-110">カスタム関数のデバッグ</span><span class="sxs-lookup"><span data-stu-id="1ff3a-110">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="1ff3a-111">カスタム関数の要件</span><span class="sxs-lookup"><span data-stu-id="1ff3a-111">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="1ff3a-112">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="1ff3a-112">Create custom functions in Excel</span></span>](custom-functions-overview.md)
