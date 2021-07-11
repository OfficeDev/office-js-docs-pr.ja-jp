---
title: セルの編集中に実行を遅らせる
description: セルの編集中に Excel.run メソッドの実行を遅延する方法について説明します。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: b7b28064ef4d313639391e63cba780351b5623f9
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349519"
---
# <a name="delay-execution-while-cell-is-being-edited"></a><span data-ttu-id="bc31d-103">セルの編集中に実行を遅らせる</span><span class="sxs-lookup"><span data-stu-id="bc31d-103">Delay execution while cell is being edited</span></span>

<span data-ttu-id="bc31d-104">`Excel.run`を使用するオーバーロード[Excel。RunOptions](/javascript/api/excel/excel.runoptions)オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="bc31d-104">`Excel.run` has an overload that takes in a [Excel.RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="bc31d-105">これには、関数の実行時にプラットフォームの動作に影響を与えるプロパティのセットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="bc31d-105">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="bc31d-106">現在、次のプロパティがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="bc31d-106">The following property is currently supported.</span></span>

- <span data-ttu-id="bc31d-107">`delayForCellEdit`: ユーザーがセル編集モードを終了するまでバッチ要求を延期するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="bc31d-107">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="bc31d-108">**true** の場合、バッチ要求は延期され、ユーザーがセル編集モードを終了した時点で実行されます。</span><span class="sxs-lookup"><span data-stu-id="bc31d-108">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="bc31d-109">**false** の場合、バッチ要求は、ユーザーがセル編集モードにある場合、自動的に失敗します (ユーザーにエラーが表示されます)。</span><span class="sxs-lookup"><span data-stu-id="bc31d-109">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="bc31d-110">`delayForCellEdit` プロパティが指定されていない場合の既定の動作は、このプロパティが **false** の場合と同じ動作となります。</span><span class="sxs-lookup"><span data-stu-id="bc31d-110">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
