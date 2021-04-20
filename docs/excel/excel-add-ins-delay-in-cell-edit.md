---
title: セルの編集中に実行を延期する
description: セルが編集されているときに、Excel の run メソッドの実行を延期する方法について説明します。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: eb33f4cb7cce3b1f8642e00f432e708e90b5b895
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409405"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>セルの編集中に実行を延期する

`Excel.run`[RunOptions](/javascript/api/excel/excel.runoptions)オブジェクトで実行されるオーバーロードがあります。 これには、関数の実行時にプラットフォームの動作に影響を与えるプロパティのセットが含まれています。 次のプロパティが現在サポートされています。

* `delayForCellEdit`: ユーザーがセル編集モードを終了するまでバッチ要求を延期するかどうかを指定します。 **true** の場合、バッチ要求は延期され、ユーザーがセル編集モードを終了した時点で実行されます。 **false** の場合、バッチ要求は、ユーザーがセル編集モードにある場合、自動的に失敗します (ユーザーにエラーが表示されます)。 `delayForCellEdit` プロパティが指定されていない場合の既定の動作は、このプロパティが **false** の場合と同じ動作となります。

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
