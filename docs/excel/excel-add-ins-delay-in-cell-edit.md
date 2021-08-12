---
title: セルの編集中に実行を遅らせる
description: セルの編集中に Excel.run メソッドの実行を遅延する方法について説明します。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 41bbfba3894bcef0c1fd075ce76557dfdc4ba4721b7bc7b19ca21756b86ccc4d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084285"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>セルの編集中に実行を遅らせる

`Excel.run`を使用するオーバーロード[Excel。RunOptions](/javascript/api/excel/excel.runoptions)オブジェクト。 これには、関数の実行時にプラットフォームの動作に影響を与えるプロパティのセットが含まれています。 現在、次のプロパティがサポートされています。

- `delayForCellEdit`: ユーザーがセル編集モードを終了するまでバッチ要求を延期するかどうかを指定します。 **true** の場合、バッチ要求は延期され、ユーザーがセル編集モードを終了した時点で実行されます。 **false** の場合、バッチ要求は、ユーザーがセル編集モードにある場合、自動的に失敗します (ユーザーにエラーが表示されます)。 `delayForCellEdit` プロパティが指定されていない場合の既定の動作は、このプロパティが **false** の場合と同じ動作となります。

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
