---
title: セルの編集中に実行を遅らせる
description: セルの編集中に Excel.run メソッドの実行を遅延する方法について説明します。
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: c5609fbb2a39d6ecc69063d4bccdfbc1da1c102d
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340807"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>セルの編集中に実行を遅らせる

`Excel.run`を使用するオーバーロードを持Excel[。RunOptions](/javascript/api/excel/excel.runoptions) オブジェクト。 これには、関数の実行時にプラットフォームの動作に影響を与えるプロパティのセットが含まれています。 現在、次のプロパティがサポートされています。

- `delayForCellEdit`: ユーザーがセル編集モードを終了するまでバッチ要求を延期するかどうかを指定します。 **true** の場合、バッチ要求は延期され、ユーザーがセル編集モードを終了した時点で実行されます。 **false** の場合、バッチ要求は、ユーザーがセル編集モードにある場合、自動的に失敗します (ユーザーにエラーが表示されます)。 `delayForCellEdit` プロパティが指定されていない場合の既定の動作は、このプロパティが **false** の場合と同じ動作となります。

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
