---
title: セルの編集中に実行を遅らせる
description: セルの編集中に Excel.run メソッドの実行を遅らせる方法について説明します。
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1abcdb382150db486033b32d2521207ab0b7f28f
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889220"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>セルの編集中に実行を遅らせる

`Excel.run` には、 [Excel.RunOptions](/javascript/api/excel/excel.runoptions) オブジェクトを受け取るオーバーロードがあります。 これには、関数の実行時にプラットフォームの動作に影響を与えるプロパティのセットが含まれています。 現在、次のプロパティがサポートされています。

- `delayForCellEdit`: ユーザーがセル編集モードを終了するまでバッチ要求を延期するかどうかを指定します。 ときに `true`、バッチ要求が遅延され、ユーザーがセル編集モードを終了したときに実行されます。 この場合 `false`、ユーザーがセル編集モードになっている (エラーがユーザーに到達する原因) 場合、バッチ要求は自動的に失敗します。 プロパティが指定されていない `delayForCellEdit` 既定の動作は、プロパティが指定されている場合 `false`と同じです。

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
