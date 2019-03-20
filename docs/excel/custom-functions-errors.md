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
# <a name="error-handling-within-custom-functions"></a>カスタム関数内でのエラー処理

カスタム関数を定義するアドインをビルドする場合は、実行時エラーを考慮して、エラー処理ロジックを含めるようにします。 カスタム関数のエラー処理は、[全体的な Excel の JavaScript API のエラー処理](excel-add-ins-error-handling.md)と同じです。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

次のコード サンプルでは、`.catch` がコード内で以前に発生したエラーを処理します。

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

## <a name="see-also"></a>関連項目

* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [カスタム関数の変更ログ](custom-functions-changelog.md)
