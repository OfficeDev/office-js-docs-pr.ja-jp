---
ms.date: 06/18/2019
description: JavaScript を使用して Excel のカスタム関数でダイアログ ボックスを作成します。
title: カスタム関数からダイアログ ボックスを表示する
localization_priority: Normal
ms.openlocfilehash: a2ef005f4c1519228f114dbd671d689807e5914c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718747"
---
# <a name="display-a-dialog-box-from-a-custom-function"></a>カスタム関数からダイアログ ボックスを表示する

ユーザーがカスタム関数を操作する必要がある場合は、[`Office.Dialog` オブジェクト](/javascript/api/office-runtime/officeruntime.dialog)を使用してダイアログ ボックスを作成できます。 ダイアログ ボックスを使用するための一般的なシナリオは、カスタム関数が web サービスにアクセスできるよう、ユーザーを認証することです。 カスタム関数を使用した認証について詳しくは、[カスタム関数認証](./custom-functions-authentication.md)を参照してください。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> `Office.Dialog` オブジェクトは、カスタム関数のランタイムの一部です。 作業ウィンドウは `Dialog` オブジェクトを使用しません。 作業ウィンドウからダイアログ ボックスを作成するには、[ダイアログ API](../develop/dialog-api-in-office-add-ins.md) を参照してください。

## <a name="dialog-box-api-example"></a>ダイアログ ボックス API の例

次のコードサンプルでは、関数`getTokenViaDialog`は`Dialog` API の`displayWebDialogOptions`関数を使用してダイアログボックスを表示します。

```js
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once, wait for previous dialog box's token
      let timeout = 5;
      let count = 0;
      var intervalId = setInterval(function () {
        count++;
        if(_cachedToken) {
          resolve(_cachedToken);
          clearInterval(intervalId);
        }
        if(count >= timeout) {
          reject("Timeout while waiting for token");
          clearInterval(intervalId);
        }
      }, 1000);
    } else {
      _dialogOpen = true;
      OfficeRuntime.displayWebDialog(url, {
        height: '50%',
        width: '50%',
        onMessage: function (message, dialog) {
          _cachedToken = message;
          resolve(message);
          dialog.close();
          return;
        },
        onRuntimeError: function(error, dialog) {
          reject(error);
        },
      }).catch(function (e) {
        reject(e);
      });
    }
  });
}
```

## <a name="next-steps"></a>次の手順
「[XLL ユーザー定義関数と互換性のある、カスタム関数を作成する](make-custom-functions-compatible-with-xll-udf.md)」で方法を確認してください。

## <a name="see-also"></a>関連項目

* [カスタム関数の認証](custom-functions-authentication.md)
* [カスタム関数でデータを受信して​​処理する](custom-functions-web-reqs.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
