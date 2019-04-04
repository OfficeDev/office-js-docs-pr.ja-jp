---
ms.date: 03/21/2019
description: JavaScriptを使用し、Excelのカスタム関数でダイアログボックスを作成します。
title: カスタム関数ダイアログ（プレビュー）
localization_priority: Priority
ms.openlocfilehash: 0f596825a7a32525a68ef45656f1390196146706
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30926672"
---
# <a name="display-a-dialog-box-in-custom-functions"></a>カスタム関数でダイアログボックスを表示します

カスタム関数がユーザーと対話する必要がある場合、 `OfficeRuntime.Dialog`オブジェクトを使用してダイアログボックスを作成できます。 ダイアログ ボックスを使用するための一般的なシナリオは、カスタム関数が web サービスにアクセスできるよう、ユーザーを認証することです。 カスタム関数を使用した認証について詳しくは、[カスタム関数認証](./custom-functions-authentication.md)を参照してください。

注意：`OfficeRuntime.Dialog` オブジェクトはカスタム関数ランタイムの一部です。 作業ウィンドウの文脈からは使用できません。 作業ウィンドウからダイアログを作成するには、[ダイアログAPI](/office/dev/add-ins/develop/dialog-api-in-office-add-ins)を参照してください。

## <a name="dialog-api-example"></a>ダイアログ API の使用例

以下のコード サンプルでは、関数 `getTokenViaDialog` がダイアログ API の `displayWebDialog` 関数を使用して、ダイアログ ボックスを表示します。

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"

function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://www.contoso.com/auth")
    .then(function (token) {

      // Use token to get stock price
      fetch("https://www.contoso.com/?token=token&ticker= + ticker")
      .then(function (result) {

        // Return stock price to cell
        resolve(result);
      });
    })
    .catch(function (error) {
      reject(error);
    });
  });
  
  //Helper
  function getToken(url) {
    return new Promise(function (resolve,reject) {
      if(_cachedToken) {
        resolve(_cachedToken);
      } else {
        getTokenViaDialog(url)
        .then(function (result) {
          resolve(result);
        })
        .catch(function (result) {
          reject(result);
        });
      }
    });
  }

  function getTokenViaDialog(url) {
    return new Promise (function (resolve, reject) {
      if (_dialogOpen) {
        // Can only have one dialog open at once, wait for previous dialog's token
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
}
```

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [カスタム関数の変更ログ](custom-functions-changelog.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
