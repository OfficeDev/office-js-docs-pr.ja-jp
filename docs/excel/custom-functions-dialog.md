---
ms.date: 03/21/2019
description: JavaScriptを使用し、Excelのカスタム関数でダイアログボックスを作成します。
title: カスタム関数ダイアログ（プレビュー）
localization_priority: Priority
ms.openlocfilehash: 0f596825a7a32525a68ef45656f1390196146706
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449261"
---
# <a name="display-a-dialog-box-in-custom-functions"></a><span data-ttu-id="a0eb8-103">カスタム関数でダイアログボックスを表示します</span><span class="sxs-lookup"><span data-stu-id="a0eb8-103">Display a dialog box in custom functions</span></span>

<span data-ttu-id="a0eb8-104">カスタム関数がユーザーと対話する必要がある場合、 `OfficeRuntime.Dialog`オブジェクトを使用してダイアログボックスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="a0eb8-104">If your custom function needs to interact with the user, you can create a dialog box using the `OfficeRuntime.Dialog` object.</span></span> <span data-ttu-id="a0eb8-105">ダイアログ ボックスを使用するための一般的なシナリオは、カスタム関数が web サービスにアクセスできるよう、ユーザーを認証することです。</span><span class="sxs-lookup"><span data-stu-id="a0eb8-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="a0eb8-106">カスタム関数を使用した認証について詳しくは、[カスタム関数認証](./custom-functions-authentication.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a0eb8-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

<span data-ttu-id="a0eb8-107">注意：`OfficeRuntime.Dialog` オブジェクトはカスタム関数ランタイムの一部です。</span><span class="sxs-lookup"><span data-stu-id="a0eb8-107">Note: The `OfficeRuntime.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="a0eb8-108">作業ウィンドウの文脈からは使用できません。</span><span class="sxs-lookup"><span data-stu-id="a0eb8-108">It cannot be used from the context of a task pane.</span></span> <span data-ttu-id="a0eb8-109">作業ウィンドウからダイアログを作成するには、[ダイアログAPI](/office/dev/add-ins/develop/dialog-api-in-office-add-ins)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a0eb8-109">To create a dialog from a task pane, see [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span></span>

## <a name="dialog-api-example"></a><span data-ttu-id="a0eb8-110">ダイアログ API の使用例</span><span class="sxs-lookup"><span data-stu-id="a0eb8-110">Dialog API example</span></span>

<span data-ttu-id="a0eb8-111">以下のコード サンプルでは、関数 `getTokenViaDialog` がダイアログ API の `displayWebDialog` 関数を使用して、ダイアログ ボックスを表示します。</span><span class="sxs-lookup"><span data-stu-id="a0eb8-111">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `displayWebDialog` function to display a dialog box.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="a0eb8-112">関連項目</span><span class="sxs-lookup"><span data-stu-id="a0eb8-112">See also</span></span>

* [<span data-ttu-id="a0eb8-113">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="a0eb8-113">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="a0eb8-114">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="a0eb8-114">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="a0eb8-115">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="a0eb8-115">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="a0eb8-116">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="a0eb8-116">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="a0eb8-117">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="a0eb8-117">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
