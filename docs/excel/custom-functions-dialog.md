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
# <a name="display-a-dialog-box-from-a-custom-function"></a><span data-ttu-id="8a1f5-103">カスタム関数からダイアログ ボックスを表示する</span><span class="sxs-lookup"><span data-stu-id="8a1f5-103">Display a dialog box from a custom function</span></span>

<span data-ttu-id="8a1f5-104">ユーザーがカスタム関数を操作する必要がある場合は、[`Office.Dialog` オブジェクト](/javascript/api/office-runtime/officeruntime.dialog)を使用してダイアログ ボックスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="8a1f5-104">If your custom function needs to interact with the user, you can create a dialog box using the [`Office.Dialog` object](/javascript/api/office-runtime/officeruntime.dialog).</span></span> <span data-ttu-id="8a1f5-105">ダイアログ ボックスを使用するための一般的なシナリオは、カスタム関数が web サービスにアクセスできるよう、ユーザーを認証することです。</span><span class="sxs-lookup"><span data-stu-id="8a1f5-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="8a1f5-106">カスタム関数を使用した認証について詳しくは、[カスタム関数認証](./custom-functions-authentication.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8a1f5-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> <span data-ttu-id="8a1f5-107">`Office.Dialog` オブジェクトは、カスタム関数のランタイムの一部です。</span><span class="sxs-lookup"><span data-stu-id="8a1f5-107">The `Office.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="8a1f5-108">作業ウィンドウは `Dialog` オブジェクトを使用しません。</span><span class="sxs-lookup"><span data-stu-id="8a1f5-108">Task panes don't use the `Dialog` object.</span></span> <span data-ttu-id="8a1f5-109">作業ウィンドウからダイアログ ボックスを作成するには、[ダイアログ API](../develop/dialog-api-in-office-add-ins.md) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8a1f5-109">To create a dialog box from a task pane, see [Dialog API](../develop/dialog-api-in-office-add-ins.md).</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="8a1f5-110">ダイアログ ボックス API の例</span><span class="sxs-lookup"><span data-stu-id="8a1f5-110">dialog box API example</span></span>

<span data-ttu-id="8a1f5-111">次のコードサンプルでは、関数`getTokenViaDialog`は`Dialog` API の`displayWebDialogOptions`関数を使用してダイアログボックスを表示します。</span><span class="sxs-lookup"><span data-stu-id="8a1f5-111">In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API's `displayWebDialogOptions` function to display a dialog box.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="8a1f5-112">次の手順</span><span class="sxs-lookup"><span data-stu-id="8a1f5-112">Next steps</span></span>
<span data-ttu-id="8a1f5-113">「[XLL ユーザー定義関数と互換性のある、カスタム関数を作成する](make-custom-functions-compatible-with-xll-udf.md)」で方法を確認してください。</span><span class="sxs-lookup"><span data-stu-id="8a1f5-113">Learn how to [make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="8a1f5-114">関連項目</span><span class="sxs-lookup"><span data-stu-id="8a1f5-114">See also</span></span>

* [<span data-ttu-id="8a1f5-115">カスタム関数の認証</span><span class="sxs-lookup"><span data-stu-id="8a1f5-115">Custom functions authentication</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="8a1f5-116">カスタム関数でデータを受信して​​処理する</span><span class="sxs-lookup"><span data-stu-id="8a1f5-116">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="8a1f5-117">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="8a1f5-117">Create custom functions in Excel</span></span>](custom-functions-overview.md)
