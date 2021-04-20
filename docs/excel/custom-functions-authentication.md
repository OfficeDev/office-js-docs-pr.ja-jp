---
ms.date: 05/17/2020
description: 作業ウィンドウを使用しない Excel のカスタム関数を使用してユーザーを認証します。
title: UI を使用するカスタム関数の認証
localization_priority: Normal
ms.openlocfilehash: bca3cd422330b6499e18c31ef8d7da6def81b546
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839860"
---
# <a name="authentication-for-ui-less-custom-functions"></a>UI を使用するカスタム関数の認証

シナリオによっては、作業ウィンドウや他のユーザー インターフェイス要素 (UI を使用しないカスタム関数) を使用しないカスタム関数は、保護されたリソースにアクセスするためにユーザーを認証する必要があります。 UI を使用するカスタム関数は、JavaScript 専用ランタイムで実行されます。 この理由から、JavaScript 専用ランタイムと、オブジェクトとダイアログ API を使用するほとんどのアドインで使用される一般的なブラウザー エンジン ランタイムとの間でデータを受け渡す必要があります。 `OfficeRuntime.storage`

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>OfficeRuntime.storage オブジェクト

UI を使用しないカスタム関数で使用される JavaScript 専用のランタイムには、通常はデータを格納するグローバル ウィンドウで使用できる `localStorage` オブジェクトが用意されません。 代わりに [、OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) を使用してデータを設定および取得することで、UI を使用するカスタム関数と作業ウィンドウの間でデータを共有する必要があります。

### <a name="suggested-usage"></a>おすすめの使用法

UI を使用していないカスタム関数から認証する必要がある場合は、アクセス トークンが既に取得 `storage` されたのか確認します。 取得されていない場合は、ダイアログ API を使用してユーザーを認証し、アクセストークンを取得して、後で使用するために `storage` に保存します。

## <a name="dialog-api"></a>ダイアログ API

トークンが存在しない場合は、ユーザーにサインインを求めるダイアログ API を表示する必要があります。 ユーザーが資格情報を入力すると、結果のアクセストークンが `storage` に保存されます。

> [!NOTE]
> JavaScript 専用ランタイムは、作業ウィンドウで使用されるブラウザー エンジン ランタイムの Dialog オブジェクトとは少し異なる Dialog オブジェクトを使用します。 これらはどちらも "ダイアログ API" と呼ばれますが、JavaScript 専用ランタイムでユーザーを認証 `OfficeRuntime.Dialog` するために使用されます。

この基本的な手順を次の図に示します。 点線は、UI を使用するカスタム関数とアドインの作業ウィンドウの両方がアドイン全体の一部であり、別々のランタイムを使用する点を示しています。

1. Excel ブックのセルから UI を使用するカスタム関数呼び出しを発行します。
2. UI を使用するカスタム関数は、 `Dialog` ユーザー資格情報を Web サイトに渡す場合に使用します。
3. 次に、この Web サイトは、UI を使用するカスタム関数にアクセス トークンを返します。
4. その後、UI を使用するカスタム関数は、このアクセス トークンを次のアクセス トークンに設定します `storage` 。
5. アドインの作業ウィンドウは、`storage` からトークンにアクセスします。

![ダイアログ API を使用してアクセス トークンを取得し、OfficeRuntime.storage API を使用して作業ウィンドウとトークンを共有するカスタム関数の図。](../images/authentication-diagram.png "認証の図。")

## <a name="storing-the-token"></a>トークンの格納

次の例は、[カスタム関数の OfficeRuntime.storage を使用 ](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)したコードサンプルです。 UI を使用するカスタム関数と作業ウィンドウの間でデータを共有する完全な例については、このコード サンプルを参照してください。

UI を使用するカスタム関数が認証を受ける場合は、アクセス トークンを受け取り、それを格納する必要があります `storage` 。 次のコードサンプルは、`storage.setItem`メソッドを呼び出して値を格納する方法を示します。 この関数は UI を使用するカスタム関数で、たとえばユーザーの値 `storeValue` を格納します。 必要なトークン値を格納するように変更できます。

```js
/**
 * Stores a key-value pair into OfficeRuntime.storage.
 * @customfunction
 * @param {string} key Key of item to put into storage.
 * @param {*} value Value of item to put into storage.
 */
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

作業ウィンドウにアクセストークンが必要な場合は、`storage`から トークンを取得できます。 次のコードサンプルは、`storage.getItem`メソッドを使用してトークンを取得する方法を示します。

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  var key = "token";
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## <a name="general-guidance"></a>一般的なガイダンス

Office アドインは web ベースで、あらゆる web 認証技術を使用できます。 UI を使用するカスタム関数を使用して独自の認証を実装するために従う必要のある特定のパターンやメソッドはありません。 さまざまな認証パターンに関するドキュメントを参照してください。 [この記事では、外部サービスによる認証について説明します。](../develop/auth-external-add-ins.md)  

カスタム関数を開発するときに、次の場所にデータを格納しないようにします。  

- `localStorage`: UI を使用しないカスタム関数は、グローバル オブジェクトにアクセスできないので、格納されている `window` データにアクセスできない `localStorage` 。
- `Office.context.document.settings`: この場所は安全ではないため、アドインを使用しているユーザーが情報を抽出できます。

## <a name="dialog-box-api-example"></a>ダイアログ ボックス API の例

次のコード サンプルでは、この関数は `getTokenViaDialog` `Dialog` API の関数を使用 `displayWebDialogOptions` してダイアログ ボックスを表示します。 このサンプルは、認証方法を示すのではなく、オブジェクトの `Dialog` 機能を示すサンプルです。

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
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
UI を使用する [カスタム関数をデバッグする方法について説明します](custom-functions-debugging.md)。

## <a name="see-also"></a>関連項目

* [UI を使用する Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)