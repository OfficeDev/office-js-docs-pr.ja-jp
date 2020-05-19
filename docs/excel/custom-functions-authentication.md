---
ms.date: 05/17/2020
description: 作業ウィンドウを使用しない Excel でカスタム関数を使用してユーザーを認証します。
title: UI レスのカスタム関数の認証
localization_priority: Normal
ms.openlocfilehash: 93073fb23f3f4d30c36faf4927a3aebdafbc887d
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278379"
---
# <a name="authentication-for-ui-less-custom-functions"></a>UI レスのカスタム関数の認証

一部のシナリオでは、作業ウィンドウやその他のユーザーインターフェイス要素を使用しないカスタム関数 (UI レスカスタム関数) は、保護されたリソースにアクセスするためにユーザーを認証する必要があります。 UI を使用しないカスタム関数は、JavaScript のみのランタイムで実行されることに注意してください。 そのため、JavaScript のみのランタイムと、オブジェクトとダイアログ API を使用してほとんどのアドインで使用される一般的なブラウザーエンジンランタイムとの間でデータをやり取りする必要があり `OfficeRuntime.storage` ます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>OfficeRuntime.storage オブジェクト

UI に含まれないカスタム関数で使用される JavaScript 専用のランタイムには、 `localStorage` 通常、データを格納するグローバルウィンドウで使用できるオブジェクトがありません。 その代わりに、" [Officeruntime](/javascript/api/office-runtime/officeruntime.storage) " を使用して UI レスのカスタム関数と作業ウィンドウ間でデータを共有する必要があります。データを設定および取得するためのストレージです。

### <a name="suggested-usage"></a>おすすめの使用法

UI を使用しないカスタム関数から認証する必要がある場合は、 `storage` アクセストークンが既に取得されているかどうかを確認します。 取得されていない場合は、ダイアログ API を使用してユーザーを認証し、アクセストークンを取得して、後で使用するために `storage` に保存します。

## <a name="dialog-api"></a>ダイアログ API

トークンが存在しない場合は、ユーザーにサインインを求めるダイアログ API を表示する必要があります。 ユーザーが資格情報を入力すると、結果のアクセストークンが `storage` に保存されます。

> [!NOTE]
> JavaScript のみのランタイムは、作業ウィンドウで使用されるブラウザーエンジンランタイムの Dialog オブジェクトとは少し異なるダイアログオブジェクトを使用します。 これらはどちらも "Dialog API" と呼ばれていますが、 `OfficeRuntime.Dialog` JavaScript のみのランタイムでユーザーを認証するために使用します。

この基本的な手順を次の図に示します。 点線は、UI を使用しないカスタム関数とアドインの作業ウィンドウがどちらもアドインの一部であることを示していますが、個別のランタイムを使用しています。

1. Excel ブックのセルから UI を使用しないカスタム関数呼び出しを発行します。
2. UI を使用しないカスタム関数は、 `Dialog` ユーザーの資格情報を web サイトに渡すために使用します。
3. この web サイトは、UI なしのカスタム関数へのアクセストークンを返します。
4. UI を使用しないカスタム関数は、このアクセストークンをに設定し `storage` ます。
5. アドインの作業ウィンドウは、`storage` からトークンにアクセスします。

![アクセストークンを取得するためにダイアログ API を使用したカスタム関数の図。次に、この方法で、作業ウィンドウを使用してトークンを保存します。ストレージ API を使用します。](../images/authentication-diagram.png "認証の図。")

## <a name="storing-the-token"></a>トークンの格納

次の例は、[カスタム関数の OfficeRuntime.storage を使用 ](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)したコードサンプルです。 UI を使用しないカスタム関数と作業ウィンドウ間でデータを共有する完全な例については、以下のコードサンプルを参照してください。

UI を省略したカスタム関数が認証された場合は、アクセストークンを受け取り、それをに格納する必要があり `storage` ます。 次のコードサンプルは、`storage.setItem`メソッドを呼び出して値を格納する方法を示します。 この `storeValue` 関数は、ユーザーの値を格納するなど、UI を使用しないカスタム関数です。 必要なトークン値を格納するように変更できます。

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

Office アドインは web ベースで、あらゆる web 認証技術を使用できます。 UI を使用しないカスタム関数を使用して独自の認証を実装するために従う必要のある特定のパターンやメソッドはありません。 さまざまな認証パターンに関するドキュメントを参照してください。 [この記事では、外部サービスによる認証について説明します。](../develop/auth-external-add-ins.md)  

カスタム関数を開発するときに、次の場所にデータを格納しないようにします。  

- `localStorage`: UI を持たないカスタム関数は、グローバルオブジェクトにアクセスできない `window` ため、に格納されているデータにはアクセスできません `localStorage` 。
- `Office.context.document.settings`: この場所は安全ではないため、アドインを使用しているユーザーが情報を抽出できます。

## <a name="dialog-box-api-example"></a>ダイアログボックス API の例

次のコードサンプルでは、関数は `getTokenViaDialog` API の関数を使用して `Dialog` `displayWebDialogOptions` ダイアログボックスを表示します。 このサンプルは、オブジェクトの機能を示すために提供されており `Dialog` 、認証方法を示すものではありません。

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
[UI のないカスタム関数をデバッグ](custom-functions-debugging.md)する方法について説明します。

## <a name="see-also"></a>関連項目

* [UI レス Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [Excel カスタム関数のチュートリアル](excel-tutorial-custom-functions.md)
