---
title: 共有ランタイムのないカスタム関数の認証
description: 共有ランタイムを使用しないカスタム関数を使用してユーザーを認証します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7ff7b1dca67e9e25f14ef07bd1c088608f254427
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958426"
---
# <a name="authentication-for-custom-functions-without-a-shared-runtime"></a>共有ランタイムのないカスタム関数の認証

一部のシナリオでは、共有ランタイムを使用しないカスタム関数が、保護されたリソースにアクセスするためにユーザーを認証する必要があります。 共有ランタイムを使用しないカスタム関数は、JavaScript 専用ランタイムで実行されます。 このため、アドインに作業ウィンドウがある場合は、JavaScript 専用ランタイムと作業ウィンドウで使用される HTML サポート ランタイムの間でデータをやり取りする必要があります。 これを行うには、 [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) オブジェクトと特別な Dialog API を使用します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>OfficeRuntime.storage オブジェクト

JavaScript 専用ランタイムには、通常データを `localStorage` 格納するグローバル ウィンドウで使用できるオブジェクトがありません。 代わりに、コードは、データの設定と取得に使用 `OfficeRuntime.storage` して、カスタム関数と作業ウィンドウの間でデータを共有する必要があります。

### <a name="suggested-usage"></a>おすすめの使用法

共有ランタイムを使用しないカスタム関数アドインから認証する必要がある場合は、コードでアクセス トークンが既に取得されているかどうかを確認 `OfficeRuntime.storage` する必要があります。 そうでない場合は、 [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) を使用してユーザーを認証し、アクセス トークンを取得し、今後使用できるようにトークン `OfficeRuntime.storage` を格納します。

## <a name="dialog-api"></a>ダイアログ API

トークンが存在しない場合は、API を `OfficeRuntime.dialog` 使用してユーザーにサインインを求める必要があります。 ユーザーが資格情報を入力した後、結果のアクセス トークンをアイテム `OfficeRuntime.storage`として格納できます。

> [!NOTE]
> JavaScript 専用ランタイムは、作業ウィンドウで使用されるブラウザー エンジン ランタイムのダイアログ オブジェクトとは若干異なるダイアログ オブジェクトを使用します。 どちらも "ダイアログ API" と呼ばれますが、[OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) を使用して、[Office.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) *ではなく* JavaScript 専用ランタイムでユーザーを認証します。

この基本的な手順を次の図に示します。 点線は、カスタム関数とアドインの作業ウィンドウがアドイン全体の一部であることを示しますが、個別のランタイムを使用します。

1. Excel ワークブックのセルからカスタム関数を発行します。
2. カスタム関数は、ユーザーの資格情報を web サイトに渡すために `OfficeRuntime.dialog` を使用します。
3. その後、web サイトは、アクセストークンをカスタム関数に返します。
4. 次に、カスタム関数によって、このアクセス トークンが .`OfficeRuntime.storage`
5. アドインの作業ウィンドウは、`OfficeRuntime.storage` からトークンにアクセスします。

![ダイアログ API を使用してアクセス トークンを取得し、OfficeRuntime.storage API を使用して作業ウィンドウでトークンを共有するカスタム関数の図。](../images/authentication-diagram.png "認証図。")

## <a name="storing-the-token"></a>トークンの格納

次の例は、[カスタム関数の OfficeRuntime.storage を使用 ](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AsyncStorage)したコードサンプルです。 カスタム関数と、共有ランタイムを使用しないアドインの作業ウィンドウ間でデータを共有する完全な例については、このコード サンプルを参照してください。

カスタム関数が認証されたら、アクセストークンを受け取り、`OfficeRuntime.storage`に保存する必要があります。 次のコードサンプルは、`storage.setItem`メソッドを呼び出して値を格納する方法を示します。 この `storeValue` 関数は、ユーザーの値を格納するカスタム関数です。 必要なトークン値を格納するように変更できます。

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

作業ウィンドウにアクセス トークンが必要な場合は、アイテムからトークンを `OfficeRuntime.storage` 取得できます。 次のコードサンプルは、`storage.getItem`メソッドを使用してトークンを取得する方法を示します。

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  const key = "token";
  const tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## <a name="general-guidance"></a>一般的なガイダンス

Office アドインは web ベースで、あらゆる web 認証技術を使用できます。 カスタム関数を使用して独自の認証を実装するのに、特定のパターンやメソッドはありません。 さまざまな認証パターンに関するドキュメントを参照してください。 [この記事では、外部サービスによる認証について説明します。](../develop/auth-external-add-ins.md)  

カスタム関数を開発するときに、次の場所にデータを格納しないようにします。

- `localStorage`: 共有ランタイムを使用しないカスタム関数は、グローバル `window` オブジェクトにアクセスできないため、格納されている `localStorage`データにアクセスできません。
- `Office.context.document.settings`: この場所はセキュリティで保護されておらず、アドインを使用して誰でも情報を抽出できます。

## <a name="dialog-box-api-example"></a>ダイアログ ボックス API の例

次のコード サンプルでは、関数は関数 `getTokenViaDialog` を `OfficeRuntime.displayWebDialog` 使用してダイアログ ボックスを表示します。 このサンプルは、認証方法ではなく、メソッドの機能を示すために提供されています。

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this isn't a sufficient example of authentication but is intended to show the capabilities of the displayWebDialog method.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
      let timeout = 5;
      let count = 0;
      const intervalId = setInterval(function () {
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

[カスタム関数をデバッグ](custom-functions-debugging.md)する方法について説明します。

## <a name="see-also"></a>関連項目

- [カスタム関数のための JavaScript 専用ランタイム](custom-functions-runtime.md)
- [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
