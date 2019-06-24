---
ms.date: 06/17/2019
description: Excel のカスタム関数を使用してユーザーを認証します。
title: カスタム関数の認証
localization_priority: Priority
ms.openlocfilehash: 91755a76751406e87eb8a1f316e4b163ada98b45
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127983"
---
# <a name="authentication-for-custom-functions"></a>カスタム関数の認証

一部のシナリオでは、保護されたリソースにアクセスするために、ユーザーを認証する必要があります。 カスタム関数は特定の認証方法を必要としませんが、カスタム関数は、アドインの作業ウィンドウと他の UI 要素とは別のランタイムで実行されます。 このため、`OfficeRuntime.storage` オブジェクトとダイアログ API を使用して 2 つのランタイム間でデータを受け渡しする必要があります。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="officeruntimestorage-object"></a>OfficeRuntime.storage オブジェクト

カスタム関数ランタイムには、通常はデータを格納するグローバルウィンドウに使用できる `localStorage` オブジェクトがありません。 代わりに、データを設定して取得するためのストレージ[OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage)を使用して、カスタム関数と作業ウィンドウの間でデータを共有する必要があります。

また、`storage`オブジェクトを使用すると便利です。セキュリティサンドボックス環境を使用するため、他のアドインがデータにアクセスすることができません。

### <a name="suggested-usage"></a>おすすめの使用法

作業ウィンドウまたはカスタム関数から認証する必要がある場合は、アクセストークンが既に取得されているかどうか `storage` を確認します。 取得されていない場合は、ダイアログ API を使用してユーザーを認証し、アクセストークンを取得して、後で使用するために `storage` に保存します。

## <a name="dialog-api"></a>ダイアログ API

トークンが存在しない場合は、ユーザーにサインインを求めるダイアログ API を表示する必要があります。 ユーザーが資格情報を入力すると、結果のアクセストークンが `storage` に保存されます。

> [!NOTE]
> カスタム関数のランタイムは、作業ウィンドウで使用されるブラウザー エンジン ランタイムのダイアログ オブジェクトとは少し異なるダイアログ オブジェクトを使用します。 いずれも "ダイアログ API" と呼ばれていますが、カスタム関数のランタイムでユーザーを認証するために `OfficeRuntime.Dialog` を使用します。

`Dialog` オブジェクトを使用する方法の詳細については、「[カスタム関数ダイアログ](/office/dev/add-ins/excel/custom-functions-dialog)」を参照してください。

認証プロセス全体を構想するときには、アドインの作業ウィンドウと UI 要素、アドインのカスタム関数部分が、`OfficeRuntime.storage` を通じて相互に通信できる個別のエンティティとして考えてみることをおすすめします。

この基本的な手順を次の図に示します。 点線は、ユーザーが個別の操作を実行している間に、カスタム関数とアドインの作業ウィンドウがアドインの一部であることを示しています。

1. Excel ワークブックのセルからカスタム関数を発行します。
2. カスタム関数は、ユーザーの資格情報を web サイトに渡すために `Dialog` を使用します。
3. その後、web サイトは、アクセストークンをカスタム関数に返します。
4. このアクセストークンは、カスタム関数によって `storage` に設定されます。
5. アドインの作業ウィンドウは、`storage` からトークンにアクセスします。

![アクセス トークンを取得するためのダイアログ API を使用したカスタム関数の図と、OfficeRuntime.storage API を通してトークンを作業ウィンドウと共有します。](../images/authentication-diagram.png "認証図。")

## <a name="storing-the-token"></a>トークンの格納

次の例は、[カスタム関数の OfficeRuntime.storage を使用 ](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)したコードサンプルです。 カスタム関数と作業ウィンドウ間のデータ共有の例については、このコードサンプルを参照してください。

カスタム関数が認証されたら、アクセストークンを受け取り、`storage`に保存する必要があります。 次のコードサンプルは、`storage.setItem`メソッドを呼び出して値を格納する方法を示します。 `storeValue` 関数は、ユーザーからの値を格納するためのカスタム関数です。 必要なトークン値を格納するように変更できます。

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

CustomFunctions.associate("STOREVALUE", storeValue);
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
CustomFunctions.associate("GETTOKEN", receiveTokenFromCustomFunction);

```

## <a name="general-guidance"></a>一般的なガイダンス

Office アドインは web ベースで、あらゆる web 認証技術を使用できます。 カスタム関数を使用して独自の認証を実装するのに、特定のパターンやメソッドはありません。 さまざまな認証パターンに関するドキュメントを参照してください。 [この記事では、外部サービスによる認証について説明します。](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js)  

カスタム関数を開発するときに、次の場所にデータを格納しないようにします。  

- `localStorage`: カスタム関数はグローバル `window` オブジェクトへのアクセス権がないため、`localStorage` に保存されているデータにはアクセスできません。
- `Office.context.document.settings`: この場所は安全ではないため、アドインを使用しているユーザーが情報を抽出できます。

## <a name="next-steps"></a>次の手順
[カスタム関数のダイアログ API ](custom-functions-dialog.md) について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数のアーキテクチャ](custom-functions-architecture.md)
* [カスタム関数でデータを受信して​​処理する](custom-functions-web-reqs.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [Excel カスタム関数のチュートリアル](excel-tutorial-custom-functions.md)
