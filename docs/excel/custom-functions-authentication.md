---
ms.date: 04/15/2019
description: Excel でカスタム関数を使用してユーザーを認証します。
title: カスタム関数の認証
ms.openlocfilehash: 75ffb82c0dc9350c35b22b1d1676990598ea0c44
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914187"
---
# <a name="authentication"></a>認証

一部のシナリオでは、カスタム関数は、保護されたリソースにアクセスするためにユーザーを認証する必要があります。 カスタム関数では、特定の認証方法を使用する必要はありませんが、カスタム関数は、アドインの作業ウィンドウや他の UI 要素とは別のランタイムで実行されることに注意してください。 そのため、 `AsyncStorage`オブジェクトとダイアログ API を使用して、2つのランタイム間でデータをやり取りする必要があります。
  
## <a name="asyncstorage-object"></a>asyncstorage オブジェクト

カスタム関数ランタイムには、通常`localStorage` 、データを格納するグローバルウィンドウで使用できるオブジェクトがありません。 代わりに、データを設定して取得するために、 [officeruntime](/javascript/api/office-runtime/officeruntime.asyncstorage)を使用して、カスタム関数と作業ウィンドウ間でデータを共有する必要があります。

さらに、を使用`AsyncStorage`するメリットがあります。セキュリティで保護されたサンドボックス環境を使用して、他のアドインがデータにアクセスできないようにします。

### <a name="suggested-usage"></a>推奨される使用法

作業ウィンドウまたはカスタム関数から認証を受ける必要がある場合は、 `AsyncStorage`アクセストークンが既に取得されているかどうかを確認します。 表示されない場合は、ダイアログ API を使用してユーザーを認証し、アクセストークンを取得し`AsyncStorage`てから、トークンを後で使用するために格納します。

## <a name="dialog-api"></a>ダイアログ API

トークンが存在しない場合は、ダイアログ API を使用して、ユーザーにサインインを要求する必要があります。 ユーザーが資格情報を入力すると、作成されたアクセストークン`AsyncStorage`がに保存されます。

> [!NOTE]
> カスタム関数ランタイムは、作業ウィンドウで使用されるブラウザーエンジンランタイムの dialog オブジェクトとは少し異なるダイアログオブジェクトを使用します。 これらはどちらも "Dialog API" と呼ばれています`Officeruntime.Dialog`が、カスタム関数ランタイムでユーザーを認証するために使用します。

の`OfficeRuntime.Dialog`使用方法については、「 [Custom Functions dialog](/office/dev/add-ins/excel/custom-functions-dialog)」を参照してください。

全体として認証プロセス全体を構想する場合は、アドインの作業ウィンドウと UI 要素、およびアドインのカスタム関数の部分を、を通じて`AsyncStorage`相互に通信できる個別のエンティティと考えることをお勧めします。

次の図は、この基本的なプロセスの概要を示しています。 点線では、個別の操作を実行する一方で、カスタム関数とアドインの作業ウィンドウは、どちらもアドインの一部であることに注意してください。

1. Excel ブックのセルからカスタム関数呼び出しを発行します。
2. カスタム関数を使用`Officeruntime.Dialog`して、ユーザーの資格情報を web サイトに渡します。
3. その後、この web サイトは、カスタム関数へのアクセストークンを返します。
4. 次に、カスタム関数は、 `AsyncStorage`このアクセストークンをに設定します。
5. アドインの作業ウィンドウで、から`AsyncStorage`トークンにアクセスできます。

![ダイアログ API を使用してアクセストークンを取得し、asyncstorage API を使用してトークンを作業ウィンドウで共有するカスタム関数の図。](../images/authentication-diagram.png "認証の図。")

## <a name="storing-the-token"></a>トークンの保存

次の例は、 [「カスタム関数の asyncstorage を使用する」](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)のコードサンプルのものです。 カスタム関数と作業ウィンドウとの間でデータを共有する完全な例については、次のコードサンプルを参照してください。

カスタム関数が認証された場合は、アクセストークンを受け取り、それをに`AsyncStorage`格納する必要があります。 次のコードサンプルは、メソッドを呼び出し`AsyncStorage.setItem`て値を格納する方法を示しています。 この`StoreValue`関数は、たとえば、ユーザーの値を格納するためのカスタム関数です。 必要なトークン値を格納するように変更することができます。

```javascript
function StoreValue(key, value) {
  return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}
```

作業ウィンドウでアクセストークンが必要になると、そのトークンを取得`AsyncStorage`することができます。 次のコードサンプルは、 `AsyncStorage.getItem`メソッドを使用してトークンを取得する方法を示しています。

```javascript
function ReceiveTokenFromCustomFunction() {
   var key = "token";
   var tokenSendStatus = document.getElementById('tokenSendStatus');
   OfficeRuntime.AsyncStorage.getItem(key).then(function (result) {
      tokenSendStatus.value = "Success: Item with key '" + key + "' read from AsyncStorage.";
      document.getElementById('tokenTextBox2').value = result;
   }, function (error) {
      tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from AsyncStorage. " + error;
   });
}
```

## <a name="general-guidance"></a>一般的なガイダンス

Office アドインは web ベースであり、任意の web 認証方法を使用できます。 カスタム関数を使用して独自の認証を実装するために従う必要のある特定のパターンやメソッドはありません。 [外部サービスによる承認については、この記事](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js)から始まるさまざまな認証パターンに関するドキュメントを参照してください。  

カスタム関数を開発する際に、次の場所を使用してデータを保存しないようにします。  

- `localStorage`: カスタム関数にはグローバル`window`オブジェクトへのアクセス権がないため、に`localStorage`格納されているデータにアクセスできません。
- `Office.context.document.settings`: この場所はセキュリティで保護されていないため、アドインを使用するすべてのユーザーが情報を抽出できます。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [Excel カスタム関数のチュートリアル](excel-tutorial-custom-functions.md)
