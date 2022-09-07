---
title: インターネット ヘッダーの取得と設定
description: Outlook アドインでメッセージのインターネット ヘッダーを取得および設定する方法。
ms.date: 08/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f8e4af70b24a96b8d00acc7ea4101acf53e2b71
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616029"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>Outlook アドインでメッセージのインターネット ヘッダーを取得および設定する

## <a name="background"></a>背景

Outlook アドイン開発の一般的な要件は、アドインに関連付けられているカスタム プロパティをさまざまなレベルで格納することです。 現時点では、カスタム プロパティはアイテムまたはメールボックス レベルで格納されます。

- アイテム レベル - 特定のアイテムに適用されるプロパティの場合は、 [CustomProperties オブジェクトを](/javascript/api/outlook/office.customproperties) 使用します。 たとえば、電子メールを送信したユーザーに関連付けられている顧客コードを格納します。
- メールボックス レベル - ユーザーのメールボックス内のすべてのメール アイテムに適用されるプロパティについては、 [RoamingSettings オブジェクトを](/javascript/api/outlook/office.roamingsettings) 使用します。 たとえば、ユーザーの好みを保存して、特定のスケールの温度を表示します。

両方の種類のプロパティは、アイテムが Exchange サーバーを離れた後は保持されないため、電子メール受信者はアイテムに設定されたプロパティを取得できません。 そのため、開発者はこれらの設定やその他の多目的インターネット メール拡張機能 (MIME) プロパティにアクセスして、読み取りシナリオを改善することはできません。

Exchange Web Services (EWS) 要求を使用してインターネット ヘッダーを設定する方法がありますが、一部のシナリオでは EWS 要求を行うことはできません。 たとえば、Outlook デスクトップの作成モードでは、アイテム ID はキャッシュ モードでは同期 `saveAsync` されません。

> [!TIP]
> これらのオプションの使用の詳細については、「 [Outlook アドインのアドイン メタデータを取得して設定する](metadata-for-an-outlook-add-in.md)」を参照してください。

## <a name="purpose-of-the-internet-headers-api"></a>インターネット ヘッダー API の目的

[要件セット 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) で導入されたインターネット ヘッダー API を使用すると、開発者は次の機能を利用できます。

- すべてのクライアント間で Exchange を離れた後に保持される電子メールに関する情報をスタンプします。
- メール読み取りシナリオですべてのクライアントで Exchange を離れた後に保持された電子メールに関する情報を読み取ります。
- メールの MIME ヘッダー全体にアクセスします。

![インターネット ヘッダーの図。 テキスト: ユーザー 1 は電子メールを送信します。 アドインは、ユーザーが電子メールを作成している間に、カスタム インターネット ヘッダーを管理します。 ユーザー 2 は電子メールを受信します。 アドインは、受信した電子メールからインターネット ヘッダーを取得し、カスタム ヘッダーを解析して使用します。](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>メッセージの作成中にインターネット ヘッダーを設定する

[item.internetHeaders](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-internetheaders-member) プロパティを使用して、現在のメッセージに配置するカスタム インターネット ヘッダーを作成モードで管理します。

### <a name="set-get-and-remove-custom-internet-headers-example"></a>カスタム インターネット ヘッダーの設定、取得、および削除の例

次の例は、カスタム インターネット ヘッダーを設定、取得、削除する方法を示しています。

```js
// Set custom internet headers.
function setCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.setAsync(
    { "x-preferred-fruit": "orange", "x-preferred-vegetable": "broccoli", "x-best-vegetable": "spinach" },
    setCallback
  );
}

function setCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully set headers");
  } else {
    console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
  }
}

// Get custom internet headers.
function getSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.getAsync(
    ["x-preferred-fruit", "x-preferred-vegetable", "x-best-vegetable", "x-nonexistent-header"],
    getCallback
  );
}

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Selected headers: " + JSON.stringify(asyncResult.value));
  } else {
    console.log("Error getting selected headers: " + JSON.stringify(asyncResult.error));
  }
}

// Remove custom internet headers.
function removeSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.removeAsync(
    ["x-best-vegetable", "x-nonexistent-header"],
    removeCallback);
}

function removeCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully removed selected headers");
  } else {
    console.log("Error removing selected headers: " + JSON.stringify(asyncResult.error));
  }
}

setCustomHeaders();
getSelectedCustomHeaders();
removeSelectedCustomHeaders();
getSelectedCustomHeaders();

/* Sample output:
Successfully set headers
Selected headers: {"x-best-vegetable":"spinach","x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
Successfully removed selected headers
Selected headers: {"x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
*/
```

## <a name="get-internet-headers-while-reading-a-message"></a>メッセージの読み取り中にインターネット ヘッダーを取得する

[item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getallinternetheadersasync-member(1)) を呼び出して、読み取りモードで現在のメッセージのインターネット ヘッダーを取得します。

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>現在の MIME ヘッダーから送信者の設定を取得する例

前のセクションの例を基に、次のコードは、現在の電子メールの MIME ヘッダーから送信者の設定を取得する方法を示しています。

```js
Office.context.mailbox.item.getAllInternetHeadersAsync(getCallback);

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Sender's preferred fruit: " + asyncResult.value.match(/x-preferred-fruit:.*/gim)[0].slice(19));
    console.log("Sender's preferred vegetable: " + asyncResult.value.match(/x-preferred-vegetable:.*/gim)[0].slice(23));
  } else {
    console.log("Error getting preferences from header: " + JSON.stringify(asyncResult.error));
  }
}

/* Sample output:
Sender's preferred fruit: orange
Sender's preferred vegetable: broccoli
*/
```

> [!IMPORTANT]
> このサンプルは、単純なケースに適しています。 より複雑な情報の取得 (RFC [2822](https://tools.ietf.org/html/rfc2822) で説明されている複数インスタンス ヘッダーや折りたたまれた値など) については、適切な MIME 解析ライブラリを使用してみてください。

## <a name="recommended-practices"></a>推奨プラクティス

現在、インターネット ヘッダーは、ユーザーのメールボックス上の有限のリソースです。 クォータが使い果たされると、そのメールボックスにインターネット ヘッダーをこれ以上作成できないため、これに依存するクライアントが予期しない動作を実行する可能性があります。

アドインでインターネット ヘッダーを作成するときは、次のガイドラインを適用します。

- 必要なヘッダーの最小数を作成します。 ヘッダー クォータは、メッセージに適用されるヘッダーの合計サイズに基づきます。 Exchange Onlineでは、ヘッダーの制限は 256 KB に制限されますが、Exchange オンプレミス環境では、組織の管理者によって制限が決定されます。 ヘッダーの制限の詳細については、「[Exchange Online メッセージの制限とメッセージの制限](/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits)[のExchange Server](/exchange/mail-flow/message-size-limits)」を参照してください。
- 後で値を再利用および更新できるように、ヘッダーに名前を付けます。 そのため、ヘッダーに変数の名前を付けないでください (たとえば、ユーザー入力、タイムスタンプなどに基づいて)。

## <a name="see-also"></a>関連項目

- [Outlook アドインのアドイン メタデータを取得および設定する](metadata-for-an-outlook-add-in.md)
