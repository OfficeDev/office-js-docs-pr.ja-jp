---
title: インターネットヘッダーを取得および設定する
description: Outlook アドインでメッセージのインターネットヘッダーを取得および設定する方法について説明します。
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: a05ba86eebd8dc01c8368b61e39d1de1d90f9efa
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609084"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>Outlook アドインでメッセージのインターネットヘッダーを取得および設定する

## <a name="background"></a>背景

Outlook アドインの開発での一般的な要件は、アドインに関連付けられたカスタムプロパティをさまざまなレベルで保存することです。 現在、カスタムプロパティは、アイテムまたはメールボックスのレベルで保存されています。

- アイテムレベル-特定のアイテムに適用されるプロパティについては、 [CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトを使用します。 たとえば、電子メールの送信者に関連付けられている顧客コードを格納します。
- メールボックスレベル-ユーザーのメールボックス内のすべてのメールアイテムに適用されるプロパティについては、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)オブジェクトを使用します。 たとえば、ユーザーの設定を保存して、特定の規模で温度を表示します。

アイテムが Exchange サーバーから離脱した後、両方の種類のプロパティは保持されないため、電子メールの受信者は、アイテムに設定されているプロパティを取得できません。 そのため、開発者はこれらの設定やその他の MIME プロパティにアクセスして、読み取りシナリオを改善することはできません。

EWS 要求を通じてインターネットヘッダーを設定する方法はありますが、一部のシナリオでは EWS 要求が機能しない場合があります。 たとえば、Outlook デスクトップの新規作成モードでは、アイテム id は  `saveAsync` キャッシュモードで同期されません   。

> [!TIP]
> これらのオプションの使用の詳細については[、「Outlook アドインのアドインメタデータを取得および設定](metadata-for-an-outlook-add-in.md)する」を参照してください。

## <a name="purpose-of-the-internet-headers-api"></a>インターネットヘッダー API の目的

[要件セット 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)で導入されたインターネットヘッダー api を使用すると、開発者は次のことを行うことができます。

- すべてのクライアント間で Exchange を残した後に保持されるメールについての情報をスタンプします。
- メールの読み取りシナリオにおいて、すべてのクライアント間で Exchange のメールが残された後に保持される電子メールの情報を読み取ります。
- 電子メールの MIME ヘッダー全体にアクセスします。

![インターネットヘッダーの図 テキスト: ユーザー1が電子メールを送信します。 アドインは、ユーザーが電子メールを作成しているときに、カスタムのインターネットヘッダーを管理します。 ユーザー2が電子メールを受信します。 アドインは受信した電子メールからインターネットヘッダーを取得し、カスタムヘッダーを解析して使用します。](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>メッセージの作成中にインターネットヘッダーを設定する

[InternetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders)プロパティを使用して、新規作成モードで現在のメッセージに配置するカスタムインターネットヘッダーを管理します。

### <a name="set-get-and-remove-custom-headers-example"></a>カスタムヘッダーの設定、取得、および削除の例

次の例は、カスタムヘッダーを設定、取得、および削除する方法を示しています。

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

## <a name="get-internet-headers-while-reading-a-message"></a>メッセージの読み取り中にインターネットヘッダーを取得する

現在のメッセージのインターネットヘッダーを閲覧モードで取得するには、 [getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-)を呼び出してください。

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>現在の MIME ヘッダーの送信者の設定を取得する例

前のセクションの例では、次のコードは、現在の電子メールの MIME ヘッダーから送信者の設定を取得する方法を示しています。

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
> このサンプルは、単純なケースで機能します。 複雑な情報取得 ( [RFC 2822](https://tools.ietf.org/html/rfc2822)で説明されているように、複数インスタンスのヘッダー、または折りたたまれた値など) を取得するには、適切な MIME 解析ライブラリを使用してみてください。

## <a name="recommended-practices"></a>推奨プラクティス

現時点では、インターネットヘッダーはユーザーのメールボックス上の有限リソースです。 クォータが不足している場合は、そのメールボックスにより多くのインターネットヘッダーを作成することはできません。これにより、これに依存するクライアントから予期しない動作が発生する可能性があります。

アドインでインターネットヘッダーを作成するときには、次のガイドラインを適用します。

- 必要なヘッダーの最小数を作成します。
- 後で再利用して値を更新できるように、名前のヘッダー。 そのため、ユーザーの入力やタイムスタンプなどに基づいて、変数の方法でヘッダーに名前を付けることは避けてください。

## <a name="see-also"></a>関連項目

- [Outlook アドインのアドイン メタデータを取得および設定する](metadata-for-an-outlook-add-in.md)
