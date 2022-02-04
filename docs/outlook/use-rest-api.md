---
title: Outlook アドインからの Outlook REST API の使用
description: Outlook アドインから Outlook REST API を使用して、アクセス トークンを取得する方法について説明します。
ms.date: 07/06/2021
ms.localizationpriority: medium
---

# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a>Outlook アドインからの Outlook REST API の使用

[Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) 名前空間は、メッセージや予定の多くの共通フィールドへのアクセスを提供します。ただし、シナリオによっては、名前空間によって公開されないデータにアドインがアクセスする必要が生じる可能性があります。たとえば、アドインは外部アプリによって設定されるカスタム プロパティを使用する場合があります。あるいは、同じ送信者からのメッセージをユーザーのメールボックスから検索する必要があります。これらのシナリオでは、[Outlook REST API](/outlook/rest) を使用して情報を取得する方法が推奨されています。

> [!IMPORTANT]
> **REST api Outlookは非推奨です**
>
> REST Outlookは 2022 年 11 月に完全に使用停止されます (詳細については、[2020](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/) 年 11 月の発表を参照してください)。 Microsoft アドインを使用するには、既存のアドインを[移行Graph。](/outlook/rest#outlook-rest-api-via-microsoft-graph) また、[REST API エンドポイントGraphとOutlook比較します](/outlook/rest/compare-graph)。

## <a name="get-an-access-token"></a>アクセス トークンを取得する

Outlook REST API では、`Authorization` ヘッダーにベアラー トークンが必要です。通常、アプリは OAuth2 フローを使用してトークンを取得します。ただし、アドインは、メールボックス要件セット 1.5 で導入されている新しい [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用することにより、OAuth2 を実装せずにトークンを取得できます。

`isRest` オプションを `true` に設定することにより、REST API と互換性があるトークンを要求できます。

### <a name="add-in-permissions-and-token-scope"></a>アドインのアクセス許可とトークンの範囲

REST API を経由してアドインが必要とするアクセスのレベルを考慮することが重要です。ほとんどの場合、`getCallbackTokenAsync` によって返されるトークンは、現在の項目への読み取り専用のアクセスのみを提供します。このことは、アドインがそのマニフェストに `ReadWriteItem` アクセス許可レベルを指定する場合にも当てはまります。

現在の項目またはユーザーのメールボックス内のその他の項目への書き込みアクセスがアドインに必要な場合、アドインがそのマニフェストに `ReadWriteMailbox` アクセス許可レベルを指定する必要があります。その場合、返されるトークンに、ユーザーのメッセージ、イベント、および連絡先に対する読み取り/書き込みアクセス権限が含まれます。

### <a name="example"></a>例

```js
Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    var accessToken = result.value;

    // Use the access token.
    getCurrentItem(accessToken);
  } else {
    // Handle the error.
  }
});
```

## <a name="get-the-item-id"></a>項目 ID を取得する

REST を経由して現在の項目を取得するには、REST 用に正しく書式設定された項目の ID がアドインに必要です。 これは [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティから取得されますが、REST 用に書式設定された ID であることを確認するためのいくつかの確認が必要です。

- Outlook Mobile の場合、`Office.context.mailbox.item.itemId` によって返される値が REST 用に形式設定された ID であり、そのまま使用できます。
- その他の Outlook クライアントの場合、`Office.context.mailbox.item.itemId` によって返される値が EWS 用に設定された ID であり、[Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用して変換する必要があります。
- また、これを使用するには、Attachment ID を REST 用に形式設定された ID に変換する必要もあります。 ID を変換する必要がある理由は、EWS ID に URL セーフ以外の値が含まれている可能性があり、その場合は REST で問題が発生するためです。

[Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostname-member) プロパティを確認することにより、アドインは読み込まれる Outlook クライアントを判別できます。

### <a name="example"></a>例

```js
function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}
```

## <a name="get-the-rest-api-url"></a>REST API URL を取得する

REST API を呼び出すためにアドインで必要な情報の最終部分は、API 要求の送信に使用するホスト名です。この情報は [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) プロパティにあります。

### <a name="example"></a>例

```js
// Example: https://outlook.office.com
var restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a>API を呼び出す

アドインがアクセス トークン、アイテム ID、および REST API URL を取得すると、REST API を呼び出すバックエンド サービスにその情報を渡すか、AJAX を使用して直接呼び出すことができるようになります。 次の例は、Outlook Mail REST API を呼び出して現在のメッセージを取得します。

> [!IMPORTANT]
> オンプレミスのExchange展開では、AJAX または類似のライブラリを使用するクライアント側の要求は、そのサーバーセットアップで CORS がサポートされていないため失敗します。

```js
function getCurrentItem(accessToken) {
  // Get the item's REST ID.
  var itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  var getMessageUrl = Office.context.mailbox.restUrl +
    '/v2.0/me/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`.
    var subject = item.Subject;
    ...
  }).fail(function(error){
    // Handle error.
  });
}
```

## <a name="see-also"></a>関連項目

- Outlook アドインから REST API を呼び出す例については、GitHub の [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) をご覧ください。
- Outlook REST API は、Microsoft Graph エンドポイントからでも利用できますが、アドインでアクセス トークンを取得する方法など、重要な違いがいくつかあります。 詳細については、「[Microsoft Graph を介して使用する Outlook REST API](/outlook/rest/index#outlook-rest-api-via-microsoft-graph)」を参照してください。