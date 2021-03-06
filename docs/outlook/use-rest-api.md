---
title: Outlook アドインからの Outlook REST API の使用
description: Outlook アドインから Outlook REST API を使用して、アクセス トークンを取得する方法について説明します。
ms.date: 02/26/2021
localization_priority: Normal
ms.openlocfilehash: c0df1df4fdbda22768562892874e09bbeb760473
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505487"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a><span data-ttu-id="b63cf-103">Outlook アドインからの Outlook REST API の使用</span><span class="sxs-lookup"><span data-stu-id="b63cf-103">Use the Outlook REST APIs from an Outlook add-in</span></span>

<span data-ttu-id="b63cf-p101">[Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) 名前空間は、メッセージや予定の多くの共通フィールドへのアクセスを提供します。ただし、シナリオによっては、名前空間によって公開されないデータにアドインがアクセスする必要が生じる可能性があります。たとえば、アドインは外部アプリによって設定されるカスタム プロパティを使用する場合があります。あるいは、同じ送信者からのメッセージをユーザーのメールボックスから検索する必要があります。これらのシナリオでは、[Outlook REST API](/outlook/rest) を使用して情報を取得する方法が推奨されています。</span><span class="sxs-lookup"><span data-stu-id="b63cf-p101">The [Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) namespace provides access to many of the common fields of messages and appointments. However, in some scenarios an add-in may need to access data that is not exposed by the namespace. For example, the add-in may rely on custom properties set by an outside app, or it needs to search the user's mailbox for messages from the same sender. In these scenarios, the [Outlook REST APIs](/outlook/rest) is the recommended method to retrieve the information.</span></span>

> [!NOTE]
> <span data-ttu-id="b63cf-108">[Microsoft Graph から Outlook REST API](/outlook/rest#outlook-rest-api-via-microsoft-graph) にアクセスすることもできますが、いくつかの重要な違いがあります。</span><span class="sxs-lookup"><span data-stu-id="b63cf-108">You can also access [Outlook REST APIs via Microsoft Graph](/outlook/rest#outlook-rest-api-via-microsoft-graph) but there are some key differences.</span></span> <span data-ttu-id="b63cf-109">詳細については、「[Microsoft Graph と Outlook の比較](/outlook/rest/compare-graph)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="b63cf-109">For more details, please [Compare Microsoft Graph and Outlook](/outlook/rest/compare-graph).</span></span>

## <a name="get-an-access-token"></a><span data-ttu-id="b63cf-110">アクセス トークンを取得する</span><span class="sxs-lookup"><span data-stu-id="b63cf-110">Get an access token</span></span>

<span data-ttu-id="b63cf-p103">Outlook REST API では、`Authorization` ヘッダーにベアラー トークンが必要です。通常、アプリは OAuth2 フローを使用してトークンを取得します。ただし、アドインは、メールボックス要件セット 1.5 で導入されている新しい [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用することにより、OAuth2 を実装せずにトークンを取得できます。</span><span class="sxs-lookup"><span data-stu-id="b63cf-p103">The Outlook REST APIs require a bearer token in the `Authorization` header. Typically apps use OAuth2 flows to retrieve a token. However, add-ins can retrieve a token without implementing OAuth2 by using the new [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method introduced in the Mailbox requirement set 1.5.</span></span>

<span data-ttu-id="b63cf-114">`isRest` オプションを `true` に設定することにより、REST API と互換性があるトークンを要求できます。</span><span class="sxs-lookup"><span data-stu-id="b63cf-114">By setting the `isRest` option to `true`, you can request a token compatible with the REST APIs.</span></span>

### <a name="add-in-permissions-and-token-scope"></a><span data-ttu-id="b63cf-115">アドインのアクセス許可とトークンの範囲</span><span class="sxs-lookup"><span data-stu-id="b63cf-115">Add-in permissions and token scope</span></span>

<span data-ttu-id="b63cf-p104">REST API を経由してアドインが必要とするアクセスのレベルを考慮することが重要です。ほとんどの場合、`getCallbackTokenAsync` によって返されるトークンは、現在の項目への読み取り専用のアクセスのみを提供します。このことは、アドインがそのマニフェストに `ReadWriteItem` アクセス許可レベルを指定する場合にも当てはまります。</span><span class="sxs-lookup"><span data-stu-id="b63cf-p104">It is important to consider what level of access your add-in will need via the REST APIs. In most cases, the token returned by `getCallbackTokenAsync` will provide read-only access to the current item only. This is true even if your add-in specifies the `ReadWriteItem` permission level in its manifest.</span></span>

<span data-ttu-id="b63cf-p105">現在の項目またはユーザーのメールボックス内のその他の項目への書き込みアクセスがアドインに必要な場合、アドインがそのマニフェストに `ReadWriteMailbox` アクセス許可レベルを指定する必要があります。その場合、返されるトークンに、ユーザーのメッセージ、イベント、および連絡先に対する読み取り/書き込みアクセス権限が含まれます。</span><span class="sxs-lookup"><span data-stu-id="b63cf-p105">If your add-in will require write access to the current item or other items in the user's mailbox, your add-in must specify the `ReadWriteMailbox` permission level in its manifest. In this case, the token returned will contain read/write access to the user's messages, events, and contacts.</span></span>

### <a name="example"></a><span data-ttu-id="b63cf-121">例</span><span class="sxs-lookup"><span data-stu-id="b63cf-121">Example</span></span>

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

## <a name="get-the-item-id"></a><span data-ttu-id="b63cf-122">項目 ID を取得する</span><span class="sxs-lookup"><span data-stu-id="b63cf-122">Get the item ID</span></span>

<span data-ttu-id="b63cf-123">REST を経由して現在の項目を取得するには、REST 用に正しく書式設定された項目の ID がアドインに必要です。</span><span class="sxs-lookup"><span data-stu-id="b63cf-123">To retrieve the current item via REST, your add-in will need the item's ID, properly formatted for REST.</span></span> <span data-ttu-id="b63cf-124">これは [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティから取得されますが、REST 用に書式設定された ID であることを確認するためのいくつかの確認が必要です。</span><span class="sxs-lookup"><span data-stu-id="b63cf-124">This is obtained from the [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property, but some checks should be made to ensure that it is a REST-formatted ID.</span></span>

- <span data-ttu-id="b63cf-125">Outlook Mobile の場合、`Office.context.mailbox.item.itemId` によって返される値が REST 用に形式設定された ID であり、そのまま使用できます。</span><span class="sxs-lookup"><span data-stu-id="b63cf-125">In Outlook Mobile, the value returned by `Office.context.mailbox.item.itemId` is a REST-formatted ID and can be used as-is.</span></span>
- <span data-ttu-id="b63cf-126">その他の Outlook クライアントの場合、`Office.context.mailbox.item.itemId` によって返される値が EWS 用に設定された ID であり、[Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b63cf-126">In other Outlook clients, the value returned by `Office.context.mailbox.item.itemId` is an EWS-formatted ID, and must be converted using the [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method.</span></span>
- <span data-ttu-id="b63cf-127">また、これを使用するには、Attachment ID を REST 用に形式設定された ID に変換する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="b63cf-127">Note you must also convert Attachment ID to a REST-formatted ID in order to use it.</span></span> <span data-ttu-id="b63cf-128">ID を変換する必要がある理由は、EWS ID に URL セーフ以外の値が含まれている可能性があり、その場合は REST で問題が発生するためです。</span><span class="sxs-lookup"><span data-stu-id="b63cf-128">The reason the IDs must be converted is that EWS IDs can contain non-URL safe values which will cause problems for REST.</span></span>

<span data-ttu-id="b63cf-129">[Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname) プロパティを確認することにより、アドインは読み込まれる Outlook クライアントを判別できます。</span><span class="sxs-lookup"><span data-stu-id="b63cf-129">Your add-in can determine which Outlook client it is loaded in by checking the [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname) property.</span></span>

### <a name="example"></a><span data-ttu-id="b63cf-130">例</span><span class="sxs-lookup"><span data-stu-id="b63cf-130">Example</span></span>

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

## <a name="get-the-rest-api-url"></a><span data-ttu-id="b63cf-131">REST API URL を取得する</span><span class="sxs-lookup"><span data-stu-id="b63cf-131">Get the REST API URL</span></span>

<span data-ttu-id="b63cf-p108">REST API を呼び出すためにアドインで必要な情報の最終部分は、API 要求の送信に使用するホスト名です。この情報は [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) プロパティにあります。</span><span class="sxs-lookup"><span data-stu-id="b63cf-p108">The final piece of information your add-in needs to call the REST API is the hostname it should use to send API requests. This information is in the [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) property.</span></span>

### <a name="example"></a><span data-ttu-id="b63cf-134">例</span><span class="sxs-lookup"><span data-stu-id="b63cf-134">Example</span></span>

```js
// Example: https://outlook.office.com
var restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a><span data-ttu-id="b63cf-135">API を呼び出す</span><span class="sxs-lookup"><span data-stu-id="b63cf-135">Call the API</span></span>

<span data-ttu-id="b63cf-136">アドインがアクセス トークン、アイテム ID、および REST API URL を取得すると、REST API を呼び出すバックエンド サービスにその情報を渡すか、AJAX を使用して直接呼び出すことができるようになります。</span><span class="sxs-lookup"><span data-stu-id="b63cf-136">After your add-in has the access token, item ID, and REST API URL, it can either pass that information to a back-end service which calls the REST API, or it can call it directly using AJAX.</span></span> <span data-ttu-id="b63cf-137">次の例は、Outlook Mail REST API を呼び出して現在のメッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="b63cf-137">The following example calls the Outlook Mail REST API to get the current message.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b63cf-138">オンプレミスの Exchange 展開では、AJAX または類似のライブラリを使用するクライアント側の要求は、そのサーバーセットアップで CORS がサポートされていないので失敗します。</span><span class="sxs-lookup"><span data-stu-id="b63cf-138">For on-premises Exchange deployments, client-side requests using AJAX or similar libraries fail because CORS isn't supported in that server setup.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="b63cf-139">関連項目</span><span class="sxs-lookup"><span data-stu-id="b63cf-139">See also</span></span>

- <span data-ttu-id="b63cf-140">Outlook アドインから REST API を呼び出す例については、GitHub の [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="b63cf-140">For an example that calls the REST APIs from an Outlook add-in, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
- <span data-ttu-id="b63cf-141">Outlook REST API は、Microsoft Graph エンドポイントからでも利用できますが、アドインでアクセス トークンを取得する方法など、重要な違いがいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="b63cf-141">Outlook REST APIs are also available through the Microsoft Graph endpoint but there are some key differences, including how your add-in gets an access token.</span></span> <span data-ttu-id="b63cf-142">詳細については、「[Microsoft Graph を介して使用する Outlook REST API](/outlook/rest/index#outlook-rest-api-via-microsoft-graph)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b63cf-142">For more information, see [Outlook REST API via Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).</span></span>