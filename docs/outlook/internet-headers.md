---
title: インターネットヘッダーを取得および設定する
description: Outlook アドインでメッセージのインターネットヘッダーを取得および設定する方法について説明します。
ms.date: 11/04/2019
localization_priority: Normal
ms.openlocfilehash: d7f38b7564683ce51ed0bd840480b4a8b2040bf6
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166565"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a><span data-ttu-id="bdff8-103">Outlook アドインでメッセージのインターネットヘッダーを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="bdff8-103">Get and set internet headers on a message in an Outlook add-in</span></span>

## <a name="background"></a><span data-ttu-id="bdff8-104">背景</span><span class="sxs-lookup"><span data-stu-id="bdff8-104">Background</span></span>

<span data-ttu-id="bdff8-105">Outlook アドインの開発での一般的な要件は、アドインに関連付けられたカスタムプロパティをさまざまなレベルで保存することです。</span><span class="sxs-lookup"><span data-stu-id="bdff8-105">A common requirement in Outlook add-ins development is to store custom properties associated with an add-in at different levels.</span></span> <span data-ttu-id="bdff8-106">現在、カスタムプロパティは、アイテムまたはメールボックスのレベルで保存されています。</span><span class="sxs-lookup"><span data-stu-id="bdff8-106">At present, custom properties are stored at the item or mailbox level.</span></span>

- <span data-ttu-id="bdff8-107">アイテムレベル-特定のアイテムに適用されるプロパティについては、 [CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="bdff8-107">Item level - For properties that apply to a specific item, use the [CustomProperties](/javascript/api/outlook/office.customproperties) object.</span></span> <span data-ttu-id="bdff8-108">たとえば、電子メールの送信者に関連付けられている顧客コードを格納します。</span><span class="sxs-lookup"><span data-stu-id="bdff8-108">For example, store a customer code associated with the person who sent the email.</span></span>
- <span data-ttu-id="bdff8-109">メールボックスレベル-ユーザーのメールボックス内のすべてのメールアイテムに適用されるプロパティについては、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="bdff8-109">Mailbox level - For properties that apply to all the mail items in the user's mailbox, use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object.</span></span> <span data-ttu-id="bdff8-110">たとえば、ユーザーの設定を保存して、特定の規模で温度を表示します。</span><span class="sxs-lookup"><span data-stu-id="bdff8-110">For example, store a user's preference to show the temperature in a particular scale.</span></span>

<span data-ttu-id="bdff8-111">アイテムが Exchange サーバーから離脱した後、両方の種類のプロパティは保持されないため、電子メールの受信者は、アイテムに設定されているプロパティを取得できません。</span><span class="sxs-lookup"><span data-stu-id="bdff8-111">Both types of properties are not preserved after the item leaves the Exchange server so the email recipients can't get any properties set on the item.</span></span> <span data-ttu-id="bdff8-112">そのため、開発者はこれらの設定やその他の MIME プロパティにアクセスして、読み取りシナリオを改善することはできません。</span><span class="sxs-lookup"><span data-stu-id="bdff8-112">Therefore, developers can't access those settings or other MIME properties to enable better read scenarios.</span></span>

<span data-ttu-id="bdff8-113">EWS 要求を通じてインターネットヘッダーを設定する方法はありますが、一部のシナリオでは EWS 要求が機能しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="bdff8-113">While there's a way for you to set the internet headers through EWS requests, in some scenarios making an EWS request won't work.</span></span> <span data-ttu-id="bdff8-114">たとえば、Outlook デスクトップの新規作成モードでは、アイテム id はキャッシュモード `saveAsync` で同期されません。</span><span class="sxs-lookup"><span data-stu-id="bdff8-114">For example, in Compose mode on Outlook desktop, the item id isn't synced on `saveAsync` in cached mode.</span></span>

> [!TIP]
> <span data-ttu-id="bdff8-115">これらのオプションの使用の詳細については[、「Outlook アドインのアドインメタデータを取得および設定](metadata-for-an-outlook-add-in.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bdff8-115">See [Get and set add-in metadata for an Outlook add-in](metadata-for-an-outlook-add-in.md) to learn more about using these options.</span></span>

## <a name="purpose-of-the-internet-headers-api"></a><span data-ttu-id="bdff8-116">インターネットヘッダー API の目的</span><span class="sxs-lookup"><span data-stu-id="bdff8-116">Purpose of the internet headers API</span></span>

<span data-ttu-id="bdff8-117">要件セット1.8 で導入されたインターネットヘッダー Api を使用すると、開発者は次のことを行うことができます。</span><span class="sxs-lookup"><span data-stu-id="bdff8-117">Introduced in requirement set 1.8, the internet headers APIs enable developers to:</span></span>

- <span data-ttu-id="bdff8-118">すべてのクライアント間で Exchange を残した後に保持されるメールについての情報をスタンプします。</span><span class="sxs-lookup"><span data-stu-id="bdff8-118">Stamp information on an email that persists after it leaves Exchange across all clients.</span></span>
- <span data-ttu-id="bdff8-119">メールの読み取りシナリオにおいて、すべてのクライアント間で Exchange のメールが残された後に保持される電子メールの情報を読み取ります。</span><span class="sxs-lookup"><span data-stu-id="bdff8-119">Read information on an email that persisted after the email left Exchange across all clients in mail read scenarios.</span></span>
- <span data-ttu-id="bdff8-120">電子メールの MIME ヘッダー全体にアクセスします。</span><span class="sxs-lookup"><span data-stu-id="bdff8-120">Access the entire MIME header of the email.</span></span>

## <a name="set-internet-headers-while-composing-a-message"></a><span data-ttu-id="bdff8-121">メッセージの作成中にインターネットヘッダーを設定する</span><span class="sxs-lookup"><span data-stu-id="bdff8-121">Set internet headers while composing a message</span></span>

<span data-ttu-id="bdff8-122">[InternetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders)プロパティを使用して、新規作成モードで現在のメッセージに配置するカスタムインターネットヘッダーを管理します。</span><span class="sxs-lookup"><span data-stu-id="bdff8-122">Try using the [item.internetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders) property to manage the custom internet headers you place on the current message in Compose mode.</span></span>

### <a name="set-get-and-remove-custom-headers-example"></a><span data-ttu-id="bdff8-123">カスタムヘッダーの設定、取得、および削除の例</span><span class="sxs-lookup"><span data-stu-id="bdff8-123">Set, get, and remove custom headers example</span></span>

<span data-ttu-id="bdff8-124">次の例は、カスタムヘッダーを設定、取得、および削除する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="bdff8-124">The following example shows how to set, get, and remove custom headers.</span></span>

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

## <a name="get-internet-headers-while-reading-a-message"></a><span data-ttu-id="bdff8-125">メッセージの読み取り中にインターネットヘッダーを取得する</span><span class="sxs-lookup"><span data-stu-id="bdff8-125">Get internet headers while reading a message</span></span>

<span data-ttu-id="bdff8-126">現在のメッセージのインターネットヘッダーを閲覧モードで取得するには、 [getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-)を呼び出してください。</span><span class="sxs-lookup"><span data-stu-id="bdff8-126">Try calling [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) to get internet headers on the current message in Read mode.</span></span>

### <a name="get-sender-preferences-from-current-mime-headers-example"></a><span data-ttu-id="bdff8-127">現在の MIME ヘッダーの送信者の設定を取得する例</span><span class="sxs-lookup"><span data-stu-id="bdff8-127">Get sender preferences from current MIME headers example</span></span>

<span data-ttu-id="bdff8-128">前のセクションの例では、次のコードは、現在の電子メールの MIME ヘッダーから送信者の設定を取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="bdff8-128">Building on the example from the previous section, the following code shows how to get the sender's preferences from the current email's MIME headers.</span></span>

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
> <span data-ttu-id="bdff8-129">このサンプルは、単純なケースで機能します。</span><span class="sxs-lookup"><span data-stu-id="bdff8-129">This sample works for simple cases.</span></span> <span data-ttu-id="bdff8-130">複雑な情報取得 ( [RFC 2822](https://tools.ietf.org/html/rfc2822)で説明されているように、複数インスタンスのヘッダー、または折りたたまれた値など) については、適切な MIME 解析ライブラリを使用してみてください。</span><span class="sxs-lookup"><span data-stu-id="bdff8-130">For more complex information retrieval (e.g., multi-instance headers or folded values as described in [RFC 2822](https://tools.ietf.org/html/rfc2822)), try using an appropriate MIME-parsing library.</span></span>

## <a name="see-also"></a><span data-ttu-id="bdff8-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="bdff8-131">See also</span></span>

- [<span data-ttu-id="bdff8-132">Outlook アドインのアドイン メタデータを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="bdff8-132">Get and set add-in metadata for an Outlook add-in</span></span>](metadata-for-an-outlook-add-in.md)
