---
title: インターネットヘッダーを取得および設定する
description: Outlook アドインでメッセージのインターネットヘッダーを取得および設定する方法について説明します。
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: 1b6bdbbe77998ce92ea1b1b43874a32a30aa160a
ms.sourcegitcommit: 0fdb78cefa669b727b817614a4147a46d249a0ed
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/28/2020
ms.locfileid: "43930289"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a><span data-ttu-id="dcb5f-103">Outlook アドインでメッセージのインターネットヘッダーを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="dcb5f-103">Get and set internet headers on a message in an Outlook add-in</span></span>

## <a name="background"></a><span data-ttu-id="dcb5f-104">背景</span><span class="sxs-lookup"><span data-stu-id="dcb5f-104">Background</span></span>

<span data-ttu-id="dcb5f-105">Outlook アドインの開発での一般的な要件は、アドインに関連付けられたカスタムプロパティをさまざまなレベルで保存することです。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-105">A common requirement in Outlook add-ins development is to store custom properties associated with an add-in at different levels.</span></span> <span data-ttu-id="dcb5f-106">現在、カスタムプロパティは、アイテムまたはメールボックスのレベルで保存されています。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-106">At present, custom properties are stored at the item or mailbox level.</span></span>

- <span data-ttu-id="dcb5f-107">アイテムレベル-特定のアイテムに適用されるプロパティについては、 [CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-107">Item level - For properties that apply to a specific item, use the [CustomProperties](/javascript/api/outlook/office.customproperties) object.</span></span> <span data-ttu-id="dcb5f-108">たとえば、電子メールの送信者に関連付けられている顧客コードを格納します。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-108">For example, store a customer code associated with the person who sent the email.</span></span>
- <span data-ttu-id="dcb5f-109">メールボックスレベル-ユーザーのメールボックス内のすべてのメールアイテムに適用されるプロパティについては、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-109">Mailbox level - For properties that apply to all the mail items in the user's mailbox, use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object.</span></span> <span data-ttu-id="dcb5f-110">たとえば、ユーザーの設定を保存して、特定の規模で温度を表示します。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-110">For example, store a user's preference to show the temperature in a particular scale.</span></span>

<span data-ttu-id="dcb5f-111">アイテムが Exchange サーバーから離脱した後、両方の種類のプロパティは保持されないため、電子メールの受信者は、アイテムに設定されているプロパティを取得できません。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-111">Both types of properties are not preserved after the item leaves the Exchange server so the email recipients can't get any properties set on the item.</span></span> <span data-ttu-id="dcb5f-112">そのため、開発者はこれらの設定やその他の MIME プロパティにアクセスして、読み取りシナリオを改善することはできません。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-112">Therefore, developers can't access those settings or other MIME properties to enable better read scenarios.</span></span>

<span data-ttu-id="dcb5f-113">EWS 要求を通じてインターネットヘッダーを設定する方法はありますが、一部のシナリオでは EWS 要求が機能しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-113">While there's a way for you to set the internet headers through EWS requests, in some scenarios making an EWS request won't work.</span></span> <span data-ttu-id="dcb5f-114">たとえば、Outlook デスクトップの新規作成モードでは、アイテム id はキャッシュモード `saveAsync` で同期されません。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-114">For example, in Compose mode on Outlook desktop, the item id isn't synced on `saveAsync` in cached mode.</span></span>

> [!TIP]
> <span data-ttu-id="dcb5f-115">これらのオプションの使用の詳細については[、「Outlook アドインのアドインメタデータを取得および設定](metadata-for-an-outlook-add-in.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-115">See [Get and set add-in metadata for an Outlook add-in](metadata-for-an-outlook-add-in.md) to learn more about using these options.</span></span>

## <a name="purpose-of-the-internet-headers-api"></a><span data-ttu-id="dcb5f-116">インターネットヘッダー API の目的</span><span class="sxs-lookup"><span data-stu-id="dcb5f-116">Purpose of the internet headers API</span></span>

<span data-ttu-id="dcb5f-117">[要件セット 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)で導入されたインターネットヘッダー api を使用すると、開発者は次のことを行うことができます。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-117">Introduced in [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), the internet headers APIs enable developers to:</span></span>

- <span data-ttu-id="dcb5f-118">すべてのクライアント間で Exchange を残した後に保持されるメールについての情報をスタンプします。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-118">Stamp information on an email that persists after it leaves Exchange across all clients.</span></span>
- <span data-ttu-id="dcb5f-119">メールの読み取りシナリオにおいて、すべてのクライアント間で Exchange のメールが残された後に保持される電子メールの情報を読み取ります。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-119">Read information on an email that persisted after the email left Exchange across all clients in mail read scenarios.</span></span>
- <span data-ttu-id="dcb5f-120">電子メールの MIME ヘッダー全体にアクセスします。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-120">Access the entire MIME header of the email.</span></span>

![インターネットヘッダーの図](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a><span data-ttu-id="dcb5f-126">メッセージの作成中にインターネットヘッダーを設定する</span><span class="sxs-lookup"><span data-stu-id="dcb5f-126">Set internet headers while composing a message</span></span>

<span data-ttu-id="dcb5f-127">[InternetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders)プロパティを使用して、新規作成モードで現在のメッセージに配置するカスタムインターネットヘッダーを管理します。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-127">Try using the [item.internetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders) property to manage the custom internet headers you place on the current message in Compose mode.</span></span>

### <a name="set-get-and-remove-custom-headers-example"></a><span data-ttu-id="dcb5f-128">カスタムヘッダーの設定、取得、および削除の例</span><span class="sxs-lookup"><span data-stu-id="dcb5f-128">Set, get, and remove custom headers example</span></span>

<span data-ttu-id="dcb5f-129">次の例は、カスタムヘッダーを設定、取得、および削除する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-129">The following example shows how to set, get, and remove custom headers.</span></span>

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

## <a name="get-internet-headers-while-reading-a-message"></a><span data-ttu-id="dcb5f-130">メッセージの読み取り中にインターネットヘッダーを取得する</span><span class="sxs-lookup"><span data-stu-id="dcb5f-130">Get internet headers while reading a message</span></span>

<span data-ttu-id="dcb5f-131">現在のメッセージのインターネットヘッダーを閲覧モードで取得するには、 [getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-)を呼び出してください。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-131">Try calling [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) to get internet headers on the current message in Read mode.</span></span>

### <a name="get-sender-preferences-from-current-mime-headers-example"></a><span data-ttu-id="dcb5f-132">現在の MIME ヘッダーの送信者の設定を取得する例</span><span class="sxs-lookup"><span data-stu-id="dcb5f-132">Get sender preferences from current MIME headers example</span></span>

<span data-ttu-id="dcb5f-133">前のセクションの例では、次のコードは、現在の電子メールの MIME ヘッダーから送信者の設定を取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-133">Building on the example from the previous section, the following code shows how to get the sender's preferences from the current email's MIME headers.</span></span>

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
> <span data-ttu-id="dcb5f-134">このサンプルは、単純なケースで機能します。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-134">This sample works for simple cases.</span></span> <span data-ttu-id="dcb5f-135">複雑な情報取得 ( [RFC 2822](https://tools.ietf.org/html/rfc2822)で説明されているように、複数インスタンスのヘッダー、または折りたたまれた値など) を取得するには、適切な MIME 解析ライブラリを使用してみてください。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-135">For more complex information retrieval (for example, multi-instance headers or folded values as described in [RFC 2822](https://tools.ietf.org/html/rfc2822)), try using an appropriate MIME-parsing library.</span></span>

## <a name="recommended-practices"></a><span data-ttu-id="dcb5f-136">推奨プラクティス</span><span class="sxs-lookup"><span data-stu-id="dcb5f-136">Recommended practices</span></span>

<span data-ttu-id="dcb5f-137">現時点では、インターネットヘッダーはユーザーのメールボックス上の有限リソースです。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-137">Currently, internet headers are a finite resource on a user's mailbox.</span></span> <span data-ttu-id="dcb5f-138">クォータが不足している場合は、そのメールボックスにより多くのインターネットヘッダーを作成することはできません。これにより、これに依存するクライアントから予期しない動作が発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-138">When the quota is exhausted, you can't create any more internet headers on that mailbox, which can result in unexpected behavior from clients that rely on this to function.</span></span>

<span data-ttu-id="dcb5f-139">アドインでインターネットヘッダーを作成するときには、次のガイドラインを適用します。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-139">Apply the following guidelines when you create internet headers in your add-in.</span></span>

- <span data-ttu-id="dcb5f-140">必要なヘッダーの最小数を作成します。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-140">Create the minimum number of headers required.</span></span>
- <span data-ttu-id="dcb5f-141">後で再利用して値を更新できるように、名前のヘッダー。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-141">Name headers so that you can reuse and update their values later.</span></span> <span data-ttu-id="dcb5f-142">そのため、ユーザーの入力やタイムスタンプなどに基づいて、変数の方法でヘッダーに名前を付けることは避けてください。</span><span class="sxs-lookup"><span data-stu-id="dcb5f-142">As such, avoid naming headers in a variable manner (for example, based on user input, timestamp, etc.).</span></span>

## <a name="see-also"></a><span data-ttu-id="dcb5f-143">関連項目</span><span class="sxs-lookup"><span data-stu-id="dcb5f-143">See also</span></span>

- [<span data-ttu-id="dcb5f-144">Outlook アドインのアドイン メタデータを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="dcb5f-144">Get and set add-in metadata for an Outlook add-in</span></span>](metadata-for-an-outlook-add-in.md)
