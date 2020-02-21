---
title: Outlook アドイン API 要件セット 1.1
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 2ecd337cd838cd6dd9deb4fe5e77ee789106f3f9
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165469"
---
# <a name="outlook-add-in-api-requirement-set-11"></a><span data-ttu-id="5cc62-102">Outlook アドイン API 要件セット 1.1</span><span class="sxs-lookup"><span data-stu-id="5cc62-102">Outlook add-in API requirement set 1.1</span></span>

<span data-ttu-id="5cc62-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="5cc62-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span> <span data-ttu-id="5cc62-104">Outlook JavaScript API 1.1 (メールボックス 1.1) は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="5cc62-104">Outlook JavaScript API 1.1 (Mailbox 1.1) is the first version of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="5cc62-105">このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="5cc62-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-11"></a><span data-ttu-id="5cc62-106">1.1 の新機能</span><span class="sxs-lookup"><span data-stu-id="5cc62-106">What's new in 1.1?</span></span>

<span data-ttu-id="5cc62-107">要件セット1.1 には、Outlook でサポートされているすべての[共通 API 要件セット](../../requirement-sets/office-add-in-requirement-sets.md)が含まれています。</span><span class="sxs-lookup"><span data-stu-id="5cc62-107">Requirement set 1.1 includes all of the [Common API requirement sets](../../requirement-sets/office-add-in-requirement-sets.md) supported in Outlook.</span></span> <span data-ttu-id="5cc62-108">アドインでメッセージと予定の本文にアクセスする機能、および現在のアイテムを変更する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="5cc62-108">It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.</span></span>

### <a name="change-log"></a><span data-ttu-id="5cc62-109">変更ログ</span><span class="sxs-lookup"><span data-stu-id="5cc62-109">Change log</span></span>

- <span data-ttu-id="5cc62-110">[Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) オブジェクトが追加されました。Outlook アドインでアイテムのコンテンツを追加および更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="5cc62-110">Added [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) object: Provides methods for adding and updating the content of an item in an Outlook add-in.</span></span>
- <span data-ttu-id="5cc62-111">[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) オブジェクトが追加されました。Outlook アドインで会議の場所を取得し設定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="5cc62-111">Added [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) object: Provides methods to get and set the location of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="5cc62-112">[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの受信者を取得および設定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="5cc62-112">Added [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="5cc62-113">[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの件名を取得および設定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="5cc62-113">Added [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="5cc62-114">[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) オブジェクトが追加されました。Outlook アドインで会議の開始時刻と終了時刻を取得および設定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="5cc62-114">Added [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="5cc62-115">[Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) が追加されました。ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="5cc62-115">Added [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): Adds a file to a message or appointment as an attachment.</span></span>
- <span data-ttu-id="5cc62-116">[Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods) が追加されました。メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="5cc62-116">Added [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods): Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>
- <span data-ttu-id="5cc62-117">[Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods) が追加されました。メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="5cc62-117">Added [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods): Removes an attachment from a message or appointment.</span></span>
- <span data-ttu-id="5cc62-118">[Office.context.mailbox.item.body](office.context.mailbox.item.md#properties) が追加されました。アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="5cc62-118">Added [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties): Gets an object that provides methods for manipulating the body of an item.</span></span>
- <span data-ttu-id="5cc62-119">メッセージの[bcc](office.context.mailbox.item.md#properties)行を追加しました。</span><span class="sxs-lookup"><span data-stu-id="5cc62-119">Added [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#properties) line of a message.</span></span>
- <span data-ttu-id="5cc62-120">[Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1) が追加されました。予定の受信者の種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="5cc62-120">Added [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1): Specifies the type of recipient for an appointment.</span></span>

## <a name="see-also"></a><span data-ttu-id="5cc62-121">関連項目</span><span class="sxs-lookup"><span data-stu-id="5cc62-121">See also</span></span>

- [<span data-ttu-id="5cc62-122">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="5cc62-122">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="5cc62-123">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="5cc62-123">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="5cc62-124">概要</span><span class="sxs-lookup"><span data-stu-id="5cc62-124">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="5cc62-125">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="5cc62-125">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
