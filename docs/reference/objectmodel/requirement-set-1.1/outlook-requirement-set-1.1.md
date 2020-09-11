---
title: Outlook アドイン API 要件セット 1.1
description: Outlook アドインおよび Office JavaScript Api for the Mailbox API 1.1 の一部として導入された機能と Api。
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: f93b6d582043641903b362121c6e5eaf89c2ad1c
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431375"
---
# <a name="outlook-add-in-api-requirement-set-11"></a><span data-ttu-id="5f2d2-103">Outlook アドイン API 要件セット 1.1</span><span class="sxs-lookup"><span data-stu-id="5f2d2-103">Outlook add-in API requirement set 1.1</span></span>

<span data-ttu-id="5f2d2-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span> <span data-ttu-id="5f2d2-105">Outlook JavaScript API 1.1 (メールボックス 1.1) は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-105">Outlook JavaScript API 1.1 (Mailbox 1.1) is the first version of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="5f2d2-106">このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-106">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-11"></a><span data-ttu-id="5f2d2-107">1.1 の新機能</span><span class="sxs-lookup"><span data-stu-id="5f2d2-107">What's new in 1.1?</span></span>

<span data-ttu-id="5f2d2-108">要件セット1.1 には、Outlook でサポートされているすべての [共通 API 要件セット](../../requirement-sets/office-add-in-requirement-sets.md) が含まれています。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-108">Requirement set 1.1 includes all of the [Common API requirement sets](../../requirement-sets/office-add-in-requirement-sets.md) supported in Outlook.</span></span> <span data-ttu-id="5f2d2-109">アドインでメッセージと予定の本文にアクセスする機能、および現在のアイテムを変更する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-109">It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.</span></span>

### <a name="change-log"></a><span data-ttu-id="5f2d2-110">変更ログ</span><span class="sxs-lookup"><span data-stu-id="5f2d2-110">Change log</span></span>

- <span data-ttu-id="5f2d2-111">[Body](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true) オブジェクトが追加されました。Outlook アドインでアイテムのコンテンツを追加および更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-111">Added [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true) object: Provides methods for adding and updating the content of an item in an Outlook add-in.</span></span>
- <span data-ttu-id="5f2d2-112">[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1&preserve-view=true) オブジェクトが追加されました。Outlook アドインで会議の場所を取得し設定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-112">Added [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1&preserve-view=true) object: Provides methods to get and set the location of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="5f2d2-113">[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの受信者を取得および設定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-113">Added [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="5f2d2-114">[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1&preserve-view=true) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの件名を取得および設定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-114">Added [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1&preserve-view=true) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="5f2d2-115">[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1&preserve-view=true) オブジェクトが追加されました。Outlook アドインで会議の開始時刻と終了時刻を取得および設定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-115">Added [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1&preserve-view=true) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="5f2d2-116">[Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) が追加されました。ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-116">Added [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): Adds a file to a message or appointment as an attachment.</span></span>
- <span data-ttu-id="5f2d2-117">[Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods) が追加されました。メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-117">Added [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods): Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>
- <span data-ttu-id="5f2d2-118">[Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods) が追加されました。メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-118">Added [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods): Removes an attachment from a message or appointment.</span></span>
- <span data-ttu-id="5f2d2-119">[Office.context.mailbox.item.body](office.context.mailbox.item.md#properties) が追加されました。アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-119">Added [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties): Gets an object that provides methods for manipulating the body of an item.</span></span>
- <span data-ttu-id="5f2d2-120">メッセージの [bcc](office.context.mailbox.item.md#properties) 行を追加しました。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-120">Added [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#properties) line of a message.</span></span>
- <span data-ttu-id="5f2d2-121">[Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1&preserve-view=true) が追加されました。予定の受信者の種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="5f2d2-121">Added [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1&preserve-view=true): Specifies the type of recipient for an appointment.</span></span>

## <a name="see-also"></a><span data-ttu-id="5f2d2-122">関連項目</span><span class="sxs-lookup"><span data-stu-id="5f2d2-122">See also</span></span>

- [<span data-ttu-id="5f2d2-123">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="5f2d2-123">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="5f2d2-124">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="5f2d2-124">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="5f2d2-125">概要</span><span class="sxs-lookup"><span data-stu-id="5f2d2-125">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="5f2d2-126">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="5f2d2-126">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
