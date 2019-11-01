---
title: Outlook アドイン API 要件セット 1.1
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 312d40d499531eb6f93d3b1555bfb057cd4651d6
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901956"
---
# <a name="outlook-add-in-api-requirement-set-11"></a><span data-ttu-id="dc45e-102">Outlook アドイン API 要件セット 1.1</span><span class="sxs-lookup"><span data-stu-id="dc45e-102">Outlook add-in API requirement set 1.1</span></span>

<span data-ttu-id="dc45e-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="dc45e-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="dc45e-104">このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="dc45e-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-11"></a><span data-ttu-id="dc45e-105">1.1 の新機能</span><span class="sxs-lookup"><span data-stu-id="dc45e-105">What's new in 1.1?</span></span>

<span data-ttu-id="dc45e-p101">要件セット 1.1 には、要件セット 1.0 のすべての機能が含まれています。アドインでメッセージと予定の本文にアクセスする機能、および現在のアイテムを変更する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="dc45e-p101">Requirement set 1.1 includes all of the features of Requirement set 1.0. It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.</span></span>

### <a name="change-log"></a><span data-ttu-id="dc45e-108">変更ログ</span><span class="sxs-lookup"><span data-stu-id="dc45e-108">Change log</span></span>

- <span data-ttu-id="dc45e-109">[Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) オブジェクトが追加されました。Outlook アドインでアイテムのコンテンツを追加および更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="dc45e-109">Added [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) object: Provides methods for adding and updating the content of an item in an Outlook add-in.</span></span>
- <span data-ttu-id="dc45e-110">[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) オブジェクトが追加されました。Outlook アドインで会議の場所を取得し設定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="dc45e-110">Added [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) object: Provides methods to get and set the location of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="dc45e-111">[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの受信者を取得および設定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="dc45e-111">Added [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="dc45e-112">[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの件名を取得および設定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="dc45e-112">Added [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="dc45e-113">[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) オブジェクトが追加されました。Outlook アドインで会議の開始時刻と終了時刻を取得および設定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="dc45e-113">Added [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="dc45e-114">[Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) が追加されました。ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="dc45e-114">Added [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): Adds a file to a message or appointment as an attachment.</span></span>
- <span data-ttu-id="dc45e-115">[Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback) が追加されました。メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="dc45e-115">Added [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback): Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>
- <span data-ttu-id="dc45e-116">[Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) が追加されました。メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="dc45e-116">Added [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback): Removes an attachment from a message or appointment.</span></span>
- <span data-ttu-id="dc45e-117">[Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body) が追加されました。アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="dc45e-117">Added [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body): Gets an object that provides methods for manipulating the body of an item.</span></span>
- <span data-ttu-id="dc45e-118">メッセージの[bcc](office.context.mailbox.item.md#bcc-recipients)行を追加しました。</span><span class="sxs-lookup"><span data-stu-id="dc45e-118">Added [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipients) line of a message.</span></span>
- <span data-ttu-id="dc45e-119">[Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1) が追加されました。予定の受信者の種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="dc45e-119">Added [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1): Specifies the type of recipient for an appointment.</span></span>

## <a name="see-also"></a><span data-ttu-id="dc45e-120">関連項目</span><span class="sxs-lookup"><span data-stu-id="dc45e-120">See also</span></span>

- [<span data-ttu-id="dc45e-121">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="dc45e-121">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="dc45e-122">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="dc45e-122">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="dc45e-123">概要</span><span class="sxs-lookup"><span data-stu-id="dc45e-123">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="dc45e-124">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="dc45e-124">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
