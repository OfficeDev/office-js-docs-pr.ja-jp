---
title: Outlook アドイン API 要件セット 1.1
description: ''
ms.date: 10/11/2018
localization_priority: Normal
ms.openlocfilehash: a074d0e38f8d872f0d75a68851aef947989c625e
ms.sourcegitcommit: c4d6ecdc41ea67291b6d155c3b246e31ec2e38b7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/19/2019
ms.locfileid: "30600257"
---
# <a name="outlook-add-in-api-requirement-set-11"></a><span data-ttu-id="9c315-102">Outlook アドイン API 要件セット 1.1</span><span class="sxs-lookup"><span data-stu-id="9c315-102">Outlook add-in API requirement set 1.1</span></span>

<span data-ttu-id="9c315-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="9c315-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="9c315-104">このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="9c315-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-11"></a><span data-ttu-id="9c315-105">1.1 の新機能</span><span class="sxs-lookup"><span data-stu-id="9c315-105">What's new in 1.1?</span></span>

<span data-ttu-id="9c315-p101">要件セット 1.1 には、要件セット 1.0 のすべての機能が含まれています。アドインでメッセージと予定の本文にアクセスする機能、および現在のアイテムを変更する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="9c315-p101">Requirement set 1.1 includes all of the features of Requirement set 1.0. It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.</span></span>

### <a name="change-log"></a><span data-ttu-id="9c315-108">変更ログ</span><span class="sxs-lookup"><span data-stu-id="9c315-108">Change log</span></span>

- <span data-ttu-id="9c315-109">[Body](/javascript/api/outlook_1_1/office.body) オブジェクトが追加されました。Outlook アドインでアイテムのコンテンツを追加および更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="9c315-109">Added [Body](/javascript/api/outlook_1_1/office.body) object: Provides methods for adding and updating the content of an item in an Outlook add-in.</span></span>
- <span data-ttu-id="9c315-110">[Location](/javascript/api/outlook_1_1/office.location) オブジェクトが追加されました。Outlook アドインで会議の場所を取得し設定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="9c315-110">Added [Location](/javascript/api/outlook_1_1/office.location) object: Provides methods to get and set the location of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="9c315-111">[Recipients](/javascript/api/outlook_1_1/office.recipients) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの受信者を取得および設定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="9c315-111">Added [Recipients](/javascript/api/outlook_1_1/office.recipients) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="9c315-112">[Subject](/javascript/api/outlook_1_1/office.subject) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの件名を取得および設定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="9c315-112">Added [Subject](/javascript/api/outlook_1_1/office.subject) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="9c315-113">[Time](/javascript/api/outlook_1_1/office.time) オブジェクトが追加されました。Outlook アドインで会議の開始時刻と終了時刻を取得および設定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="9c315-113">Added [Time](/javascript/api/outlook_1_1/office.time) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="9c315-114">[Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) が追加されました。ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="9c315-114">Added [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): Adds a file to a message or appointment as an attachment.</span></span>
- <span data-ttu-id="9c315-115">[Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback) が追加されました。メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="9c315-115">Added [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback): Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>
- <span data-ttu-id="9c315-116">[Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) が追加されました。メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="9c315-116">Added [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback): Removes an attachment from a message or appointment.</span></span>
- <span data-ttu-id="9c315-117">[Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body) が追加されました。アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="9c315-117">Added [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body): Gets an object that provides methods for manipulating the body of an item.</span></span>
- <span data-ttu-id="9c315-118">メッセージの[bcc](office.context.mailbox.item.md#bcc-recipients)行を追加しました。</span><span class="sxs-lookup"><span data-stu-id="9c315-118">Added [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipients) line of a message.</span></span>
- <span data-ttu-id="9c315-119">[Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype) が追加されました。予定の受信者の種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="9c315-119">Added [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype): Specifies the type of recipient for an appointment.</span></span>

## <a name="see-also"></a><span data-ttu-id="9c315-120">関連項目</span><span class="sxs-lookup"><span data-stu-id="9c315-120">See also</span></span>

- [<span data-ttu-id="9c315-121">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="9c315-121">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="9c315-122">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="9c315-122">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="9c315-123">作業の開始</span><span class="sxs-lookup"><span data-stu-id="9c315-123">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)
