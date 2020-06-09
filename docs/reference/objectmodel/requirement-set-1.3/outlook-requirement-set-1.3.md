---
title: Outlook アドイン API 要件セット 1.3
description: Outlook アドインおよび Office JavaScript Api for the Mailbox API 1.3 の一部として導入された機能と Api。
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 2f51a275e00853b2b3626c710a4c072a83ba8c0a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611997"
---
# <a name="outlook-add-in-api-requirement-set-13"></a><span data-ttu-id="cebe8-103">Outlook アドイン API 要件セット 1.3</span><span class="sxs-lookup"><span data-stu-id="cebe8-103">Outlook add-in API requirement set 1.3</span></span>

<span data-ttu-id="cebe8-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="cebe8-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="cebe8-105">このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="cebe8-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-13"></a><span data-ttu-id="cebe8-106">1.3 の新機能</span><span class="sxs-lookup"><span data-stu-id="cebe8-106">What's new in 1.3?</span></span>

<span data-ttu-id="cebe8-p101">要件セット 1.3 には、[要件セット 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) のすべての機能が含まれています。次の機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="cebe8-p101">Requirement set 1.3 includes all of the features of [Requirement set 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). It added the following features.</span></span>

- <span data-ttu-id="cebe8-109">[アドイン コマンド](../../../outlook/add-in-commands-for-outlook.md)のサポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="cebe8-109">Added support for [add-in commands](../../../outlook/add-in-commands-for-outlook.md).</span></span>
- <span data-ttu-id="cebe8-110">作成中のアイテムを保存または閉じる機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="cebe8-110">Added ability to save or close an item being composed.</span></span>
- <span data-ttu-id="cebe8-111">アドインが本文全体を取得または設定できるようにする、拡張された[Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="cebe8-111">Enhanced [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3) object to allow add-ins to get or set the entire body.</span></span>
- <span data-ttu-id="cebe8-112">EWS 形式と REST 形式間で ID を変換する変換メソッドが追加されました。</span><span class="sxs-lookup"><span data-stu-id="cebe8-112">Added conversion methods to convert IDs between EWS and REST formats.</span></span>
- <span data-ttu-id="cebe8-113">アイテム上にある情報バーに通知メッセージを追加する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="cebe8-113">Added ability to add notification messages to the info bar on items.</span></span>

### <a name="change-log"></a><span data-ttu-id="cebe8-114">変更ログ</span><span class="sxs-lookup"><span data-stu-id="cebe8-114">Change log</span></span>

- <span data-ttu-id="cebe8-115">[Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) が追加されました。現在の本文を指定された形式で返します。</span><span class="sxs-lookup"><span data-stu-id="cebe8-115">Added [Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-): Returns the current body in a specified format.</span></span>
- <span data-ttu-id="cebe8-116">[Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#setasync-data--options--callback-) が追加されました。本文全体を指定されたテキストに置換します。</span><span class="sxs-lookup"><span data-stu-id="cebe8-116">Added [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#setasync-data--options--callback-): Replaces the entire body with the specified text.</span></span>
- <span data-ttu-id="cebe8-p102">[Event](/javascript/api/office/office.addincommands.event) オブジェクトが追加されました。パラメーターとして、Outlook アドインの UI を使用しないコマンド関数に渡されます。処理の完了を通知するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="cebe8-p102">Added [Event](/javascript/api/office/office.addincommands.event) object: Passed as a parameter to UI-less command functions in an Outlook add-in. Used to signal completion of processing.</span></span>
- <span data-ttu-id="cebe8-119">[Office.context.mailbox.item.close](office.context.mailbox.item.md#methods) が追加されました。作成中の現在のアイテムを閉じます。</span><span class="sxs-lookup"><span data-stu-id="cebe8-119">Added [Office.context.mailbox.item.close](office.context.mailbox.item.md#methods): Closes the current item that is being composed.</span></span>
- <span data-ttu-id="cebe8-120">[Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#methods) が追加されました。アイテムを非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="cebe8-120">Added [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#methods): Asynchronously saves an item.</span></span>
- <span data-ttu-id="cebe8-121">[Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#properties) が追加されました。アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="cebe8-121">Added [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#properties): Gets the notification messages for an item.</span></span>
- <span data-ttu-id="cebe8-122">[Office.context.mailbox.convertToEwsId](office.context.mailbox.md#methods) が追加されました。REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="cebe8-122">Added [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#methods): Converts an item ID formatted for REST into EWS format.</span></span>
- <span data-ttu-id="cebe8-123">[Office.context.mailbox.convertToRestId](office.context.mailbox.md#methods) が追加されました。EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="cebe8-123">Added [Office.context.mailbox.convertToRestId](office.context.mailbox.md#methods): Converts an item ID formatted for EWS into REST format.</span></span>
- <span data-ttu-id="cebe8-124">[Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3) が追加されました。予定またはメッセージの通知メッセージの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="cebe8-124">Added [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3): Specifies the notification message type for an appointment or message.</span></span>
- <span data-ttu-id="cebe8-125">[Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3) が追加されました。REST 形式のアイテム ID に対応する REST API のバージョンを指定します。</span><span class="sxs-lookup"><span data-stu-id="cebe8-125">Added [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3): Specifies the version of the REST API that corresponds to a REST-formatted item ID.</span></span>
- <span data-ttu-id="cebe8-126">[NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3) オブジェクトが追加されました。Outlook アドインの通知メッセージにアクセスするメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="cebe8-126">Added [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3) object: Provides methods for accessing notification messages in an Outlook add-in.</span></span>
- <span data-ttu-id="cebe8-127">[NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3) 型を追加しました。`NotificationMessages.getAllAsync` メソッドによって返されます。</span><span class="sxs-lookup"><span data-stu-id="cebe8-127">Added [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3) type: Returned by the `NotificationMessages.getAllAsync` method.</span></span>

## <a name="see-also"></a><span data-ttu-id="cebe8-128">関連項目</span><span class="sxs-lookup"><span data-stu-id="cebe8-128">See also</span></span>

- [<span data-ttu-id="cebe8-129">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="cebe8-129">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="cebe8-130">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="cebe8-130">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="cebe8-131">概要</span><span class="sxs-lookup"><span data-stu-id="cebe8-131">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="cebe8-132">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="cebe8-132">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
