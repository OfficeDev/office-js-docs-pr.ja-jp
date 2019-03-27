---
title: Outlook アドイン API 要件セット 1.3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 276096870b128896e987bcb303b4cccdb77e0e50
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871277"
---
# <a name="outlook-add-in-api-requirement-set-13"></a><span data-ttu-id="e1e2d-102">Outlook アドイン API 要件セット 1.3</span><span class="sxs-lookup"><span data-stu-id="e1e2d-102">Outlook add-in API requirement set 1.3</span></span>

<span data-ttu-id="e1e2d-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e2d-104">このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-13"></a><span data-ttu-id="e1e2d-105">1.3 の新機能</span><span class="sxs-lookup"><span data-stu-id="e1e2d-105">What's new in 1.3?</span></span>

<span data-ttu-id="e1e2d-p101">要件セット 1.3 には、[要件セット 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) のすべての機能が含まれています。次の機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-p101">Requirement set 1.3 includes all of the features of [Requirement set 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). It added the following features.</span></span>

- <span data-ttu-id="e1e2d-108">[アドイン コマンド](/outlook/add-ins/add-in-commands-for-outlook)のサポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-108">Added support for [add-in commands](/outlook/add-ins/add-in-commands-for-outlook).</span></span>
- <span data-ttu-id="e1e2d-109">作成中のアイテムを保存または閉じる機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-109">Added ability to save or close an item being composed.</span></span>
- <span data-ttu-id="e1e2d-110">アドインで本文全体を取得または設定できるようにする [Body](/javascript/api/outlook_1_3/office.body) オブジェクトが強化されました。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-110">Enhanced [Body](/javascript/api/outlook_1_3/office.body) object to allow addins to get or set the entire body.</span></span>
- <span data-ttu-id="e1e2d-111">EWS 形式と REST 形式間で ID を変換する変換メソッドが追加されました。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-111">Added conversion methods to convert IDs between EWS and REST formats.</span></span>
- <span data-ttu-id="e1e2d-112">アイテム上にある情報バーに通知メッセージを追加する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-112">Added ability to add notification messages to the info bar on items.</span></span>

### <a name="change-log"></a><span data-ttu-id="e1e2d-113">変更ログ</span><span class="sxs-lookup"><span data-stu-id="e1e2d-113">Change log</span></span>

- <span data-ttu-id="e1e2d-114">[Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) が追加されました。現在の本文を指定された形式で返します。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-114">Added [Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-): Returns the current body in a specified format.</span></span>
- <span data-ttu-id="e1e2d-115">[Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-) が追加されました。本文全体を指定されたテキストに置換します。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-115">Added [Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-): Replaces the entire body with the specified text.</span></span>
- <span data-ttu-id="e1e2d-116">[Office.context.officeTheme](office.context.md#officetheme-object) が追加されました。Office テーマの色にアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-116">Added [Office.context.officeTheme](office.context.md#officetheme-object): Provides access to the Office theme colors.</span></span>
- <span data-ttu-id="e1e2d-p102">[Event](/javascript/api/office/office.addincommands.event) オブジェクトが追加されました。パラメーターとして、Outlook アドインの UI を使用しないコマンド関数に渡されます。処理の完了を通知するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-p102">Added [Event](/javascript/api/office/office.addincommands.event) object: Passed as a parameter to UI-less command functions in an Outlook add-in. Used to signal completion of processing.</span></span>
- <span data-ttu-id="e1e2d-119">[Office.context.mailbox.item.close](office.context.mailbox.item.md#close) が追加されました。作成中の現在のアイテムを閉じます。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-119">Added [Office.context.mailbox.item.close](office.context.mailbox.item.md#close): Closes the current item that is being composed.</span></span>
- <span data-ttu-id="e1e2d-120">[Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback) が追加されました。アイテムを非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-120">Added [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback): Asynchronously saves an item.</span></span>
- <span data-ttu-id="e1e2d-121">[Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessages) が追加されました。アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-121">Added [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessages): Gets the notification messages for an item.</span></span>
- <span data-ttu-id="e1e2d-122">[Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string) が追加されました。REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-122">Added [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string): Converts an item ID formatted for REST into EWS format.</span></span>
- <span data-ttu-id="e1e2d-123">[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) が追加されました。EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-123">Added [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string): Converts an item ID formatted for EWS into REST format.</span></span>
- <span data-ttu-id="e1e2d-124">[Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype) が追加されました。予定またはメッセージの通知メッセージの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-124">Added [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype): Specifies the notification message type for an appointment or message.</span></span>
- <span data-ttu-id="e1e2d-125">[Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion) が追加されました。REST 形式のアイテム ID に対応する REST API のバージョンを指定します。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-125">Added [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion): Specifies the version of the REST API that corresponds to a REST-formatted item ID.</span></span>
- <span data-ttu-id="e1e2d-126">[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages) オブジェクトが追加されました。Outlook アドインの通知メッセージにアクセスするメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-126">Added [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages) object: Provides methods for accessing notification messages in an Outlook add-in.</span></span>
- <span data-ttu-id="e1e2d-127">[NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails) 型を追加しました。`NotificationMessages.getAllAsync` メソッドによって返されます。</span><span class="sxs-lookup"><span data-stu-id="e1e2d-127">Added [NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails) type: Returned by the `NotificationMessages.getAllAsync` method.</span></span>

## <a name="see-also"></a><span data-ttu-id="e1e2d-128">関連項目</span><span class="sxs-lookup"><span data-stu-id="e1e2d-128">See also</span></span>

- [<span data-ttu-id="e1e2d-129">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="e1e2d-129">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="e1e2d-130">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="e1e2d-130">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="e1e2d-131">作業の開始</span><span class="sxs-lookup"><span data-stu-id="e1e2d-131">Get started</span></span>](/outlook/add-ins/quick-start)
