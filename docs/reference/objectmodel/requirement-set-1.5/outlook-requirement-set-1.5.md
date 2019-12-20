---
title: Outlook アドイン API 要件セット 1.5
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 1a12156feb7a03e596e521650a757fe7198b4d76
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814746"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="05afb-102">Outlook アドイン API 要件セット 1.5</span><span class="sxs-lookup"><span data-stu-id="05afb-102">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="05afb-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="05afb-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="05afb-104">このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="05afb-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="05afb-105">1.5 の新機能</span><span class="sxs-lookup"><span data-stu-id="05afb-105">What's new in 1.5?</span></span>

<span data-ttu-id="05afb-p101">要件セット 1.5 には、[要件セット 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) のすべての機能が含まれています。次の機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="05afb-p101">Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.</span></span>

- <span data-ttu-id="05afb-108">[ピン留め可能な作業ウィンドウ](/outlook/add-ins/pinnable-taskpane)のサポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="05afb-108">Added support for [pinnable task panes](/outlook/add-ins/pinnable-taskpane).</span></span>
- <span data-ttu-id="05afb-109">[REST API](/outlook/add-ins/use-rest-api) の呼び出しのサポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="05afb-109">Added support for calling [REST APIs](/outlook/add-ins/use-rest-api).</span></span>
- <span data-ttu-id="05afb-110">インラインで添付ファイルにマークを付ける機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="05afb-110">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="05afb-111">作業ウィンドウまたはダイアログを閉じる機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="05afb-111">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="05afb-112">変更ログ</span><span class="sxs-lookup"><span data-stu-id="05afb-112">Change log</span></span>

- <span data-ttu-id="05afb-113">[Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods) が追加されました。サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="05afb-113">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="05afb-114">追加された、サポートされているイベントの種類のイベントハンドラを削除[し](office.context.mailbox.md#methods)ました。</span><span class="sxs-lookup"><span data-stu-id="05afb-114">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="05afb-115">[Office.EventType](office.md#eventtype-string) が追加されました。イベント ハンドラーに関連付けられているイベントを指定し、ItemChanged イベントのサポートが含まれるようになります。</span><span class="sxs-lookup"><span data-stu-id="05afb-115">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="05afb-116">[Office.context.mailbox.restUrl](office.context.mailbox.md#properties) が追加されました。この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="05afb-116">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="05afb-p102">[Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods) が変更されました。このメソッドの新しい署名付きの新しいバージョン (`getCallbackTokenAsync([options], callback)`) が追加されました。元のバージョンは引き続き使用でき、変更されていません。</span><span class="sxs-lookup"><span data-stu-id="05afb-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="05afb-119">[Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) が追加されました。</span><span class="sxs-lookup"><span data-stu-id="05afb-119">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="05afb-120">[Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) が変更されました。`isInline` と呼ばれる `options` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="05afb-120">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="05afb-121">[Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) が変更されました。`formData.attachments` と呼ばれる `isInline` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="05afb-121">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="05afb-122">[Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) が変更されました。`isInline` と呼ばれる `formData.attachments` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="05afb-122">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="05afb-123">関連項目</span><span class="sxs-lookup"><span data-stu-id="05afb-123">See also</span></span>

- [<span data-ttu-id="05afb-124">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="05afb-124">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="05afb-125">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="05afb-125">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="05afb-126">概要</span><span class="sxs-lookup"><span data-stu-id="05afb-126">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="05afb-127">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="05afb-127">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
