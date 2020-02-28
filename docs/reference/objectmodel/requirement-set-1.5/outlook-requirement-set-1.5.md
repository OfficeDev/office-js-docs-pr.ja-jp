---
title: Outlook アドイン API 要件セット 1.5
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: e9dda19a4dd9831ca1bb7924558ac99c26baf763
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324793"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="19e68-102">Outlook アドイン API 要件セット 1.5</span><span class="sxs-lookup"><span data-stu-id="19e68-102">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="19e68-103">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="19e68-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="19e68-104">このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="19e68-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="19e68-105">1.5 の新機能</span><span class="sxs-lookup"><span data-stu-id="19e68-105">What's new in 1.5?</span></span>

<span data-ttu-id="19e68-p101">要件セット 1.5 には、[要件セット 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) のすべての機能が含まれています。次の機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="19e68-p101">Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.</span></span>

- <span data-ttu-id="19e68-108">[ピン留め可能な作業ウィンドウ](../../../outlook/pinnable-taskpane.md)のサポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="19e68-108">Added support for [pinnable task panes](../../../outlook/pinnable-taskpane.md).</span></span>
- <span data-ttu-id="19e68-109">[REST API](../../../outlook/use-rest-api.md) の呼び出しのサポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="19e68-109">Added support for calling [REST APIs](../../../outlook/use-rest-api.md).</span></span>
- <span data-ttu-id="19e68-110">インラインで添付ファイルにマークを付ける機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="19e68-110">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="19e68-111">作業ウィンドウまたはダイアログを閉じる機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="19e68-111">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="19e68-112">変更ログ</span><span class="sxs-lookup"><span data-stu-id="19e68-112">Change log</span></span>

- <span data-ttu-id="19e68-113">[Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods) が追加されました。サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="19e68-113">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="19e68-114">追加された、サポートされているイベントの種類のイベントハンドラを削除[し](office.context.mailbox.md#methods)ました。</span><span class="sxs-lookup"><span data-stu-id="19e68-114">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="19e68-115">[Office.EventType](office.md#eventtype-string) が追加されました。イベント ハンドラーに関連付けられているイベントを指定し、ItemChanged イベントのサポートが含まれるようになります。</span><span class="sxs-lookup"><span data-stu-id="19e68-115">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="19e68-116">[Office.context.mailbox.restUrl](office.context.mailbox.md#properties) が追加されました。この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="19e68-116">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="19e68-p102">[Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods) が変更されました。このメソッドの新しい署名付きの新しいバージョン (`getCallbackTokenAsync([options], callback)`) が追加されました。元のバージョンは引き続き使用でき、変更されていません。</span><span class="sxs-lookup"><span data-stu-id="19e68-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="19e68-119">[Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) が追加されました。</span><span class="sxs-lookup"><span data-stu-id="19e68-119">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="19e68-120">[Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) が変更されました。`isInline` と呼ばれる `options` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="19e68-120">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="19e68-121">[Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) が変更されました。`formData.attachments` と呼ばれる `isInline` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="19e68-121">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="19e68-122">[Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) が変更されました。`isInline` と呼ばれる `formData.attachments` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="19e68-122">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="19e68-123">関連項目</span><span class="sxs-lookup"><span data-stu-id="19e68-123">See also</span></span>

- [<span data-ttu-id="19e68-124">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="19e68-124">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="19e68-125">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="19e68-125">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="19e68-126">概要</span><span class="sxs-lookup"><span data-stu-id="19e68-126">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="19e68-127">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="19e68-127">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
