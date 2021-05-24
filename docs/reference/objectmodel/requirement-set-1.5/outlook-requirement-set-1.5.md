---
title: Outlook アドイン API 要件セット 1.5
description: メールボックス API 1.5 の一部Outlook JavaScript API および Office JavaScript API 用に導入された機能と API。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7d780538a77f54db6f1234a6d29a3bcdea9533b0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590842"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="b7250-103">Outlook アドイン API 要件セット 1.5</span><span class="sxs-lookup"><span data-stu-id="b7250-103">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="b7250-104">Office Outlook JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="b7250-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b7250-105">このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="b7250-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="b7250-106">1.5 の新機能</span><span class="sxs-lookup"><span data-stu-id="b7250-106">What's new in 1.5?</span></span>

<span data-ttu-id="b7250-107">要件セット 1.5 には、要件セット [1.4 のすべての機能が含まれています](../requirement-set-1.4/outlook-requirement-set-1.4.md)。</span><span class="sxs-lookup"><span data-stu-id="b7250-107">Requirement set 1.5 includes all of the features of [requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md).</span></span> <span data-ttu-id="b7250-108">次の機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="b7250-108">It added the following features.</span></span>

- <span data-ttu-id="b7250-109">[ピン留め可能な作業ウィンドウ](../../../outlook/pinnable-taskpane.md)のサポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="b7250-109">Added support for [pinnable task panes](../../../outlook/pinnable-taskpane.md).</span></span>
- <span data-ttu-id="b7250-110">[REST API](../../../outlook/use-rest-api.md) の呼び出しのサポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="b7250-110">Added support for calling [REST APIs](../../../outlook/use-rest-api.md).</span></span>
- <span data-ttu-id="b7250-111">インラインで添付ファイルにマークを付ける機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="b7250-111">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="b7250-112">作業ウィンドウまたはダイアログを閉じる機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="b7250-112">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="b7250-113">変更ログ</span><span class="sxs-lookup"><span data-stu-id="b7250-113">Change log</span></span>

- <span data-ttu-id="b7250-114">[Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods) が追加されました。サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="b7250-114">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="b7250-115">[Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="b7250-115">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="b7250-116">[Office.EventType](office.md#eventtype-string) が追加されました。イベント ハンドラーに関連付けられているイベントを指定し、ItemChanged イベントのサポートが含まれるようになります。</span><span class="sxs-lookup"><span data-stu-id="b7250-116">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="b7250-117">[Office.context.mailbox.restUrl](office.context.mailbox.md#properties) が追加されました。この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="b7250-117">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="b7250-p102">[Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods) が変更されました。このメソッドの新しい署名付きの新しいバージョン (`getCallbackTokenAsync([options], callback)`) が追加されました。元のバージョンは引き続き使用でき、変更されていません。</span><span class="sxs-lookup"><span data-stu-id="b7250-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="b7250-120">[Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) が追加されました。</span><span class="sxs-lookup"><span data-stu-id="b7250-120">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="b7250-121">[Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) が変更されました。`isInline` と呼ばれる `options` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="b7250-121">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="b7250-122">[Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) が変更されました。`formData.attachments` と呼ばれる `isInline` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="b7250-122">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="b7250-123">[Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) が変更されました。`isInline` と呼ばれる `formData.attachments` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="b7250-123">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="b7250-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="b7250-124">See also</span></span>

- [<span data-ttu-id="b7250-125">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="b7250-125">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="b7250-126">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="b7250-126">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="b7250-127">概要</span><span class="sxs-lookup"><span data-stu-id="b7250-127">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="b7250-128">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="b7250-128">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
