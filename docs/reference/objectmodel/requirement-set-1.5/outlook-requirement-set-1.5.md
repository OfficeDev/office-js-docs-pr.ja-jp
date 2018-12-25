---
title: Outlook アドイン API 要件セット 1.5
description: ''
ms.date: 11/14/2018
ms.openlocfilehash: dc6432c3e55ed75c120c2872233ca0f275010e73
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433930"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="c5687-102">Outlook アドイン API 要件セット 1.5</span><span class="sxs-lookup"><span data-stu-id="c5687-102">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="c5687-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c5687-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="c5687-104">このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="c5687-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="c5687-105">1.5 の新機能</span><span class="sxs-lookup"><span data-stu-id="c5687-105">What's new in 1.5?</span></span>

<span data-ttu-id="c5687-p101">要件セット 1.5 には、[要件セット 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) のすべての機能が含まれています。次の機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c5687-p101">Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.</span></span>

- <span data-ttu-id="c5687-108">[ピン留め可能な作業ウィンドウ](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane)のサポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c5687-108">Added support for [pinnable task panes](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span></span>
- <span data-ttu-id="c5687-109">[REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) の呼び出しのサポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c5687-109">Added support for calling [REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>
- <span data-ttu-id="c5687-110">インラインで添付ファイルにマークを付ける機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c5687-110">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="c5687-111">作業ウィンドウまたはダイアログを閉じる機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c5687-111">Added ability to close a taskpane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="c5687-112">変更ログ</span><span class="sxs-lookup"><span data-stu-id="c5687-112">Change log</span></span>

- <span data-ttu-id="c5687-113">[Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback) が追加されました。サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="c5687-113">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="c5687-114">[Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-handler-options-callback) が追加されました。サポートされているイベントのイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="c5687-114">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-handler-options-callback): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="c5687-115">[Office.EventType](office.md#eventtype-string) が追加されました。イベント ハンドラーに関連付けられているイベントを指定し、ItemChanged イベントのサポートが含まれるようになります。</span><span class="sxs-lookup"><span data-stu-id="c5687-115">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="c5687-116">[Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string) が追加されました。この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="c5687-116">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="c5687-p102">[Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback) が変更されました。このメソッドの新しい署名付きの新しいバージョン (`getCallbackTokenAsync([options], callback)`) が追加されました。元のバージョンは引き続き使用でき、変更されていません。</span><span class="sxs-lookup"><span data-stu-id="c5687-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="c5687-119">[Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c5687-119">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--):</span></span>
- <span data-ttu-id="c5687-120">[Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) が変更されました。`isInline` と呼ばれる `options` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="c5687-120">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="c5687-121">[Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata) が変更されました。`isInline` と呼ばれる `formData.attachments` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="c5687-121">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="c5687-122">[Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata) が変更されました。`isInline` と呼ばれる `formData.attachments` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="c5687-122">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="c5687-123">関連項目</span><span class="sxs-lookup"><span data-stu-id="c5687-123">See also</span></span>

- [<span data-ttu-id="c5687-124">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="c5687-124">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="c5687-125">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="c5687-125">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="c5687-126">作業の開始</span><span class="sxs-lookup"><span data-stu-id="c5687-126">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)