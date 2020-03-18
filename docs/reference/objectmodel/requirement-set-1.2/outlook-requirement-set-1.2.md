---
title: Outlook アドイン API 要件セット 1.2
description: Outlook アドインおよび Office JavaScript Api for the Mailbox API 1.2 の一部として導入された機能と Api。
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: e1605bb2a0d8cc7de0562833cf9cafc9fd932ad4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717783"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="94b85-103">Outlook アドイン API 要件セット 1.2</span><span class="sxs-lookup"><span data-stu-id="94b85-103">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="94b85-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="94b85-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="94b85-105">このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="94b85-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-12"></a><span data-ttu-id="94b85-106">1.2 の新機能</span><span class="sxs-lookup"><span data-stu-id="94b85-106">What's new in 1.2?</span></span>

<span data-ttu-id="94b85-p101">要件セット 1.2 には、[要件セット 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) のすべての機能が含まれています。アドインを使用して、メッセージの件名または本文内のいずれかで、ユーザーのカーソル位置にテキストを挿入する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="94b85-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="94b85-109">変更ログ</span><span class="sxs-lookup"><span data-stu-id="94b85-109">Change log</span></span>

- <span data-ttu-id="94b85-110">[Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods) が追加されました: メッセージの件名または本文から、選択されたデータを非同期的に返します。</span><span class="sxs-lookup"><span data-stu-id="94b85-110">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="94b85-111">[Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods) が追加されました。メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="94b85-111">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="94b85-112">[Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) が変更されました。`attachments` パラメーターに `formData` プロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="94b85-112">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="94b85-113">[Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) が変更されました。`formData` パラメーターに `attachments` プロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="94b85-113">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="94b85-114">関連項目</span><span class="sxs-lookup"><span data-stu-id="94b85-114">See also</span></span>

- [<span data-ttu-id="94b85-115">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="94b85-115">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="94b85-116">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="94b85-116">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="94b85-117">概要</span><span class="sxs-lookup"><span data-stu-id="94b85-117">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="94b85-118">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="94b85-118">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
