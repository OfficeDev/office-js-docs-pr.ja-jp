---
title: Outlook アドイン API 要件セット 1.2
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 78ad7c02985b1a61d6a187d7beaff52bef9411b6
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42597062"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="99bfd-102">Outlook アドイン API 要件セット 1.2</span><span class="sxs-lookup"><span data-stu-id="99bfd-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="99bfd-103">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="99bfd-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="99bfd-104">このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="99bfd-104">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-12"></a><span data-ttu-id="99bfd-105">1.2 の新機能</span><span class="sxs-lookup"><span data-stu-id="99bfd-105">What's new in 1.2?</span></span>

<span data-ttu-id="99bfd-p101">要件セット 1.2 には、[要件セット 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) のすべての機能が含まれています。アドインを使用して、メッセージの件名または本文内のいずれかで、ユーザーのカーソル位置にテキストを挿入する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="99bfd-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="99bfd-108">変更ログ</span><span class="sxs-lookup"><span data-stu-id="99bfd-108">Change log</span></span>

- <span data-ttu-id="99bfd-109">[Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods) が追加されました: メッセージの件名または本文から、選択されたデータを非同期的に返します。</span><span class="sxs-lookup"><span data-stu-id="99bfd-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="99bfd-110">[Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods) が追加されました。メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="99bfd-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="99bfd-111">[Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) が変更されました。`attachments` パラメーターに `formData` プロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="99bfd-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="99bfd-112">[Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) が変更されました。`formData` パラメーターに `attachments` プロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="99bfd-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="99bfd-113">関連項目</span><span class="sxs-lookup"><span data-stu-id="99bfd-113">See also</span></span>

- [<span data-ttu-id="99bfd-114">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="99bfd-114">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="99bfd-115">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="99bfd-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="99bfd-116">概要</span><span class="sxs-lookup"><span data-stu-id="99bfd-116">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="99bfd-117">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="99bfd-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
