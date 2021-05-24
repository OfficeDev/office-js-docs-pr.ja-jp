---
title: Outlook アドイン API 要件セット 1.2
description: メールボックス API 1.2 のOutlook JavaScript API および Office JavaScript API 用に導入された機能と API。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: d643f0fdf07c5f22d8d863075b894cfc05b21363
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590401"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="c2e49-103">Outlook アドイン API 要件セット 1.2</span><span class="sxs-lookup"><span data-stu-id="c2e49-103">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="c2e49-104">Office Outlook JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c2e49-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="c2e49-105">このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="c2e49-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-12"></a><span data-ttu-id="c2e49-106">1.2 の新機能</span><span class="sxs-lookup"><span data-stu-id="c2e49-106">What's new in 1.2?</span></span>

<span data-ttu-id="c2e49-107">要件セット 1.2 には、要件セット [1.1 のすべての機能が含まれています](../requirement-set-1.1/outlook-requirement-set-1.1.md)。</span><span class="sxs-lookup"><span data-stu-id="c2e49-107">Requirement set 1.2 includes all of the features of [requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md).</span></span> <span data-ttu-id="c2e49-108">アドインを使用して、メッセージの件名または本文内のいずれかで、ユーザーのカーソル位置にテキストを挿入する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c2e49-108">It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="c2e49-109">変更ログ</span><span class="sxs-lookup"><span data-stu-id="c2e49-109">Change log</span></span>

- <span data-ttu-id="c2e49-110">[Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods) が追加されました: メッセージの件名または本文から、選択されたデータを非同期的に返します。</span><span class="sxs-lookup"><span data-stu-id="c2e49-110">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="c2e49-111">[Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods) が追加されました。メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="c2e49-111">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="c2e49-112">[Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) が変更されました。`attachments` パラメーターに `formData` プロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c2e49-112">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="c2e49-113">[Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) が変更されました。`formData` パラメーターに `attachments` プロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c2e49-113">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="c2e49-114">関連項目</span><span class="sxs-lookup"><span data-stu-id="c2e49-114">See also</span></span>

- [<span data-ttu-id="c2e49-115">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="c2e49-115">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="c2e49-116">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="c2e49-116">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="c2e49-117">概要</span><span class="sxs-lookup"><span data-stu-id="c2e49-117">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="c2e49-118">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="c2e49-118">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
