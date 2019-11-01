---
title: Outlook アドイン API 要件セット 1.2
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 898e768dfc1828ba44f29e9da5c4baa61de186cb
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902096"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="35d78-102">Outlook アドイン API 要件セット 1.2</span><span class="sxs-lookup"><span data-stu-id="35d78-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="35d78-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="35d78-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="35d78-104">このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="35d78-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="35d78-105">1.2 の新機能</span><span class="sxs-lookup"><span data-stu-id="35d78-105">What's new in 1.2?</span></span>

<span data-ttu-id="35d78-p101">要件セット 1.2 には、[要件セット 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) のすべての機能が含まれています。アドインを使用して、メッセージの件名または本文内のいずれかで、ユーザーのカーソル位置にテキストを挿入する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="35d78-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="35d78-108">変更ログ</span><span class="sxs-lookup"><span data-stu-id="35d78-108">Change log</span></span>

- <span data-ttu-id="35d78-109">[Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string) が追加されました: メッセージの件名または本文から、選択されたデータを非同期的に返します。</span><span class="sxs-lookup"><span data-stu-id="35d78-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="35d78-110">[Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback) が追加されました。メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="35d78-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="35d78-111">[Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback) が変更されました。`attachments` パラメーターに `formData` プロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="35d78-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="35d78-112">[Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback) が変更されました。`formData` パラメーターに `attachments` プロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="35d78-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="35d78-113">関連項目</span><span class="sxs-lookup"><span data-stu-id="35d78-113">See also</span></span>

- [<span data-ttu-id="35d78-114">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="35d78-114">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="35d78-115">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="35d78-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="35d78-116">概要</span><span class="sxs-lookup"><span data-stu-id="35d78-116">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="35d78-117">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="35d78-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
