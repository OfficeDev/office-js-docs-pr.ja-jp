---
title: Outlook アドイン API 要件セット 1.4
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: be700af413a041502cddd491f304a693c259da28
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432369"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="c7967-102">Outlook アドイン API 要件セット 1.4</span><span class="sxs-lookup"><span data-stu-id="c7967-102">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="c7967-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c7967-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="c7967-104">このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="c7967-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="c7967-105">1.4 の新機能</span><span class="sxs-lookup"><span data-stu-id="c7967-105">What's new in 1.4?</span></span>

<span data-ttu-id="c7967-p101">要件セット 1.4 には、[要件セット 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) のすべての機能が含まれています。`Office.ui` 名前空間へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c7967-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="c7967-108">変更ログ</span><span class="sxs-lookup"><span data-stu-id="c7967-108">Change log</span></span>

- <span data-ttu-id="c7967-109">[Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) が追加されました。Office ホストでダイアログ ボックスを表示します。</span><span class="sxs-lookup"><span data-stu-id="c7967-109">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="c7967-110">[Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-) が追加されました。メッセージをダイアログ ボックスからその親/オープナー ページに配信します。</span><span class="sxs-lookup"><span data-stu-id="c7967-110">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="c7967-111">[Dialog](/javascript/api/office/office.dialog) オブジェクトが追加されました。このオブジェクトは、[`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) メソッドが呼び出されたときに返されます。</span><span class="sxs-lookup"><span data-stu-id="c7967-111">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="c7967-112">関連項目</span><span class="sxs-lookup"><span data-stu-id="c7967-112">See also</span></span>

- [<span data-ttu-id="c7967-113">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="c7967-113">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="c7967-114">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="c7967-114">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="c7967-115">作業の開始</span><span class="sxs-lookup"><span data-stu-id="c7967-115">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)