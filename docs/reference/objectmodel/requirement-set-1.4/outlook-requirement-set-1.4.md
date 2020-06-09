---
title: Outlook アドイン API 要件セット 1.4
description: Outlook アドインおよび Office JavaScript Api for the Mailbox API 1.4 の一部として導入された機能と Api。
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 6154acc357dd0e0e663d658c8de2d54b641e080a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609819"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="e1b3a-103">Outlook アドイン API 要件セット 1.4</span><span class="sxs-lookup"><span data-stu-id="e1b3a-103">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="e1b3a-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="e1b3a-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e1b3a-105">このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="e1b3a-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="e1b3a-106">1.4 の新機能</span><span class="sxs-lookup"><span data-stu-id="e1b3a-106">What's new in 1.4?</span></span>

<span data-ttu-id="e1b3a-p101">要件セット 1.4 には、[要件セット 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) のすべての機能が含まれています。`Office.ui` 名前空間へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="e1b3a-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="e1b3a-109">変更ログ</span><span class="sxs-lookup"><span data-stu-id="e1b3a-109">Change log</span></span>

- <span data-ttu-id="e1b3a-110">[Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) が追加されました。Office ホストでダイアログ ボックスを表示します。</span><span class="sxs-lookup"><span data-stu-id="e1b3a-110">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="e1b3a-111">[Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-) が追加されました。メッセージをダイアログ ボックスからその親/オープナー ページに配信します。</span><span class="sxs-lookup"><span data-stu-id="e1b3a-111">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="e1b3a-112">[Dialog](/javascript/api/office/office.dialog) オブジェクトが追加されました。このオブジェクトは、[`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) メソッドが呼び出されたときに返されます。</span><span class="sxs-lookup"><span data-stu-id="e1b3a-112">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="e1b3a-113">関連項目</span><span class="sxs-lookup"><span data-stu-id="e1b3a-113">See also</span></span>

- [<span data-ttu-id="e1b3a-114">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="e1b3a-114">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="e1b3a-115">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="e1b3a-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="e1b3a-116">概要</span><span class="sxs-lookup"><span data-stu-id="e1b3a-116">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="e1b3a-117">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="e1b3a-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
