---
title: Office アドインでサポートされていない Window オブジェクト
description: この記事では、Office アドインでは動作しない window ランタイムオブジェクトの一部について説明します。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: d2560748841bd1e2a7708b25a8e51133563d1534
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160506"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a><span data-ttu-id="6ea02-103">Office アドインでサポートされていない Window オブジェクト</span><span class="sxs-lookup"><span data-stu-id="6ea02-103">Window objects that are unsupported in Office Add-ins</span></span>

<span data-ttu-id="6ea02-104">Windows および Office の一部のバージョンでは、アドインは Internet Explorer 11 ランタイムで実行されます。</span><span class="sxs-lookup"><span data-stu-id="6ea02-104">For some versions of Windows and Office, add-ins run in an Internet Explorer 11 runtime.</span></span> <span data-ttu-id="6ea02-105">(詳細については、「 [Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください)。グローバルオブジェクトの一部のプロパティまたはサブプロパティは、 `window` Internet Explorer 11 ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6ea02-105">(For details, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Some properties or subproperties of the global `window` object are not supported in Internet Explorer 11.</span></span> <span data-ttu-id="6ea02-106">アドインで使用されているブラウザーに関係なく、すべてのユーザーに一貫した機能を提供するために、これらのプロパティはアドインで無効になっています。</span><span class="sxs-lookup"><span data-stu-id="6ea02-106">These properties are disabled in add-ins to ensure that your add-in provides a consistent experience to all users, regardless of which browser the add-in is using.</span></span> <span data-ttu-id="6ea02-107">これは、AngularJS が適切に読み込まれるのにも役に立ちます。</span><span class="sxs-lookup"><span data-stu-id="6ea02-107">This also helps AngularJS load properly.</span></span>

<span data-ttu-id="6ea02-108">無効にされたプロパティの一覧を次に示します。</span><span class="sxs-lookup"><span data-stu-id="6ea02-108">The following is a list of the disabled properties.</span></span> <span data-ttu-id="6ea02-109">リストは処理中です。</span><span class="sxs-lookup"><span data-stu-id="6ea02-109">The list is a work in progress.</span></span> <span data-ttu-id="6ea02-110">アドインで機能しない他のプロパティが見つかった場合は、 `window` 次のフィードバックツールを使用してご確認ください。</span><span class="sxs-lookup"><span data-stu-id="6ea02-110">If you discover additional `window` properties that do not work in add-ins, please use the feedback tool below to tell us.</span></span>

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a><span data-ttu-id="6ea02-111">関連項目</span><span class="sxs-lookup"><span data-stu-id="6ea02-111">See also</span></span>

- [<span data-ttu-id="6ea02-112">Office アドインによって使用されるブラウザー</span><span class="sxs-lookup"><span data-stu-id="6ea02-112">Browsers used by Office Add-ins</span></span>](../concepts/browsers-used-by-office-web-add-ins.md)