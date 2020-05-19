---
title: マニフェストファイルの LaunchEvent (プレビュー)
description: LaunchEvent 要素は、サポートされているイベントに基づいてアクティブになるようにアドインを構成します。
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: a4f5208ec7f735d926c3a878cae34973c3992cf9
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278559"
---
# <a name="launchevent-element-preview"></a><span data-ttu-id="446bd-103">LaunchEvent 要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="446bd-103">LaunchEvent element (preview)</span></span>

<span data-ttu-id="446bd-104">サポートされているイベントに基づいて、アドインをアクティブにするように構成します。</span><span class="sxs-lookup"><span data-stu-id="446bd-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="446bd-105">要素の子 [`<LaunchEvents>`](launchevents.md) 。</span><span class="sxs-lookup"><span data-stu-id="446bd-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="446bd-106">詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="446bd-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="446bd-107">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="446bd-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="446bd-108">イベントベースのライセンス認証は現在[プレビュー段階で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)あり、web 上の Outlook でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="446bd-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="446bd-109">詳細については、「[イベントベースのライセンス認証機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="446bd-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="446bd-110">構文</span><span class="sxs-lookup"><span data-stu-id="446bd-110">Syntax</span></span>

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## <a name="contained-in"></a><span data-ttu-id="446bd-111">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="446bd-111">Contained in</span></span>

- [<span data-ttu-id="446bd-112">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="446bd-112">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="446bd-113">属性</span><span class="sxs-lookup"><span data-stu-id="446bd-113">Attributes</span></span>

|  <span data-ttu-id="446bd-114">属性</span><span class="sxs-lookup"><span data-stu-id="446bd-114">Attribute</span></span>  |  <span data-ttu-id="446bd-115">必須</span><span class="sxs-lookup"><span data-stu-id="446bd-115">Required</span></span>  |  <span data-ttu-id="446bd-116">説明</span><span class="sxs-lookup"><span data-stu-id="446bd-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="446bd-117">**種類**</span><span class="sxs-lookup"><span data-stu-id="446bd-117">**Type**</span></span>  |  <span data-ttu-id="446bd-118">はい</span><span class="sxs-lookup"><span data-stu-id="446bd-118">Yes</span></span>  | <span data-ttu-id="446bd-119">サポートされているイベントの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="446bd-119">Specifies a supported event type.</span></span> <span data-ttu-id="446bd-120">使用できる型は `OnNewMessageCompose` 、および `OnNewAppointmentOrganizer` です。</span><span class="sxs-lookup"><span data-stu-id="446bd-120">Available types are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> |
|  <span data-ttu-id="446bd-121">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="446bd-121">**FunctionName**</span></span>  |  <span data-ttu-id="446bd-122">はい</span><span class="sxs-lookup"><span data-stu-id="446bd-122">Yes</span></span>  | <span data-ttu-id="446bd-123">属性で指定されたイベントを処理する JavaScript 関数の名前を指定し `Type` ます。</span><span class="sxs-lookup"><span data-stu-id="446bd-123">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="446bd-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="446bd-124">See also</span></span>

- [<span data-ttu-id="446bd-125">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="446bd-125">LaunchEvents</span></span>](launchevents.md)
