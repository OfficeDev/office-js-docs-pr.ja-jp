---
title: マニフェスト ファイル内の LaunchEvents (プレビュー)
description: LaunchEvents 要素は、サポートされているイベントに基づいてアクティブ化するアドインを構成します。
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 9df059879018d79a61f1c900888c8d197e0b9880
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104813"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="6e8b2-103">LaunchEvents 要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="6e8b2-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="6e8b2-104">サポートされているイベントに基づいてアクティブ化するアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="6e8b2-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="6e8b2-105">要素の [`<ExtensionPoint>`](extensionpoint.md) 子。</span><span class="sxs-lookup"><span data-stu-id="6e8b2-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="6e8b2-106">詳細については、「イベント ベースのアクティブ [化用に Outlook アドインを構成する」を参照してください](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="6e8b2-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="6e8b2-107">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="6e8b2-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6e8b2-108">イベント ベースのアクティブ化は現在 [プレビュー中であり](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 、Outlook on the web および Windows でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="6e8b2-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and Windows.</span></span> <span data-ttu-id="6e8b2-109">詳細については、イベント ベースの [アクティブ化機能をプレビューする方法を参照してください](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="6e8b2-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="6e8b2-110">構文</span><span class="sxs-lookup"><span data-stu-id="6e8b2-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="6e8b2-111">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="6e8b2-111">Contained in</span></span>

<span data-ttu-id="6e8b2-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** メール アドイン)</span><span class="sxs-lookup"><span data-stu-id="6e8b2-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="6e8b2-113">子要素</span><span class="sxs-lookup"><span data-stu-id="6e8b2-113">Child elements</span></span>

|  <span data-ttu-id="6e8b2-114">要素</span><span class="sxs-lookup"><span data-stu-id="6e8b2-114">Element</span></span> |  <span data-ttu-id="6e8b2-115">必須</span><span class="sxs-lookup"><span data-stu-id="6e8b2-115">Required</span></span>  |  <span data-ttu-id="6e8b2-116">説明</span><span class="sxs-lookup"><span data-stu-id="6e8b2-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="6e8b2-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="6e8b2-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="6e8b2-118">はい</span><span class="sxs-lookup"><span data-stu-id="6e8b2-118">Yes</span></span> |  <span data-ttu-id="6e8b2-119">アドインのアクティブ化のために、サポートされているイベントを JavaScript ファイル内の関数にマップします。</span><span class="sxs-lookup"><span data-stu-id="6e8b2-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="6e8b2-120">関連項目</span><span class="sxs-lookup"><span data-stu-id="6e8b2-120">See also</span></span>

- [<span data-ttu-id="6e8b2-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="6e8b2-121">LaunchEvent</span></span>](launchevent.md)
