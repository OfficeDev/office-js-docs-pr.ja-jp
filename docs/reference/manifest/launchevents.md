---
title: マニフェスト ファイル内の LaunchEvents (プレビュー)
description: LaunchEvents 要素は、サポートされているイベントに基づいてアクティブ化するアドインを構成します。
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 59c52aa3f60e69e2bdda84718c6123f02942fedc
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237981"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="2d09a-103">LaunchEvents 要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="2d09a-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="2d09a-104">サポートされているイベントに基づいてアクティブ化するアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="2d09a-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="2d09a-105">要素の [`<ExtensionPoint>`](extensionpoint.md) 子。</span><span class="sxs-lookup"><span data-stu-id="2d09a-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="2d09a-106">詳細については、「イベント ベースのアクティブ [化用に Outlook アドインを構成する」を参照してください](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="2d09a-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="2d09a-107">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="2d09a-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2d09a-108">イベント ベースのアクティブ化は現在 [プレビュー中](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) で、Outlook on the web および Windows でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="2d09a-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and on Windows.</span></span> <span data-ttu-id="2d09a-109">詳細については、イベント ベースの [アクティブ化機能をプレビューする方法を参照してください](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="2d09a-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="2d09a-110">構文</span><span class="sxs-lookup"><span data-stu-id="2d09a-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="2d09a-111">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="2d09a-111">Contained in</span></span>

<span data-ttu-id="2d09a-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** メール アドイン)</span><span class="sxs-lookup"><span data-stu-id="2d09a-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="2d09a-113">子要素</span><span class="sxs-lookup"><span data-stu-id="2d09a-113">Child elements</span></span>

|  <span data-ttu-id="2d09a-114">要素</span><span class="sxs-lookup"><span data-stu-id="2d09a-114">Element</span></span> |  <span data-ttu-id="2d09a-115">必須</span><span class="sxs-lookup"><span data-stu-id="2d09a-115">Required</span></span>  |  <span data-ttu-id="2d09a-116">説明</span><span class="sxs-lookup"><span data-stu-id="2d09a-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="2d09a-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="2d09a-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="2d09a-118">はい</span><span class="sxs-lookup"><span data-stu-id="2d09a-118">Yes</span></span> |  <span data-ttu-id="2d09a-119">アドインのアクティブ化のために、サポートされているイベントを JavaScript ファイル内の関数にマップします。</span><span class="sxs-lookup"><span data-stu-id="2d09a-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="2d09a-120">関連項目</span><span class="sxs-lookup"><span data-stu-id="2d09a-120">See also</span></span>

- [<span data-ttu-id="2d09a-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="2d09a-121">LaunchEvent</span></span>](launchevent.md)
