---
title: マニフェスト ファイルの LaunchEvents
description: LaunchEvents 要素は、サポートされているイベントに基づいてアクティブ化するアドインを構成します。
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 16d721ca6d9402d2bd5d19787707e146358044f0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590918"
---
# <a name="launchevents-element"></a><span data-ttu-id="2fc87-103">LaunchEvents 要素</span><span class="sxs-lookup"><span data-stu-id="2fc87-103">LaunchEvents element</span></span>

<span data-ttu-id="2fc87-104">サポートされているイベントに基づいてアクティブ化するアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="2fc87-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="2fc87-105">要素の [`<ExtensionPoint>`](extensionpoint.md) 子。</span><span class="sxs-lookup"><span data-stu-id="2fc87-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="2fc87-106">詳細については、「イベント ベース[のアクティブ化Outlookアドインを構成する」を参照してください](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="2fc87-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="2fc87-107">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="2fc87-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2fc87-108">構文</span><span class="sxs-lookup"><span data-stu-id="2fc87-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="2fc87-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="2fc87-109">Contained in</span></span>

<span data-ttu-id="2fc87-110">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** メール アドイン)</span><span class="sxs-lookup"><span data-stu-id="2fc87-110">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="2fc87-111">子要素</span><span class="sxs-lookup"><span data-stu-id="2fc87-111">Child elements</span></span>

|  <span data-ttu-id="2fc87-112">要素</span><span class="sxs-lookup"><span data-stu-id="2fc87-112">Element</span></span> |  <span data-ttu-id="2fc87-113">必須</span><span class="sxs-lookup"><span data-stu-id="2fc87-113">Required</span></span>  |  <span data-ttu-id="2fc87-114">説明</span><span class="sxs-lookup"><span data-stu-id="2fc87-114">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="2fc87-115">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="2fc87-115">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="2fc87-116">必要</span><span class="sxs-lookup"><span data-stu-id="2fc87-116">Yes</span></span> |  <span data-ttu-id="2fc87-117">サポートされているイベントを JavaScript ファイル内の関数にマップして、アドインのアクティブ化を行います。</span><span class="sxs-lookup"><span data-stu-id="2fc87-117">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="2fc87-118">関連項目</span><span class="sxs-lookup"><span data-stu-id="2fc87-118">See also</span></span>

- [<span data-ttu-id="2fc87-119">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="2fc87-119">LaunchEvent</span></span>](launchevent.md)
