---
title: マニフェストファイル内の LaunchEvents (プレビュー)
description: LaunchEvents 要素は、サポートされているイベントに基づいてアクティブになるようにアドインを構成します。
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 92416f8c646326410a8cd9ee7831e17a5c5f1ffc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611772"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="03fe7-103">LaunchEvents 要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="03fe7-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="03fe7-104">サポートされているイベントに基づいて、アドインをアクティブにするように構成します。</span><span class="sxs-lookup"><span data-stu-id="03fe7-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="03fe7-105">要素の子 [`<ExtensionPoint>`](extensionpoint.md) 。</span><span class="sxs-lookup"><span data-stu-id="03fe7-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="03fe7-106">詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="03fe7-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="03fe7-107">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="03fe7-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="03fe7-108">イベントベースのライセンス認証は現在[プレビュー段階で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)あり、web 上の Outlook でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="03fe7-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="03fe7-109">詳細については、「[イベントベースのライセンス認証機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="03fe7-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="03fe7-110">構文</span><span class="sxs-lookup"><span data-stu-id="03fe7-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="03fe7-111">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="03fe7-111">Contained in</span></span>

<span data-ttu-id="03fe7-112">[Extensionpoint](extensionpoint.md) (**launchevent**メールアドイン)</span><span class="sxs-lookup"><span data-stu-id="03fe7-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="03fe7-113">子要素</span><span class="sxs-lookup"><span data-stu-id="03fe7-113">Child elements</span></span>

|  <span data-ttu-id="03fe7-114">要素</span><span class="sxs-lookup"><span data-stu-id="03fe7-114">Element</span></span> |  <span data-ttu-id="03fe7-115">必須</span><span class="sxs-lookup"><span data-stu-id="03fe7-115">Required</span></span>  |  <span data-ttu-id="03fe7-116">説明</span><span class="sxs-lookup"><span data-stu-id="03fe7-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="03fe7-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="03fe7-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="03fe7-118">はい</span><span class="sxs-lookup"><span data-stu-id="03fe7-118">Yes</span></span> |  <span data-ttu-id="03fe7-119">アドインをアクティブ化するために、JavaScript ファイルの関数にサポートされているイベントをマップします。</span><span class="sxs-lookup"><span data-stu-id="03fe7-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="03fe7-120">関連項目</span><span class="sxs-lookup"><span data-stu-id="03fe7-120">See also</span></span>

- [<span data-ttu-id="03fe7-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="03fe7-121">LaunchEvent</span></span>](launchevent.md)
