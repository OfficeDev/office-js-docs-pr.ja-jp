---
title: マニフェスト ファイルの LaunchEvent
description: LaunchEvent 要素は、サポートされているイベントに基づいてアクティブ化するアドインを構成します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: c866a085ed6b7a33c8d7bf02d25e6ec748629e07
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591080"
---
# <a name="launchevent-element"></a><span data-ttu-id="e146d-103">LaunchEvent 要素</span><span class="sxs-lookup"><span data-stu-id="e146d-103">LaunchEvent element</span></span>

<span data-ttu-id="e146d-104">サポートされているイベントに基づいてアクティブ化するアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="e146d-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="e146d-105">要素の [`<LaunchEvents>`](launchevents.md) 子。</span><span class="sxs-lookup"><span data-stu-id="e146d-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="e146d-106">詳細については、「イベント ベース[のアクティブ化Outlookアドインを構成する」を参照してください](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="e146d-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="e146d-107">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="e146d-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e146d-108">構文</span><span class="sxs-lookup"><span data-stu-id="e146d-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="e146d-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="e146d-109">Contained in</span></span>

- [<span data-ttu-id="e146d-110">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="e146d-110">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="e146d-111">属性</span><span class="sxs-lookup"><span data-stu-id="e146d-111">Attributes</span></span>

|  <span data-ttu-id="e146d-112">属性</span><span class="sxs-lookup"><span data-stu-id="e146d-112">Attribute</span></span>  |  <span data-ttu-id="e146d-113">必須</span><span class="sxs-lookup"><span data-stu-id="e146d-113">Required</span></span>  |  <span data-ttu-id="e146d-114">説明</span><span class="sxs-lookup"><span data-stu-id="e146d-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e146d-115">**Type**</span><span class="sxs-lookup"><span data-stu-id="e146d-115">**Type**</span></span>  |  <span data-ttu-id="e146d-116">はい</span><span class="sxs-lookup"><span data-stu-id="e146d-116">Yes</span></span>  | <span data-ttu-id="e146d-117">サポートされているイベントの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="e146d-117">Specifies a supported event type.</span></span> <span data-ttu-id="e146d-118">サポートされている一連の種類については、「イベント ベースのライセンス認証Outlookアドインを構成する[」を参照してください](../../outlook/autolaunch.md#supported-events)。</span><span class="sxs-lookup"><span data-stu-id="e146d-118">For the set of supported types, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md#supported-events).</span></span> |
|  <span data-ttu-id="e146d-119">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="e146d-119">**FunctionName**</span></span>  |  <span data-ttu-id="e146d-120">はい</span><span class="sxs-lookup"><span data-stu-id="e146d-120">Yes</span></span>  | <span data-ttu-id="e146d-121">属性で指定されたイベントを処理する JavaScript 関数の名前を指定 `Type` します。</span><span class="sxs-lookup"><span data-stu-id="e146d-121">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="e146d-122">関連項目</span><span class="sxs-lookup"><span data-stu-id="e146d-122">See also</span></span>

- [<span data-ttu-id="e146d-123">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="e146d-123">LaunchEvents</span></span>](launchevents.md)
