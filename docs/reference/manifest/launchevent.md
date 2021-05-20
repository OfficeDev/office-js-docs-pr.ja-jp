---
title: マニフェスト ファイル内の起動イベント (プレビュー)
description: LaunchEvent 要素は、サポートされているイベントに基づいてアクティブ化するようにアドインを構成します。
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 7283e9aba9ca57793019ffe027a7f4d6e3243aa8
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555312"
---
# <a name="launchevent-element-preview"></a><span data-ttu-id="98662-103">起動イベント要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="98662-103">LaunchEvent element (preview)</span></span>

<span data-ttu-id="98662-104">サポートされているイベントに基づいてアクティブ化するようにアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="98662-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="98662-105">要素の子 [`<LaunchEvents>`](launchevents.md) 。</span><span class="sxs-lookup"><span data-stu-id="98662-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="98662-106">詳細については、「イベント[ベースのアクティブ化用にOutlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="98662-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="98662-107">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="98662-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="98662-108">イベントベースのアクティブ化は現在[プレビュー段階にあり](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)、web 上およびWindowsでOutlookでのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="98662-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and on Windows.</span></span> <span data-ttu-id="98662-109">詳細については、「 [イベント ベースのアクティブ化機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="98662-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="98662-110">構文</span><span class="sxs-lookup"><span data-stu-id="98662-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="98662-111">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="98662-111">Contained in</span></span>

- [<span data-ttu-id="98662-112">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="98662-112">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="98662-113">属性</span><span class="sxs-lookup"><span data-stu-id="98662-113">Attributes</span></span>

|  <span data-ttu-id="98662-114">属性</span><span class="sxs-lookup"><span data-stu-id="98662-114">Attribute</span></span>  |  <span data-ttu-id="98662-115">必須</span><span class="sxs-lookup"><span data-stu-id="98662-115">Required</span></span>  |  <span data-ttu-id="98662-116">説明</span><span class="sxs-lookup"><span data-stu-id="98662-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="98662-117">**Type**</span><span class="sxs-lookup"><span data-stu-id="98662-117">**Type**</span></span>  |  <span data-ttu-id="98662-118">はい</span><span class="sxs-lookup"><span data-stu-id="98662-118">Yes</span></span>  | <span data-ttu-id="98662-119">サポートされているイベントの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="98662-119">Specifies a supported event type.</span></span> <span data-ttu-id="98662-120">サポートされている種類のセットについては、「 [イベントベースのアクティブ化機能をプレビューする方法](../../outlook/autolaunch.md#supported-events)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="98662-120">For the set of supported types, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#supported-events).</span></span> |
|  <span data-ttu-id="98662-121">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="98662-121">**FunctionName**</span></span>  |  <span data-ttu-id="98662-122">はい</span><span class="sxs-lookup"><span data-stu-id="98662-122">Yes</span></span>  | <span data-ttu-id="98662-123">属性で指定されたイベントを処理する JavaScript 関数の名前を指定します `Type` 。</span><span class="sxs-lookup"><span data-stu-id="98662-123">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="98662-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="98662-124">See also</span></span>

- [<span data-ttu-id="98662-125">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="98662-125">LaunchEvents</span></span>](launchevents.md)
