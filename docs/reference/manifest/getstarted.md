---
title: マニフェスト ファイルの GetStarted 要素
description: Word、Excel、PowerPoint、およびアドインにアドインがインストールされている場合に表示される吹き出しでPowerPoint情報をOneNote。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a637f3f9031d9f8e09d14f17f2095ca0647c4d50
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348686"
---
# <a name="getstarted-element"></a><span data-ttu-id="4f910-103">GetStarted 要素</span><span class="sxs-lookup"><span data-stu-id="4f910-103">GetStarted element</span></span>

<span data-ttu-id="4f910-104">Word、Excel、PowerPoint、およびアドインにアドインがインストールされている場合に表示される吹き出しでPowerPoint情報をOneNote。</span><span class="sxs-lookup"><span data-stu-id="4f910-104">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote.</span></span> <span data-ttu-id="4f910-105">**GetStarted 要素** は [DesktopFormFactor の子要素です](desktopformfactor.md)。</span><span class="sxs-lookup"><span data-stu-id="4f910-105">The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="4f910-106">子要素</span><span class="sxs-lookup"><span data-stu-id="4f910-106">Child elements</span></span>

| <span data-ttu-id="4f910-107">要素</span><span class="sxs-lookup"><span data-stu-id="4f910-107">Element</span></span>                       | <span data-ttu-id="4f910-108">必須</span><span class="sxs-lookup"><span data-stu-id="4f910-108">Required</span></span> | <span data-ttu-id="4f910-109">説明</span><span class="sxs-lookup"><span data-stu-id="4f910-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="4f910-110">Title</span><span class="sxs-lookup"><span data-stu-id="4f910-110">Title</span></span>](#title)               | <span data-ttu-id="4f910-111">はい</span><span class="sxs-lookup"><span data-stu-id="4f910-111">Yes</span></span>      | <span data-ttu-id="4f910-112">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="4f910-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="4f910-113">説明</span><span class="sxs-lookup"><span data-stu-id="4f910-113">Description</span></span>](#description)   | <span data-ttu-id="4f910-114">はい</span><span class="sxs-lookup"><span data-stu-id="4f910-114">Yes</span></span>      | <span data-ttu-id="4f910-115">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="4f910-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="4f910-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="4f910-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="4f910-117">はい</span><span class="sxs-lookup"><span data-stu-id="4f910-117">Yes</span></span>       | <span data-ttu-id="4f910-118">アドインの詳細を説明するページの URL。</span><span class="sxs-lookup"><span data-stu-id="4f910-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="4f910-119">タイトル</span><span class="sxs-lookup"><span data-stu-id="4f910-119">Title</span></span> 

<span data-ttu-id="4f910-120">必須。</span><span class="sxs-lookup"><span data-stu-id="4f910-120">Required.</span></span> <span data-ttu-id="4f910-121">吹き出しの一番上に使用するタイトル。</span><span class="sxs-lookup"><span data-stu-id="4f910-121">The title used for the top of the callout.</span></span> <span data-ttu-id="4f910-122">**resid 属性** は、[リソース] セクションの **ShortStrings** 要素 [](resources.md)の有効な ID を参照し、32 文字以内で指定できます。</span><span class="sxs-lookup"><span data-stu-id="4f910-122">The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="description"></a><span data-ttu-id="4f910-123">説明</span><span class="sxs-lookup"><span data-stu-id="4f910-123">Description</span></span>

<span data-ttu-id="4f910-124">必須。</span><span class="sxs-lookup"><span data-stu-id="4f910-124">Required.</span></span> <span data-ttu-id="4f910-125">吹き出しの説明/本文の内容。</span><span class="sxs-lookup"><span data-stu-id="4f910-125">The description / body content for the callout.</span></span> <span data-ttu-id="4f910-126">**resid 属性** は、[リソース] セクションの **LongStrings** 要素 [](resources.md)の有効な ID を参照し、32 文字以内で指定できます。</span><span class="sxs-lookup"><span data-stu-id="4f910-126">The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="4f910-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="4f910-127">LearnMoreUrl</span></span>

<span data-ttu-id="4f910-128">必須。</span><span class="sxs-lookup"><span data-stu-id="4f910-128">Required.</span></span> <span data-ttu-id="4f910-129">ユーザーがアドインの詳細を参照できるページの URL。</span><span class="sxs-lookup"><span data-stu-id="4f910-129">The URL to a page where the user can learn more about your add-in.</span></span> <span data-ttu-id="4f910-130">**resid 属性** は、[リソース] セクションの **Urls** 要素 [](resources.md)の有効な ID を参照し、32 文字以内で指定できます。</span><span class="sxs-lookup"><span data-stu-id="4f910-130">The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

> [!NOTE]
> <span data-ttu-id="4f910-131">**LearnMoreUrl** は現在、Word、Excel、または PowerPoint のクライアントではレンダリングされません。</span><span class="sxs-lookup"><span data-stu-id="4f910-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="4f910-132">これが利用可能になったときに URL がレンダリングされるよう、すべてのクライアントにこの URL を追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="4f910-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="4f910-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="4f910-133">See also</span></span>

<span data-ttu-id="4f910-134">次のコード サンプルでは **、GetStarted 要素を使用** します。</span><span class="sxs-lookup"><span data-stu-id="4f910-134">The following code samples use the **GetStarted** element.</span></span>

* [<span data-ttu-id="4f910-135">テーブルとグラフの書式設定を操作するための Excel Web アドイン</span><span class="sxs-lookup"><span data-stu-id="4f910-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="4f910-136">Word アドインの JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="4f910-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="4f910-137">PowerPoint アドインで Microsoft Graph を使用して Excel グラフを挿入する</span><span class="sxs-lookup"><span data-stu-id="4f910-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
