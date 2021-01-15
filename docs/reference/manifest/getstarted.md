---
title: マニフェスト ファイルの GetStarted 要素
description: Word、Excel、PowerPoint、OneNote にアドインがインストールされている場合に表示される吹き出しで使用される情報を提供します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0ad6196dc45e4ea06c2b43ac5da66a560ab0b899
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771418"
---
# <a name="getstarted-element"></a><span data-ttu-id="65fa0-103">GetStarted 要素</span><span class="sxs-lookup"><span data-stu-id="65fa0-103">GetStarted element</span></span>

<span data-ttu-id="65fa0-104">Word、Excel、PowerPoint、OneNote にアドインがインストールされている場合に表示される吹き出しで使用される情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="65fa0-104">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote.</span></span> <span data-ttu-id="65fa0-105">**GetStarted 要素は** [DesktopFormFactor の子要素です](desktopformfactor.md)。</span><span class="sxs-lookup"><span data-stu-id="65fa0-105">The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="65fa0-106">子要素</span><span class="sxs-lookup"><span data-stu-id="65fa0-106">Child elements</span></span>

| <span data-ttu-id="65fa0-107">要素</span><span class="sxs-lookup"><span data-stu-id="65fa0-107">Element</span></span>                       | <span data-ttu-id="65fa0-108">必須</span><span class="sxs-lookup"><span data-stu-id="65fa0-108">Required</span></span> | <span data-ttu-id="65fa0-109">説明</span><span class="sxs-lookup"><span data-stu-id="65fa0-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="65fa0-110">Title</span><span class="sxs-lookup"><span data-stu-id="65fa0-110">Title</span></span>](#title)               | <span data-ttu-id="65fa0-111">はい</span><span class="sxs-lookup"><span data-stu-id="65fa0-111">Yes</span></span>      | <span data-ttu-id="65fa0-112">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="65fa0-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="65fa0-113">説明</span><span class="sxs-lookup"><span data-stu-id="65fa0-113">Description</span></span>](#description)   | <span data-ttu-id="65fa0-114">はい</span><span class="sxs-lookup"><span data-stu-id="65fa0-114">Yes</span></span>      | <span data-ttu-id="65fa0-115">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="65fa0-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="65fa0-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="65fa0-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="65fa0-117">はい</span><span class="sxs-lookup"><span data-stu-id="65fa0-117">Yes</span></span>       | <span data-ttu-id="65fa0-118">アドインの詳細を説明するページの URL。</span><span class="sxs-lookup"><span data-stu-id="65fa0-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="65fa0-119">タイトル</span><span class="sxs-lookup"><span data-stu-id="65fa0-119">Title</span></span> 

<span data-ttu-id="65fa0-120">必須。</span><span class="sxs-lookup"><span data-stu-id="65fa0-120">Required.</span></span> <span data-ttu-id="65fa0-121">吹き出しの一番上に使用するタイトル。</span><span class="sxs-lookup"><span data-stu-id="65fa0-121">The title used for the top of the callout.</span></span> <span data-ttu-id="65fa0-122">**resid 属性** は [、Resources](resources.md)セクションの **ShortStrings** 要素内の有効な ID を参照し、32 文字以内で指定できます。</span><span class="sxs-lookup"><span data-stu-id="65fa0-122">The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="description"></a><span data-ttu-id="65fa0-123">説明</span><span class="sxs-lookup"><span data-stu-id="65fa0-123">Description</span></span>

<span data-ttu-id="65fa0-124">必須。</span><span class="sxs-lookup"><span data-stu-id="65fa0-124">Required.</span></span> <span data-ttu-id="65fa0-125">吹き出しの説明/本文の内容。</span><span class="sxs-lookup"><span data-stu-id="65fa0-125">The description / body content for the callout.</span></span> <span data-ttu-id="65fa0-126">**resid 属性** は [、Resources](resources.md)セクションの **LongStrings** 要素内の有効な ID を参照し、32 文字以内で指定できます。</span><span class="sxs-lookup"><span data-stu-id="65fa0-126">The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="65fa0-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="65fa0-127">LearnMoreUrl</span></span>

<span data-ttu-id="65fa0-128">必須。</span><span class="sxs-lookup"><span data-stu-id="65fa0-128">Required.</span></span> <span data-ttu-id="65fa0-129">ユーザーがアドインの詳細を参照できるページの URL。</span><span class="sxs-lookup"><span data-stu-id="65fa0-129">The URL to a page where the user can learn more about your add-in.</span></span> <span data-ttu-id="65fa0-130">**resid 属性** は [、Resources](resources.md)セクションの **Urls** 要素内の有効な ID を参照し、32 文字以内で指定できます。</span><span class="sxs-lookup"><span data-stu-id="65fa0-130">The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

> [!NOTE]
> <span data-ttu-id="65fa0-131">**LearnMoreUrl** は現在、Word、Excel、または PowerPoint のクライアントではレンダリングされません。</span><span class="sxs-lookup"><span data-stu-id="65fa0-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="65fa0-132">これが利用可能になったときに URL がレンダリングされるよう、すべてのクライアントにこの URL を追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="65fa0-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="65fa0-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="65fa0-133">See also</span></span>

<span data-ttu-id="65fa0-134">次のコード サンプルでは、**GetStarted** 要素を使用しています。</span><span class="sxs-lookup"><span data-stu-id="65fa0-134">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="65fa0-135">テーブルとグラフの書式設定を操作するための Excel Web アドイン</span><span class="sxs-lookup"><span data-stu-id="65fa0-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="65fa0-136">Word アドインの JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="65fa0-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="65fa0-137">PowerPoint アドインで Microsoft Graph を使用して Excel グラフを挿入する</span><span class="sxs-lookup"><span data-stu-id="65fa0-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
