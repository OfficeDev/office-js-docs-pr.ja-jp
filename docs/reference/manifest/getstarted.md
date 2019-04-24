---
title: マニフェスト ファイルの GetStarted 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d9ebcba7881b388544eeb3e2c3028bff9bdcf9a6
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452082"
---
# <a name="getstarted-element"></a><span data-ttu-id="b7bd8-102">GetStarted 要素</span><span class="sxs-lookup"><span data-stu-id="b7bd8-102">GetStarted element</span></span>

<span data-ttu-id="b7bd8-p101">アドインが、Word、Excel、PowerPoint、OneNote のホストにインストールされているときに表示される吹き出しで使用される情報を提供します。**GetStarted** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="b7bd8-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="b7bd8-105">子要素</span><span class="sxs-lookup"><span data-stu-id="b7bd8-105">Child elements</span></span>

| <span data-ttu-id="b7bd8-106">要素</span><span class="sxs-lookup"><span data-stu-id="b7bd8-106">Element</span></span>                       | <span data-ttu-id="b7bd8-107">必須</span><span class="sxs-lookup"><span data-stu-id="b7bd8-107">Required</span></span> | <span data-ttu-id="b7bd8-108">説明</span><span class="sxs-lookup"><span data-stu-id="b7bd8-108">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="b7bd8-109">Title</span><span class="sxs-lookup"><span data-stu-id="b7bd8-109">Title</span></span>](#title)               | <span data-ttu-id="b7bd8-110">はい</span><span class="sxs-lookup"><span data-stu-id="b7bd8-110">Yes</span></span>      | <span data-ttu-id="b7bd8-111">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="b7bd8-111">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="b7bd8-112">説明</span><span class="sxs-lookup"><span data-stu-id="b7bd8-112">Description</span></span>](#description)   | <span data-ttu-id="b7bd8-113">はい</span><span class="sxs-lookup"><span data-stu-id="b7bd8-113">Yes</span></span>      | <span data-ttu-id="b7bd8-114">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="b7bd8-114">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="b7bd8-115">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="b7bd8-115">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="b7bd8-116">いいえ</span><span class="sxs-lookup"><span data-stu-id="b7bd8-116">No</span></span>       | <span data-ttu-id="b7bd8-117">アドインの詳細を説明するページの URL。</span><span class="sxs-lookup"><span data-stu-id="b7bd8-117">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="b7bd8-118">タイトル</span><span class="sxs-lookup"><span data-stu-id="b7bd8-118">Title</span></span> 

<span data-ttu-id="b7bd8-p102">必須。 吹き出しの一番上に使用するタイトル。 **resid** 属性は **Resources** セクションの [ShortStrings](resources.md) 要素にある有効な ID を参照します。</span><span class="sxs-lookup"><span data-stu-id="b7bd8-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="b7bd8-122">説明</span><span class="sxs-lookup"><span data-stu-id="b7bd8-122">Description</span></span>

<span data-ttu-id="b7bd8-p103">必須。 吹き出しの説明/本文の内容。 **resid** 属性は **Resources** セクションの [LongStrings](resources.md) 要素にある有効な ID を参照します。</span><span class="sxs-lookup"><span data-stu-id="b7bd8-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="b7bd8-126">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="b7bd8-126">LearnMoreUrl</span></span>

<span data-ttu-id="b7bd8-p104">必須。ユーザーがアドインの詳細を参照できるページの URL。**resid** 属性は [Resources](resources.md) セクションの **Urls** 要素にある有効な ID を参照します。</span><span class="sxs-lookup"><span data-stu-id="b7bd8-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="b7bd8-130">**LearnMoreUrl** は現在、Word、Excel、または PowerPoint のクライアントではレンダリングされません。</span><span class="sxs-lookup"><span data-stu-id="b7bd8-130">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="b7bd8-131">これが利用可能になったときに URL がレンダリングされるよう、すべてのクライアントにこの URL を追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="b7bd8-131">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="b7bd8-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="b7bd8-132">See also</span></span>

<span data-ttu-id="b7bd8-133">次のコード サンプルでは、**GetStarted** 要素を使用しています。</span><span class="sxs-lookup"><span data-stu-id="b7bd8-133">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="b7bd8-134">テーブルとグラフの書式設定を操作するための Excel Web アドイン</span><span class="sxs-lookup"><span data-stu-id="b7bd8-134">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="b7bd8-135">Word アドインの JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="b7bd8-135">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="b7bd8-136">PowerPoint アドインで Microsoft Graph を使用して Excel グラフを挿入する</span><span class="sxs-lookup"><span data-stu-id="b7bd8-136">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
