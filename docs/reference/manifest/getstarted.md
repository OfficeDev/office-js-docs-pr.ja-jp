---
title: マニフェスト ファイルの GetStarted 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: e6fb1c56d051e9de607e97979225e484adb9affb
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433115"
---
# <a name="getstarted-element"></a><span data-ttu-id="a4153-102">GetStarted 要素</span><span class="sxs-lookup"><span data-stu-id="a4153-102">GetStarted element</span></span>

<span data-ttu-id="a4153-p101">アドインが、Word、Excel、PowerPoint、OneNote のホストにインストールされているときに表示される吹き出しで使用される情報を提供します。**GetStarted** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="a4153-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="a4153-105">子要素</span><span class="sxs-lookup"><span data-stu-id="a4153-105">Child elements</span></span>

| <span data-ttu-id="a4153-106">要素</span><span class="sxs-lookup"><span data-stu-id="a4153-106">Element</span></span>                       | <span data-ttu-id="a4153-107">必須</span><span class="sxs-lookup"><span data-stu-id="a4153-107">Required</span></span> | <span data-ttu-id="a4153-108">説明</span><span class="sxs-lookup"><span data-stu-id="a4153-108">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="a4153-109">Title</span><span class="sxs-lookup"><span data-stu-id="a4153-109">Title</span></span>](#title)               | <span data-ttu-id="a4153-110">はい</span><span class="sxs-lookup"><span data-stu-id="a4153-110">Yes</span></span>      | <span data-ttu-id="a4153-111">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="a4153-111">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="a4153-112">説明</span><span class="sxs-lookup"><span data-stu-id="a4153-112">Description</span></span>](#description)   | <span data-ttu-id="a4153-113">はい</span><span class="sxs-lookup"><span data-stu-id="a4153-113">Yes</span></span>      | <span data-ttu-id="a4153-114">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="a4153-114">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="a4153-115">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="a4153-115">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="a4153-116">いいえ</span><span class="sxs-lookup"><span data-stu-id="a4153-116">No</span></span>       | <span data-ttu-id="a4153-117">アドインの詳細を説明するページの URL。</span><span class="sxs-lookup"><span data-stu-id="a4153-117">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="a4153-118">Title</span><span class="sxs-lookup"><span data-stu-id="a4153-118">Title</span></span> 

<span data-ttu-id="a4153-p102">必須。吹き出しの一番上に使用するタイトル。**resid** 属性は [Resources](resources.md) セクションの **ShortStrings** 要素にある有効な ID を参照します。</span><span class="sxs-lookup"><span data-stu-id="a4153-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="a4153-122">説明</span><span class="sxs-lookup"><span data-stu-id="a4153-122">Description</span></span>

<span data-ttu-id="a4153-p103">必須。吹き出しの説明/本文の内容。**resid** 属性は [Resources](resources.md) セクションの **LongStrings** 要素にある有効な ID を参照します。</span><span class="sxs-lookup"><span data-stu-id="a4153-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="a4153-126">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="a4153-126">LearnMoreUrl</span></span>

<span data-ttu-id="a4153-p104">必須。ユーザーがアドインの詳細を参照できるページの URL。**resid** 属性は [Resources](resources.md) セクションの **Urls** 要素にある有効な ID を参照します。</span><span class="sxs-lookup"><span data-stu-id="a4153-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="a4153-130">**LearnMoreUrl** は現在、Word、Excel、または PowerPoint のクライアントではレンダリングされません。</span><span class="sxs-lookup"><span data-stu-id="a4153-130">NOTE:**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="a4153-131">これが利用可能になったときに URL がレンダリングされるよう、すべてのクライアントにこの URL を追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="a4153-131">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="a4153-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="a4153-132">See also</span></span>

<span data-ttu-id="a4153-133">次のコード サンプルでは、**GetStarted** 要素を使用しています。</span><span class="sxs-lookup"><span data-stu-id="a4153-133">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="a4153-134">テーブルとグラフの書式設定を操作するための Excel Web アドイン</span><span class="sxs-lookup"><span data-stu-id="a4153-134">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="a4153-135">Word アドインの JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="a4153-135">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="a4153-136">PowerPoint アドインで Microsoft Graph を使用して Excel グラフを挿入する</span><span class="sxs-lookup"><span data-stu-id="a4153-136">Insert Excel charts using Microsoft Graph in a PowerPoint Add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
