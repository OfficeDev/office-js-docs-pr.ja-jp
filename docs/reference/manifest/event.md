---
title: マニフェストファイルの Event 要素
description: アドインでイベント ハンドラーを定義します。
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 80f21d1819e3d7e335389070ccac0db583026045
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275708"
---
# <a name="event-element"></a><span data-ttu-id="77818-103">Event 要素</span><span class="sxs-lookup"><span data-stu-id="77818-103">Event element</span></span>

<span data-ttu-id="77818-104">アドインでイベント ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="77818-104">Defines an event handler in an add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="77818-105">サポートと使用法の詳細については、「 [Outlook アドインの送信時機能](../../outlook/outlook-on-send-addins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="77818-105">For information about support and usage, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="77818-106">属性</span><span class="sxs-lookup"><span data-stu-id="77818-106">Attributes</span></span>

|  <span data-ttu-id="77818-107">属性</span><span class="sxs-lookup"><span data-stu-id="77818-107">Attribute</span></span>  |  <span data-ttu-id="77818-108">必須</span><span class="sxs-lookup"><span data-stu-id="77818-108">Required</span></span>  |  <span data-ttu-id="77818-109">説明</span><span class="sxs-lookup"><span data-stu-id="77818-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="77818-110">種類</span><span class="sxs-lookup"><span data-stu-id="77818-110">Type</span></span>](#type-attribute)  |  <span data-ttu-id="77818-111">はい</span><span class="sxs-lookup"><span data-stu-id="77818-111">Yes</span></span>  | <span data-ttu-id="77818-112">処理するイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="77818-112">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="77818-113">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="77818-113">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="77818-114">はい</span><span class="sxs-lookup"><span data-stu-id="77818-114">Yes</span></span>  | <span data-ttu-id="77818-p101">イベント ハンドラーの実行スタイル (非同期または同期) を指定します。現在サポートされているのは同期イベント ハンドラーのみです。</span><span class="sxs-lookup"><span data-stu-id="77818-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="77818-117">FunctionName</span><span class="sxs-lookup"><span data-stu-id="77818-117">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="77818-118">はい</span><span class="sxs-lookup"><span data-stu-id="77818-118">Yes</span></span>  | <span data-ttu-id="77818-119">イベント ハンドラーの関数名を指定します。</span><span class="sxs-lookup"><span data-stu-id="77818-119">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="77818-120">Type 属性</span><span class="sxs-lookup"><span data-stu-id="77818-120">Type attribute</span></span>

<span data-ttu-id="77818-p102">必須です。イベント ハンドラーを呼び出すイベントを指定します。この属性の使用可能な値は、次の表のとおりです。</span><span class="sxs-lookup"><span data-stu-id="77818-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="77818-124">イベントの種類</span><span class="sxs-lookup"><span data-stu-id="77818-124">Event type</span></span>  |  <span data-ttu-id="77818-125">説明</span><span class="sxs-lookup"><span data-stu-id="77818-125">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="77818-126">ユーザーがメッセージまたは会議出席依頼を送信すると、イベント ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="77818-126">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="77818-127">FunctionExecution 属性</span><span class="sxs-lookup"><span data-stu-id="77818-127">FunctionExecution attribute</span></span>

<span data-ttu-id="77818-128">必須です。</span><span class="sxs-lookup"><span data-stu-id="77818-128">Required.</span></span> <span data-ttu-id="77818-129">に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="77818-129">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="77818-130">FunctionName 属性</span><span class="sxs-lookup"><span data-stu-id="77818-130">FunctionName attribute</span></span>

<span data-ttu-id="77818-p104">必須です。イベント ハンドラーの関数名を指定します。この値は、アドインの[関数ファイル](functionfile.md)内の関数名と一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="77818-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
