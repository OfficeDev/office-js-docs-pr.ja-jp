---
title: マニフェストファイルの Event 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 51bbcd5a3d5abe60b850e88e4063e6bbc2da37bc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450591"
---
# <a name="event-element"></a><span data-ttu-id="747e7-102">Event 要素</span><span class="sxs-lookup"><span data-stu-id="747e7-102">Event element</span></span>

<span data-ttu-id="747e7-103">アドインでイベント ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="747e7-103">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="747e7-104">この`Event`要素は、現在 Office 365 の Outlook on the web でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="747e7-104">The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="747e7-105">属性</span><span class="sxs-lookup"><span data-stu-id="747e7-105">Attributes</span></span>

|  <span data-ttu-id="747e7-106">属性</span><span class="sxs-lookup"><span data-stu-id="747e7-106">Attribute</span></span>  |  <span data-ttu-id="747e7-107">必須</span><span class="sxs-lookup"><span data-stu-id="747e7-107">Required</span></span>  |  <span data-ttu-id="747e7-108">説明</span><span class="sxs-lookup"><span data-stu-id="747e7-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="747e7-109">Type</span><span class="sxs-lookup"><span data-stu-id="747e7-109">Type</span></span>](#type-attribute)  |  <span data-ttu-id="747e7-110">はい</span><span class="sxs-lookup"><span data-stu-id="747e7-110">Yes</span></span>  | <span data-ttu-id="747e7-111">処理するイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="747e7-111">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="747e7-112">functionexecution</span><span class="sxs-lookup"><span data-stu-id="747e7-112">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="747e7-113">はい</span><span class="sxs-lookup"><span data-stu-id="747e7-113">Yes</span></span>  | <span data-ttu-id="747e7-p101">イベント ハンドラーの実行スタイル (非同期または同期) を指定します。現在サポートされているのは同期イベント ハンドラーのみです。</span><span class="sxs-lookup"><span data-stu-id="747e7-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="747e7-116">FunctionName</span><span class="sxs-lookup"><span data-stu-id="747e7-116">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="747e7-117">はい</span><span class="sxs-lookup"><span data-stu-id="747e7-117">Yes</span></span>  | <span data-ttu-id="747e7-118">イベント ハンドラーの関数名を指定します。</span><span class="sxs-lookup"><span data-stu-id="747e7-118">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="747e7-119">Type 属性</span><span class="sxs-lookup"><span data-stu-id="747e7-119">Type attribute</span></span>

<span data-ttu-id="747e7-p102">必須です。イベント ハンドラーを呼び出すイベントを指定します。この属性の使用可能な値は、次の表のとおりです。</span><span class="sxs-lookup"><span data-stu-id="747e7-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="747e7-123">イベントの種類</span><span class="sxs-lookup"><span data-stu-id="747e7-123">Event type</span></span>  |  <span data-ttu-id="747e7-124">説明</span><span class="sxs-lookup"><span data-stu-id="747e7-124">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="747e7-125">ユーザーがメッセージまたは会議出席依頼を送信すると、イベント ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="747e7-125">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="747e7-126">FunctionExecution 属性</span><span class="sxs-lookup"><span data-stu-id="747e7-126">FunctionExecution attribute</span></span>

<span data-ttu-id="747e7-127">必須です。</span><span class="sxs-lookup"><span data-stu-id="747e7-127">Required.</span></span> <span data-ttu-id="747e7-128">に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="747e7-128">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="747e7-129">FunctionName 属性</span><span class="sxs-lookup"><span data-stu-id="747e7-129">FunctionName attribute</span></span>

<span data-ttu-id="747e7-p104">必須です。イベント ハンドラーの関数名を指定します。この値は、アドインの[関数ファイル](functionfile.md)内の関数名と一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="747e7-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```
