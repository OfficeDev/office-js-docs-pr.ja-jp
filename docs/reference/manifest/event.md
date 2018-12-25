---
title: マニフェスト ファイルの Event 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: eda895b01e106d67eef70f199be64086e9372bef
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432740"
---
# <a name="event-element"></a><span data-ttu-id="4bed9-102">Event 要素</span><span class="sxs-lookup"><span data-stu-id="4bed9-102">Event element</span></span>

<span data-ttu-id="4bed9-103">アドインでイベント ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="4bed9-103">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="4bed9-104">`Event` 要素は現在、Office 365 の Outlook on the web でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="4bed9-104">Note: The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="4bed9-105">属性</span><span class="sxs-lookup"><span data-stu-id="4bed9-105">Attributes</span></span>

|  <span data-ttu-id="4bed9-106">属性</span><span class="sxs-lookup"><span data-stu-id="4bed9-106">Attribute</span></span>  |  <span data-ttu-id="4bed9-107">必須</span><span class="sxs-lookup"><span data-stu-id="4bed9-107">Required</span></span>  |  <span data-ttu-id="4bed9-108">説明</span><span class="sxs-lookup"><span data-stu-id="4bed9-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4bed9-109">Type</span><span class="sxs-lookup"><span data-stu-id="4bed9-109">Type</span></span>](#type-attribute)  |  <span data-ttu-id="4bed9-110">はい</span><span class="sxs-lookup"><span data-stu-id="4bed9-110">Yes</span></span>  | <span data-ttu-id="4bed9-111">処理するイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="4bed9-111">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="4bed9-112">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="4bed9-112">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="4bed9-113">はい</span><span class="sxs-lookup"><span data-stu-id="4bed9-113">Yes</span></span>  | <span data-ttu-id="4bed9-p101">イベント ハンドラーの実行スタイル (非同期または同期) を指定します。現在サポートされているのは同期イベント ハンドラーのみです。</span><span class="sxs-lookup"><span data-stu-id="4bed9-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="4bed9-116">FunctionName</span><span class="sxs-lookup"><span data-stu-id="4bed9-116">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="4bed9-117">はい</span><span class="sxs-lookup"><span data-stu-id="4bed9-117">Yes</span></span>  | <span data-ttu-id="4bed9-118">イベント ハンドラーの関数名を指定します。</span><span class="sxs-lookup"><span data-stu-id="4bed9-118">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="4bed9-119">Type 属性</span><span class="sxs-lookup"><span data-stu-id="4bed9-119">Type attribute</span></span>

<span data-ttu-id="4bed9-p102">必須です。イベント ハンドラーを呼び出すイベントを指定します。この属性の使用可能な値は、次の表のとおりです。</span><span class="sxs-lookup"><span data-stu-id="4bed9-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="4bed9-123">イベントの種類</span><span class="sxs-lookup"><span data-stu-id="4bed9-123">Event type</span></span>  |  <span data-ttu-id="4bed9-124">説明</span><span class="sxs-lookup"><span data-stu-id="4bed9-124">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="4bed9-125">ユーザーがメッセージまたは会議出席依頼を送信すると、イベント ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="4bed9-125">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="4bed9-126">FunctionExecution 属性</span><span class="sxs-lookup"><span data-stu-id="4bed9-126">FunctionExecution attribute</span></span>

<span data-ttu-id="4bed9-p103">必須です。`synchronous` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4bed9-p103">Required. MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="4bed9-129">FunctionName 属性</span><span class="sxs-lookup"><span data-stu-id="4bed9-129">FunctionName attribute</span></span>

<span data-ttu-id="4bed9-p104">必須です。イベント ハンドラーの関数名を指定します。この値は、アドインの[関数ファイル](functionfile.md)内の関数名と一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4bed9-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```