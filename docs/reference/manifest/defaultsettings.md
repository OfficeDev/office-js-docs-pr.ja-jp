---
title: マニフェスト ファイルの DefaultSettings 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0c109d5d893cf9d3502f1cbf1724007f01e623e6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433755"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="de292-102">DefaultSettings 要素</span><span class="sxs-lookup"><span data-stu-id="de292-102">DefaultSettings element</span></span>

<span data-ttu-id="de292-103">コンテンツまたは作業ウィンドウ アドインの既定のソースの場所と他の既定の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="de292-103">Specifies the default source location and other default settings for your content or task pane add-in .</span></span>

<span data-ttu-id="de292-104">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="de292-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="de292-105">構文</span><span class="sxs-lookup"><span data-stu-id="de292-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="de292-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="de292-106">Contained in</span></span>

[<span data-ttu-id="de292-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="de292-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="de292-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="de292-108">Can contain</span></span>

|<span data-ttu-id="de292-109">**要素**</span><span class="sxs-lookup"><span data-stu-id="de292-109">**Element**</span></span>|<span data-ttu-id="de292-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="de292-110">**Content**</span></span>|<span data-ttu-id="de292-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="de292-111">**Mail**</span></span>|<span data-ttu-id="de292-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="de292-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="de292-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="de292-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="de292-114">x</span><span class="sxs-lookup"><span data-stu-id="de292-114">x</span></span>||<span data-ttu-id="de292-115">x</span><span class="sxs-lookup"><span data-stu-id="de292-115">x</span></span>|
|[<span data-ttu-id="de292-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="de292-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="de292-117">x</span><span class="sxs-lookup"><span data-stu-id="de292-117">x</span></span>|||
|[<span data-ttu-id="de292-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="de292-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="de292-119">x</span><span class="sxs-lookup"><span data-stu-id="de292-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="de292-120">注釈</span><span class="sxs-lookup"><span data-stu-id="de292-120">Remarks</span></span>

<span data-ttu-id="de292-121">**DefaultSettings** 要素のソースの場所と他の設定が適用されるのは、コンテンツ アドインと作業ウィンドウ アドインのみです。メール アドインの場合は、ソース ファイルの既定の場所とその他の既定の設定を [FormSettings](formsettings.md) 要素に指定します。</span><span class="sxs-lookup"><span data-stu-id="de292-121">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

