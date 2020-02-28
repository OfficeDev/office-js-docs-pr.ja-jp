---
title: マニフェスト ファイルの DefaultSettings 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 824c575b39a99c6028ffd603390d2b41ee0ad7dd
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324885"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="d4aa2-102">DefaultSettings 要素</span><span class="sxs-lookup"><span data-stu-id="d4aa2-102">DefaultSettings element</span></span>

<span data-ttu-id="d4aa2-103">コンテンツまたは作業ウィンドウ アドインの既定のソースの場所と他の既定の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="d4aa2-103">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="d4aa2-104">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="d4aa2-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="d4aa2-105">構文</span><span class="sxs-lookup"><span data-stu-id="d4aa2-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="d4aa2-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="d4aa2-106">Contained in</span></span>

[<span data-ttu-id="d4aa2-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="d4aa2-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="d4aa2-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="d4aa2-108">Can contain</span></span>

|<span data-ttu-id="d4aa2-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="d4aa2-109">**Element**</span></span>|<span data-ttu-id="d4aa2-110">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="d4aa2-110">**Content**</span></span>|<span data-ttu-id="d4aa2-111">**メール**</span><span class="sxs-lookup"><span data-stu-id="d4aa2-111">**Mail**</span></span>|<span data-ttu-id="d4aa2-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="d4aa2-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="d4aa2-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="d4aa2-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="d4aa2-114">x</span><span class="sxs-lookup"><span data-stu-id="d4aa2-114">x</span></span>||<span data-ttu-id="d4aa2-115">x</span><span class="sxs-lookup"><span data-stu-id="d4aa2-115">x</span></span>|
|[<span data-ttu-id="d4aa2-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="d4aa2-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="d4aa2-117">x</span><span class="sxs-lookup"><span data-stu-id="d4aa2-117">x</span></span>|||
|[<span data-ttu-id="d4aa2-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="d4aa2-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="d4aa2-119">x</span><span class="sxs-lookup"><span data-stu-id="d4aa2-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="d4aa2-120">注釈</span><span class="sxs-lookup"><span data-stu-id="d4aa2-120">Remarks</span></span>

<span data-ttu-id="d4aa2-121">**DefaultSettings**要素のソースの場所とその他の設定は、コンテンツアドインと作業ウィンドウアドインにのみ適用されます。メールアドインの場合は、 [formsettings](formsettings.md)要素に、ソースファイルとその他の既定の設定の既定の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="d4aa2-121">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

