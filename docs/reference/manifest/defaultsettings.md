---
title: マニフェスト ファイルの DefaultSettings 要素
description: コンテンツまたは作業ウィンドウ アドインの既定のソースの場所と他の既定の設定を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b97f692a1fd39e4b1f55080f6ed77e623be0000c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718371"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="b993d-103">DefaultSettings 要素</span><span class="sxs-lookup"><span data-stu-id="b993d-103">DefaultSettings element</span></span>

<span data-ttu-id="b993d-104">コンテンツまたは作業ウィンドウ アドインの既定のソースの場所と他の既定の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="b993d-104">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="b993d-105">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="b993d-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="b993d-106">構文</span><span class="sxs-lookup"><span data-stu-id="b993d-106">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="b993d-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="b993d-107">Contained in</span></span>

[<span data-ttu-id="b993d-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b993d-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="b993d-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="b993d-109">Can contain</span></span>

|<span data-ttu-id="b993d-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="b993d-110">**Element**</span></span>|<span data-ttu-id="b993d-111">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="b993d-111">**Content**</span></span>|<span data-ttu-id="b993d-112">**メール**</span><span class="sxs-lookup"><span data-stu-id="b993d-112">**Mail**</span></span>|<span data-ttu-id="b993d-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="b993d-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="b993d-114">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="b993d-114">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="b993d-115">x</span><span class="sxs-lookup"><span data-stu-id="b993d-115">x</span></span>||<span data-ttu-id="b993d-116">x</span><span class="sxs-lookup"><span data-stu-id="b993d-116">x</span></span>|
|[<span data-ttu-id="b993d-117">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="b993d-117">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="b993d-118">x</span><span class="sxs-lookup"><span data-stu-id="b993d-118">x</span></span>|||
|[<span data-ttu-id="b993d-119">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="b993d-119">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="b993d-120">x</span><span class="sxs-lookup"><span data-stu-id="b993d-120">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="b993d-121">注釈</span><span class="sxs-lookup"><span data-stu-id="b993d-121">Remarks</span></span>

<span data-ttu-id="b993d-122">**DefaultSettings**要素のソースの場所とその他の設定は、コンテンツアドインと作業ウィンドウアドインにのみ適用されます。メールアドインの場合は、 [formsettings](formsettings.md)要素に、ソースファイルとその他の既定の設定の既定の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="b993d-122">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

