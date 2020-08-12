---
title: マニフェスト ファイルの DefaultSettings 要素
description: コンテンツまたは作業ウィンドウ アドインの既定のソースの場所と他の既定の設定を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a9711fb44390bcbda8979b8018eed1318c5579bc
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641467"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="d8c3f-103">DefaultSettings 要素</span><span class="sxs-lookup"><span data-stu-id="d8c3f-103">DefaultSettings element</span></span>

<span data-ttu-id="d8c3f-104">コンテンツまたは作業ウィンドウ アドインの既定のソースの場所と他の既定の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="d8c3f-104">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="d8c3f-105">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="d8c3f-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="d8c3f-106">構文</span><span class="sxs-lookup"><span data-stu-id="d8c3f-106">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="d8c3f-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="d8c3f-107">Contained in</span></span>

[<span data-ttu-id="d8c3f-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="d8c3f-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="d8c3f-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="d8c3f-109">Can contain</span></span>

|<span data-ttu-id="d8c3f-110">要素</span><span class="sxs-lookup"><span data-stu-id="d8c3f-110">Element</span></span>|<span data-ttu-id="d8c3f-111">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="d8c3f-111">Content</span></span>|<span data-ttu-id="d8c3f-112">メール</span><span class="sxs-lookup"><span data-stu-id="d8c3f-112">Mail</span></span>|<span data-ttu-id="d8c3f-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="d8c3f-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="d8c3f-114">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="d8c3f-114">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="d8c3f-115">x</span><span class="sxs-lookup"><span data-stu-id="d8c3f-115">x</span></span>||<span data-ttu-id="d8c3f-116">x</span><span class="sxs-lookup"><span data-stu-id="d8c3f-116">x</span></span>|
|[<span data-ttu-id="d8c3f-117">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="d8c3f-117">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="d8c3f-118">x</span><span class="sxs-lookup"><span data-stu-id="d8c3f-118">x</span></span>|||
|[<span data-ttu-id="d8c3f-119">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="d8c3f-119">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="d8c3f-120">x</span><span class="sxs-lookup"><span data-stu-id="d8c3f-120">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="d8c3f-121">注釈</span><span class="sxs-lookup"><span data-stu-id="d8c3f-121">Remarks</span></span>

<span data-ttu-id="d8c3f-122">**DefaultSettings**要素のソースの場所とその他の設定は、コンテンツアドインと作業ウィンドウアドインにのみ適用されます。メールアドインの場合は、 [formsettings](formsettings.md)要素に、ソースファイルとその他の既定の設定の既定の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="d8c3f-122">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>
