---
title: マニフェスト ファイルの Supertip 要素
description: ヒント要素は、リッチツールヒント (タイトルと説明の両方) を定義します。
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: cf88473b72979c839e5d55f44938fda19be24084
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720352"
---
# <a name="supertip"></a><span data-ttu-id="d0d17-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="d0d17-103">Supertip</span></span>

<span data-ttu-id="d0d17-p101">豊富なヒント (タイトルと説明の両方) を定義します。[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの両方で使用されます。</span><span class="sxs-lookup"><span data-stu-id="d0d17-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="d0d17-106">子要素</span><span class="sxs-lookup"><span data-stu-id="d0d17-106">Child elements</span></span>

|  <span data-ttu-id="d0d17-107">要素</span><span class="sxs-lookup"><span data-stu-id="d0d17-107">Element</span></span> |  <span data-ttu-id="d0d17-108">必須</span><span class="sxs-lookup"><span data-stu-id="d0d17-108">Required</span></span>  |  <span data-ttu-id="d0d17-109">説明</span><span class="sxs-lookup"><span data-stu-id="d0d17-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="d0d17-110">Title</span><span class="sxs-lookup"><span data-stu-id="d0d17-110">Title</span></span>](#title) | <span data-ttu-id="d0d17-111">はい</span><span class="sxs-lookup"><span data-stu-id="d0d17-111">Yes</span></span> | <span data-ttu-id="d0d17-112">ヒントのテキストです。</span><span class="sxs-lookup"><span data-stu-id="d0d17-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="d0d17-113">説明</span><span class="sxs-lookup"><span data-stu-id="d0d17-113">Description</span></span>](#description) | <span data-ttu-id="d0d17-114">はい</span><span class="sxs-lookup"><span data-stu-id="d0d17-114">Yes</span></span> | <span data-ttu-id="d0d17-115">ヒントの説明です。</span><span class="sxs-lookup"><span data-stu-id="d0d17-115">The description for the supertip.</span></span><br><span data-ttu-id="d0d17-116">**注**: (Outlook) は、Windows および Mac クライアントのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="d0d17-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="d0d17-117">Title</span><span class="sxs-lookup"><span data-stu-id="d0d17-117">Title</span></span>

<span data-ttu-id="d0d17-118">必須です。</span><span class="sxs-lookup"><span data-stu-id="d0d17-118">Required.</span></span> <span data-ttu-id="d0d17-119">ヒントのテキスト。</span><span class="sxs-lookup"><span data-stu-id="d0d17-119">The text for the supertip.</span></span> <span data-ttu-id="d0d17-120">**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d0d17-120">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="d0d17-121">説明</span><span class="sxs-lookup"><span data-stu-id="d0d17-121">Description</span></span>

<span data-ttu-id="d0d17-122">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="d0d17-122">Required.</span></span> <span data-ttu-id="d0d17-123">ヒントの記述です。</span><span class="sxs-lookup"><span data-stu-id="d0d17-123">The description for the supertip.</span></span> <span data-ttu-id="d0d17-124">**Resid**属性は、 [Resources](resources.md)要素の**longstrings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d0d17-124">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="d0d17-125">Outlook の場合、Windows と Mac のクライアントのみが**Description**要素をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="d0d17-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="d0d17-126">例</span><span class="sxs-lookup"><span data-stu-id="d0d17-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
