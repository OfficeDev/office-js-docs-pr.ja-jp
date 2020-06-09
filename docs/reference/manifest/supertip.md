---
title: マニフェスト ファイルの Supertip 要素
description: ヒント要素は、リッチツールヒント (タイトルと説明の両方) を定義します。
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 8061c9dcd7903db0f1265084498d6c86654e1dfa
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608720"
---
# <a name="supertip"></a><span data-ttu-id="cecab-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="cecab-103">Supertip</span></span>

<span data-ttu-id="cecab-p101">豊富なヒント (タイトルと説明の両方) を定義します。[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの両方で使用されます。</span><span class="sxs-lookup"><span data-stu-id="cecab-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="cecab-106">子要素</span><span class="sxs-lookup"><span data-stu-id="cecab-106">Child elements</span></span>

|  <span data-ttu-id="cecab-107">要素</span><span class="sxs-lookup"><span data-stu-id="cecab-107">Element</span></span> |  <span data-ttu-id="cecab-108">必須</span><span class="sxs-lookup"><span data-stu-id="cecab-108">Required</span></span>  |  <span data-ttu-id="cecab-109">説明</span><span class="sxs-lookup"><span data-stu-id="cecab-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="cecab-110">Title</span><span class="sxs-lookup"><span data-stu-id="cecab-110">Title</span></span>](#title) | <span data-ttu-id="cecab-111">はい</span><span class="sxs-lookup"><span data-stu-id="cecab-111">Yes</span></span> | <span data-ttu-id="cecab-112">ヒントのテキストです。</span><span class="sxs-lookup"><span data-stu-id="cecab-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="cecab-113">説明</span><span class="sxs-lookup"><span data-stu-id="cecab-113">Description</span></span>](#description) | <span data-ttu-id="cecab-114">はい</span><span class="sxs-lookup"><span data-stu-id="cecab-114">Yes</span></span> | <span data-ttu-id="cecab-115">ヒントの説明です。</span><span class="sxs-lookup"><span data-stu-id="cecab-115">The description for the supertip.</span></span><br><span data-ttu-id="cecab-116">**注**: (Outlook) は、Windows および Mac クライアントのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="cecab-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="cecab-117">Title</span><span class="sxs-lookup"><span data-stu-id="cecab-117">Title</span></span>

<span data-ttu-id="cecab-118">必須です。</span><span class="sxs-lookup"><span data-stu-id="cecab-118">Required.</span></span> <span data-ttu-id="cecab-119">ヒントのテキスト。</span><span class="sxs-lookup"><span data-stu-id="cecab-119">The text for the supertip.</span></span> <span data-ttu-id="cecab-120">**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cecab-120">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="cecab-121">説明</span><span class="sxs-lookup"><span data-stu-id="cecab-121">Description</span></span>

<span data-ttu-id="cecab-122">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="cecab-122">Required.</span></span> <span data-ttu-id="cecab-123">ヒントの記述です。</span><span class="sxs-lookup"><span data-stu-id="cecab-123">The description for the supertip.</span></span> <span data-ttu-id="cecab-124">**Resid**属性は、 [Resources](resources.md)要素の**longstrings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cecab-124">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="cecab-125">Outlook の場合、Windows と Mac のクライアントのみが**Description**要素をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="cecab-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="cecab-126">例</span><span class="sxs-lookup"><span data-stu-id="cecab-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
