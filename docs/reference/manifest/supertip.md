---
title: マニフェスト ファイルの Supertip 要素
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: ab280ec550a58f85082c36a24f5f7c3b4112a214
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325235"
---
# <a name="supertip"></a><span data-ttu-id="dd95e-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="dd95e-102">Supertip</span></span>

<span data-ttu-id="dd95e-p101">豊富なヒント (タイトルと説明の両方) を定義します。[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの両方で使用されます。</span><span class="sxs-lookup"><span data-stu-id="dd95e-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="dd95e-105">子要素</span><span class="sxs-lookup"><span data-stu-id="dd95e-105">Child elements</span></span>

|  <span data-ttu-id="dd95e-106">要素</span><span class="sxs-lookup"><span data-stu-id="dd95e-106">Element</span></span> |  <span data-ttu-id="dd95e-107">必須</span><span class="sxs-lookup"><span data-stu-id="dd95e-107">Required</span></span>  |  <span data-ttu-id="dd95e-108">説明</span><span class="sxs-lookup"><span data-stu-id="dd95e-108">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="dd95e-109">Title</span><span class="sxs-lookup"><span data-stu-id="dd95e-109">Title</span></span>](#title) | <span data-ttu-id="dd95e-110">はい</span><span class="sxs-lookup"><span data-stu-id="dd95e-110">Yes</span></span> | <span data-ttu-id="dd95e-111">ヒントのテキストです。</span><span class="sxs-lookup"><span data-stu-id="dd95e-111">The text for the supertip.</span></span> |
| [<span data-ttu-id="dd95e-112">説明</span><span class="sxs-lookup"><span data-stu-id="dd95e-112">Description</span></span>](#description) | <span data-ttu-id="dd95e-113">はい</span><span class="sxs-lookup"><span data-stu-id="dd95e-113">Yes</span></span> | <span data-ttu-id="dd95e-114">ヒントの説明です。</span><span class="sxs-lookup"><span data-stu-id="dd95e-114">The description for the supertip.</span></span><br><span data-ttu-id="dd95e-115">**注**: (Outlook) は、Windows および Mac クライアントのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="dd95e-115">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="dd95e-116">Title</span><span class="sxs-lookup"><span data-stu-id="dd95e-116">Title</span></span>

<span data-ttu-id="dd95e-117">必須です。</span><span class="sxs-lookup"><span data-stu-id="dd95e-117">Required.</span></span> <span data-ttu-id="dd95e-118">ヒントのテキスト。</span><span class="sxs-lookup"><span data-stu-id="dd95e-118">The text for the supertip.</span></span> <span data-ttu-id="dd95e-119">**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dd95e-119">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="dd95e-120">説明</span><span class="sxs-lookup"><span data-stu-id="dd95e-120">Description</span></span>

<span data-ttu-id="dd95e-121">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="dd95e-121">Required.</span></span> <span data-ttu-id="dd95e-122">ヒントの記述です。</span><span class="sxs-lookup"><span data-stu-id="dd95e-122">The description for the supertip.</span></span> <span data-ttu-id="dd95e-123">**Resid**属性は、 [Resources](resources.md)要素の**longstrings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dd95e-123">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="dd95e-124">Outlook の場合、Windows と Mac のクライアントのみが**Description**要素をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="dd95e-124">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="dd95e-125">例</span><span class="sxs-lookup"><span data-stu-id="dd95e-125">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
