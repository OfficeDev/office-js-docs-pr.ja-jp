---
title: マニフェスト ファイルの Supertip 要素
description: Supertip 要素は、豊富なヒント (タイトルと説明の両方) を定義します。
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 5e8b3850d99f6791726b1b2f0545c5fb4b52c554
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771299"
---
# <a name="supertip"></a><span data-ttu-id="a8304-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="a8304-103">Supertip</span></span>

<span data-ttu-id="a8304-p101">豊富なヒント (タイトルと説明の両方) を定義します。[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの両方で使用されます。</span><span class="sxs-lookup"><span data-stu-id="a8304-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="a8304-106">子要素</span><span class="sxs-lookup"><span data-stu-id="a8304-106">Child elements</span></span>

|  <span data-ttu-id="a8304-107">要素</span><span class="sxs-lookup"><span data-stu-id="a8304-107">Element</span></span> |  <span data-ttu-id="a8304-108">必須</span><span class="sxs-lookup"><span data-stu-id="a8304-108">Required</span></span>  |  <span data-ttu-id="a8304-109">説明</span><span class="sxs-lookup"><span data-stu-id="a8304-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="a8304-110">Title</span><span class="sxs-lookup"><span data-stu-id="a8304-110">Title</span></span>](#title) | <span data-ttu-id="a8304-111">はい</span><span class="sxs-lookup"><span data-stu-id="a8304-111">Yes</span></span> | <span data-ttu-id="a8304-112">ヒントのテキストです。</span><span class="sxs-lookup"><span data-stu-id="a8304-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="a8304-113">説明</span><span class="sxs-lookup"><span data-stu-id="a8304-113">Description</span></span>](#description) | <span data-ttu-id="a8304-114">はい</span><span class="sxs-lookup"><span data-stu-id="a8304-114">Yes</span></span> | <span data-ttu-id="a8304-115">ヒントの説明です。</span><span class="sxs-lookup"><span data-stu-id="a8304-115">The description for the supertip.</span></span><br><span data-ttu-id="a8304-116">**注**: (Outlook) Windows および Mac クライアントだけがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="a8304-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="a8304-117">タイトル</span><span class="sxs-lookup"><span data-stu-id="a8304-117">Title</span></span>

<span data-ttu-id="a8304-118">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="a8304-118">Required.</span></span> <span data-ttu-id="a8304-119">ヒントのテキストです。</span><span class="sxs-lookup"><span data-stu-id="a8304-119">The text for the supertip.</span></span> <span data-ttu-id="a8304-120">**resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a8304-120">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="a8304-121">説明</span><span class="sxs-lookup"><span data-stu-id="a8304-121">Description</span></span>

<span data-ttu-id="a8304-122">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="a8304-122">Required.</span></span> <span data-ttu-id="a8304-123">ヒントの記述です。</span><span class="sxs-lookup"><span data-stu-id="a8304-123">The description for the supertip.</span></span> <span data-ttu-id="a8304-124">**resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **LongStrings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a8304-124">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="a8304-125">Outlook では、Description 要素をサポートしているのは Windows クライアントと Mac **クライアント** のみです。</span><span class="sxs-lookup"><span data-stu-id="a8304-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="a8304-126">例</span><span class="sxs-lookup"><span data-stu-id="a8304-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
