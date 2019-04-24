---
title: マニフェスト ファイルの Supertip 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cdbba342fa591ddff3faf94ecd63a4740fb904da
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450542"
---
# <a name="supertip"></a><span data-ttu-id="b8fb2-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="b8fb2-102">Supertip</span></span>

<span data-ttu-id="b8fb2-p101">豊富なヒント (タイトルと説明の両方) を定義します。[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの両方で使用されます。</span><span class="sxs-lookup"><span data-stu-id="b8fb2-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b8fb2-105">子要素</span><span class="sxs-lookup"><span data-stu-id="b8fb2-105">Child elements</span></span>

|  <span data-ttu-id="b8fb2-106">要素</span><span class="sxs-lookup"><span data-stu-id="b8fb2-106">Element</span></span> |  <span data-ttu-id="b8fb2-107">必須</span><span class="sxs-lookup"><span data-stu-id="b8fb2-107">Required</span></span>  |  <span data-ttu-id="b8fb2-108">説明</span><span class="sxs-lookup"><span data-stu-id="b8fb2-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b8fb2-109">Title</span><span class="sxs-lookup"><span data-stu-id="b8fb2-109">Title</span></span>](#title)        | <span data-ttu-id="b8fb2-110">はい</span><span class="sxs-lookup"><span data-stu-id="b8fb2-110">Yes</span></span> |   <span data-ttu-id="b8fb2-111">ヒントのテキストです。</span><span class="sxs-lookup"><span data-stu-id="b8fb2-111">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="b8fb2-112">説明</span><span class="sxs-lookup"><span data-stu-id="b8fb2-112">Description</span></span>](#description)  | <span data-ttu-id="b8fb2-113">はい</span><span class="sxs-lookup"><span data-stu-id="b8fb2-113">Yes</span></span> |  <span data-ttu-id="b8fb2-114">ヒントの説明です。</span><span class="sxs-lookup"><span data-stu-id="b8fb2-114">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="b8fb2-115">タイトル</span><span class="sxs-lookup"><span data-stu-id="b8fb2-115">Title</span></span>

<span data-ttu-id="b8fb2-p102">必ず指定します。ヒントのテキストです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b8fb2-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="b8fb2-119">説明</span><span class="sxs-lookup"><span data-stu-id="b8fb2-119">Description</span></span>

<span data-ttu-id="b8fb2-p103">必ず指定します。ヒントの記述です。 **resid** 属性には、 **Resources** 要素の **LongStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b8fb2-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="b8fb2-123">例</span><span class="sxs-lookup"><span data-stu-id="b8fb2-123">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
