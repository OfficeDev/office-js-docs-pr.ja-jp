---
title: マニフェスト ファイルの Supertip 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: bae997eda8e1055c5be76382456ba83acca7b91c
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433671"
---
# <a name="supertip"></a><span data-ttu-id="8df98-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="8df98-102">Supertip</span></span>

<span data-ttu-id="8df98-p101">豊富なヒント (タイトルと説明の両方) を定義します。[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの両方で使用されます。</span><span class="sxs-lookup"><span data-stu-id="8df98-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="8df98-105">子要素</span><span class="sxs-lookup"><span data-stu-id="8df98-105">Child elements</span></span>

|  <span data-ttu-id="8df98-106">要素</span><span class="sxs-lookup"><span data-stu-id="8df98-106">Element</span></span> |  <span data-ttu-id="8df98-107">必須</span><span class="sxs-lookup"><span data-stu-id="8df98-107">Required</span></span>  |  <span data-ttu-id="8df98-108">説明</span><span class="sxs-lookup"><span data-stu-id="8df98-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8df98-109">Title</span><span class="sxs-lookup"><span data-stu-id="8df98-109">Title</span></span>](#title)        | <span data-ttu-id="8df98-110">はい</span><span class="sxs-lookup"><span data-stu-id="8df98-110">Yes</span></span> |   <span data-ttu-id="8df98-111">ヒントのテキストです。</span><span class="sxs-lookup"><span data-stu-id="8df98-111">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="8df98-112">説明</span><span class="sxs-lookup"><span data-stu-id="8df98-112">Description</span></span>](#description)  | <span data-ttu-id="8df98-113">はい</span><span class="sxs-lookup"><span data-stu-id="8df98-113">Yes</span></span> |  <span data-ttu-id="8df98-114">ヒントの説明です。</span><span class="sxs-lookup"><span data-stu-id="8df98-114">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="8df98-115">タイトル</span><span class="sxs-lookup"><span data-stu-id="8df98-115">Title</span></span>

<span data-ttu-id="8df98-p102">必ず指定します。ヒントのテキストです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8df98-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="8df98-119">説明</span><span class="sxs-lookup"><span data-stu-id="8df98-119">Description</span></span>

<span data-ttu-id="8df98-p103">必ず指定します。ヒントの記述です。 **resid** 属性には、 **Resources** 要素の **LongStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8df98-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="8df98-123">例</span><span class="sxs-lookup"><span data-stu-id="8df98-123">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
