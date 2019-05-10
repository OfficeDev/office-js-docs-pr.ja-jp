---
title: マニフェスト ファイルの Supertip 要素
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 269a3723db6f98cdb25c61e5a88608c5fb5f3191
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659657"
---
# <a name="supertip"></a><span data-ttu-id="5c1ea-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="5c1ea-102">Supertip</span></span>

<span data-ttu-id="5c1ea-p101">豊富なヒント (タイトルと説明の両方) を定義します。[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの両方で使用されます。</span><span class="sxs-lookup"><span data-stu-id="5c1ea-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5c1ea-105">子要素</span><span class="sxs-lookup"><span data-stu-id="5c1ea-105">Child elements</span></span>

|  <span data-ttu-id="5c1ea-106">要素</span><span class="sxs-lookup"><span data-stu-id="5c1ea-106">Element</span></span> |  <span data-ttu-id="5c1ea-107">必須</span><span class="sxs-lookup"><span data-stu-id="5c1ea-107">Required</span></span>  |  <span data-ttu-id="5c1ea-108">説明</span><span class="sxs-lookup"><span data-stu-id="5c1ea-108">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="5c1ea-109">Title</span><span class="sxs-lookup"><span data-stu-id="5c1ea-109">Title</span></span>](#title) | <span data-ttu-id="5c1ea-110">はい</span><span class="sxs-lookup"><span data-stu-id="5c1ea-110">Yes</span></span> | <span data-ttu-id="5c1ea-111">ヒントのテキストです。</span><span class="sxs-lookup"><span data-stu-id="5c1ea-111">The text for the supertip.</span></span> |
| [<span data-ttu-id="5c1ea-112">説明</span><span class="sxs-lookup"><span data-stu-id="5c1ea-112">Description</span></span>](#description) | <span data-ttu-id="5c1ea-113">はい</span><span class="sxs-lookup"><span data-stu-id="5c1ea-113">Yes</span></span> | <span data-ttu-id="5c1ea-114">ヒントの説明です。</span><span class="sxs-lookup"><span data-stu-id="5c1ea-114">The description for the supertip.</span></span><br><span data-ttu-id="5c1ea-115">**注**: (Outlook) は、Windows および Mac クライアントのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="5c1ea-115">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="5c1ea-116">タイトル</span><span class="sxs-lookup"><span data-stu-id="5c1ea-116">Title</span></span>

<span data-ttu-id="5c1ea-p102">必ず指定します。ヒントのテキストです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5c1ea-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="5c1ea-120">説明</span><span class="sxs-lookup"><span data-stu-id="5c1ea-120">Description</span></span>

<span data-ttu-id="5c1ea-p103">必ず指定します。ヒントの記述です。 **resid** 属性には、 **Resources** 要素の **LongStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5c1ea-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="5c1ea-124">Outlook の場合、Windows と Mac のクライアントのみが**Description**要素をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="5c1ea-124">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="5c1ea-125">例</span><span class="sxs-lookup"><span data-stu-id="5c1ea-125">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
