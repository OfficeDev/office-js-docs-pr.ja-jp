---
title: マニフェスト ファイルの Version 要素
description: Version 要素は、アドインOfficeバージョンを指定します。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 48a2be94d95ece597e47468bb18db2a7962a51e9
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173935"
---
# <a name="version-element"></a><span data-ttu-id="33aa9-103">Version 要素</span><span class="sxs-lookup"><span data-stu-id="33aa9-103">Version element</span></span>

<span data-ttu-id="33aa9-104">Office アドインのバージョンを指定します。</span><span class="sxs-lookup"><span data-stu-id="33aa9-104">Specifies the version of your Office Add-in.</span></span> <span data-ttu-id="33aa9-105">バージョン番号は、1、2、3、または 4 つの部分 (つまり、n、n.n、n.n.n、または n.n.n.n) です。</span><span class="sxs-lookup"><span data-stu-id="33aa9-105">The version number can be 1, 2, 3, or 4 parts (i.e., n, n.n, n.n.n, or n.n.n.n).</span></span>

<span data-ttu-id="33aa9-106">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="33aa9-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="33aa9-107">構文</span><span class="sxs-lookup"><span data-stu-id="33aa9-107">Syntax</span></span>

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a><span data-ttu-id="33aa9-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="33aa9-108">Contained in</span></span>

[<span data-ttu-id="33aa9-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="33aa9-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="33aa9-110">注釈</span><span class="sxs-lookup"><span data-stu-id="33aa9-110">Remarks</span></span>

<span data-ttu-id="33aa9-111">バージョン番号の各部分には、最大 5 桁の数字を指定できます。</span><span class="sxs-lookup"><span data-stu-id="33aa9-111">Each part of the version number can be a maximum of 5 digits.</span></span>
