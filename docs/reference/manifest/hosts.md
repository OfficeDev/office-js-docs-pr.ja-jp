---
title: マニフェスト ファイルの Hosts 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 59010c0f6c0d14d8721856f81def11540db28704
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433412"
---
# <a name="hosts-element"></a><span data-ttu-id="ed8a8-102">Hosts 要素</span><span class="sxs-lookup"><span data-stu-id="ed8a8-102">Hosts element</span></span>

<span data-ttu-id="ed8a8-p101">Office アドインをアクティブにする Office クライアント アプリケーションを指定します。**Host** 要素のコレクションとその設定が含まれます。</span><span class="sxs-lookup"><span data-stu-id="ed8a8-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="ed8a8-105">[VersionOverrides](versionoverrides.md) ノードに含まれる場合、この要素は、マニフェストの親部分の **Hosts** 要素よりも優先されます。</span><span class="sxs-lookup"><span data-stu-id="ed8a8-105">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="ed8a8-106">子要素</span><span class="sxs-lookup"><span data-stu-id="ed8a8-106">Child elements</span></span>

|  <span data-ttu-id="ed8a8-107">要素</span><span class="sxs-lookup"><span data-stu-id="ed8a8-107">Element</span></span> |  <span data-ttu-id="ed8a8-108">必須</span><span class="sxs-lookup"><span data-stu-id="ed8a8-108">Required</span></span>  |  <span data-ttu-id="ed8a8-109">説明</span><span class="sxs-lookup"><span data-stu-id="ed8a8-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ed8a8-110">Host</span><span class="sxs-lookup"><span data-stu-id="ed8a8-110">Host</span></span>](host.md)    |  <span data-ttu-id="ed8a8-111">はい</span><span class="sxs-lookup"><span data-stu-id="ed8a8-111">Yes</span></span>   |  <span data-ttu-id="ed8a8-112">ホストとその設定について説明します。</span><span class="sxs-lookup"><span data-stu-id="ed8a8-112">Describes a host and its settings.</span></span> |
