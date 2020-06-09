---
title: マニフェスト ファイルの Hosts 要素
description: Office アドインをアクティブにする Office クライアント アプリケーションを指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 037ac2b5fedbfb1b59b7523382574942fe59a00a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611807"
---
# <a name="hosts-element"></a><span data-ttu-id="86af6-103">Hosts 要素</span><span class="sxs-lookup"><span data-stu-id="86af6-103">Hosts element</span></span>

<span data-ttu-id="86af6-p101">Office アドインをアクティブにする Office クライアント アプリケーションを指定します。 **Host** 要素のコレクションとその設定が含まれます。</span><span class="sxs-lookup"><span data-stu-id="86af6-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="86af6-106">[VersionOverrides](versionoverrides.md) ノードに含まれる場合、この要素は、マニフェストの親部分の **Hosts** 要素よりも優先されます。</span><span class="sxs-lookup"><span data-stu-id="86af6-106">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="86af6-107">子要素</span><span class="sxs-lookup"><span data-stu-id="86af6-107">Child elements</span></span>

|  <span data-ttu-id="86af6-108">要素</span><span class="sxs-lookup"><span data-stu-id="86af6-108">Element</span></span> |  <span data-ttu-id="86af6-109">必須</span><span class="sxs-lookup"><span data-stu-id="86af6-109">Required</span></span>  |  <span data-ttu-id="86af6-110">説明</span><span class="sxs-lookup"><span data-stu-id="86af6-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="86af6-111">Host</span><span class="sxs-lookup"><span data-stu-id="86af6-111">Host</span></span>](host.md)    |  <span data-ttu-id="86af6-112">はい</span><span class="sxs-lookup"><span data-stu-id="86af6-112">Yes</span></span>   |  <span data-ttu-id="86af6-113">ホストとその設定について説明します。</span><span class="sxs-lookup"><span data-stu-id="86af6-113">Describes a host and its settings.</span></span> |
