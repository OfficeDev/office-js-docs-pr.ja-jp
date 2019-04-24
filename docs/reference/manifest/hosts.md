---
title: マニフェスト ファイルの Hosts 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 606073977366e37ecc4419f468f01bfb25647a7d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452026"
---
# <a name="hosts-element"></a><span data-ttu-id="b0e20-102">Hosts 要素</span><span class="sxs-lookup"><span data-stu-id="b0e20-102">Hosts element</span></span>

<span data-ttu-id="b0e20-p101">Office アドインをアクティブにする Office クライアント アプリケーションを指定します。 **Host** 要素のコレクションとその設定が含まれます。</span><span class="sxs-lookup"><span data-stu-id="b0e20-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="b0e20-105">[VersionOverrides](versionoverrides.md) ノードに含まれる場合、この要素は、マニフェストの親部分の **Hosts** 要素よりも優先されます。</span><span class="sxs-lookup"><span data-stu-id="b0e20-105">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="b0e20-106">子要素</span><span class="sxs-lookup"><span data-stu-id="b0e20-106">Child elements</span></span>

|  <span data-ttu-id="b0e20-107">要素</span><span class="sxs-lookup"><span data-stu-id="b0e20-107">Element</span></span> |  <span data-ttu-id="b0e20-108">必須</span><span class="sxs-lookup"><span data-stu-id="b0e20-108">Required</span></span>  |  <span data-ttu-id="b0e20-109">説明</span><span class="sxs-lookup"><span data-stu-id="b0e20-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b0e20-110">Host</span><span class="sxs-lookup"><span data-stu-id="b0e20-110">Host</span></span>](host.md)    |  <span data-ttu-id="b0e20-111">はい</span><span class="sxs-lookup"><span data-stu-id="b0e20-111">Yes</span></span>   |  <span data-ttu-id="b0e20-112">ホストとその設定について説明します。</span><span class="sxs-lookup"><span data-stu-id="b0e20-112">Describes a host and its settings.</span></span> |
