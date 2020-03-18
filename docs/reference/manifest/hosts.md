---
title: マニフェスト ファイルの Hosts 要素
description: Office アドインをアクティブにする Office クライアント アプリケーションを指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cd4e0eecce610b10fdc9dafcde7b807fde425b14
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718105"
---
# <a name="hosts-element"></a><span data-ttu-id="54716-103">Hosts 要素</span><span class="sxs-lookup"><span data-stu-id="54716-103">Hosts element</span></span>

<span data-ttu-id="54716-p101">Office アドインをアクティブにする Office クライアント アプリケーションを指定します。 **Host** 要素のコレクションとその設定が含まれます。</span><span class="sxs-lookup"><span data-stu-id="54716-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="54716-106">[VersionOverrides](versionoverrides.md) ノードに含まれる場合、この要素は、マニフェストの親部分の **Hosts** 要素よりも優先されます。</span><span class="sxs-lookup"><span data-stu-id="54716-106">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="54716-107">子要素</span><span class="sxs-lookup"><span data-stu-id="54716-107">Child elements</span></span>

|  <span data-ttu-id="54716-108">要素</span><span class="sxs-lookup"><span data-stu-id="54716-108">Element</span></span> |  <span data-ttu-id="54716-109">必須</span><span class="sxs-lookup"><span data-stu-id="54716-109">Required</span></span>  |  <span data-ttu-id="54716-110">説明</span><span class="sxs-lookup"><span data-stu-id="54716-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="54716-111">Host</span><span class="sxs-lookup"><span data-stu-id="54716-111">Host</span></span>](host.md)    |  <span data-ttu-id="54716-112">はい</span><span class="sxs-lookup"><span data-stu-id="54716-112">Yes</span></span>   |  <span data-ttu-id="54716-113">ホストとその設定について説明します。</span><span class="sxs-lookup"><span data-stu-id="54716-113">Describes a host and its settings.</span></span> |
