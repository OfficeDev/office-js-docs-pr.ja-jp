---
title: マニフェスト ファイルの SupportUrl 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 18b9b7c4df9def70ab42ae213066188ac04c07a7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450416"
---
# <a name="supporturl-element"></a><span data-ttu-id="5915c-102">SupportUrl 要素</span><span class="sxs-lookup"><span data-stu-id="5915c-102">SupportUrl element</span></span>

<span data-ttu-id="5915c-103">アドインのサポート情報を提供するページの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="5915c-103">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="5915c-104">構文</span><span class="sxs-lookup"><span data-stu-id="5915c-104">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="5915c-105">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="5915c-105">Contained in</span></span>

[<span data-ttu-id="5915c-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="5915c-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="5915c-107">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="5915c-107">Can contain</span></span>

|  <span data-ttu-id="5915c-108">要素</span><span class="sxs-lookup"><span data-stu-id="5915c-108">Element</span></span> | <span data-ttu-id="5915c-109">必須</span><span class="sxs-lookup"><span data-stu-id="5915c-109">Required</span></span> | <span data-ttu-id="5915c-110">説明</span><span class="sxs-lookup"><span data-stu-id="5915c-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5915c-111">Override</span><span class="sxs-lookup"><span data-stu-id="5915c-111">Override</span></span>](override.md)   | <span data-ttu-id="5915c-112">なし</span><span class="sxs-lookup"><span data-stu-id="5915c-112">No</span></span> | <span data-ttu-id="5915c-113">追加のロケール URL の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="5915c-113">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="5915c-114">属性</span><span class="sxs-lookup"><span data-stu-id="5915c-114">Attributes</span></span>

|<span data-ttu-id="5915c-115">**属性**</span><span class="sxs-lookup"><span data-stu-id="5915c-115">**Attribute**</span></span>|<span data-ttu-id="5915c-116">**型**</span><span class="sxs-lookup"><span data-stu-id="5915c-116">**Type**</span></span>|<span data-ttu-id="5915c-117">**必須**</span><span class="sxs-lookup"><span data-stu-id="5915c-117">**Required**</span></span>|<span data-ttu-id="5915c-118">**説明**</span><span class="sxs-lookup"><span data-stu-id="5915c-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5915c-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="5915c-119">DefaultValue</span></span>|<span data-ttu-id="5915c-120">URL</span><span class="sxs-lookup"><span data-stu-id="5915c-120">URL</span></span>|<span data-ttu-id="5915c-121">必須</span><span class="sxs-lookup"><span data-stu-id="5915c-121">required</span></span>|<span data-ttu-id="5915c-122">この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。</span><span class="sxs-lookup"><span data-stu-id="5915c-122">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
