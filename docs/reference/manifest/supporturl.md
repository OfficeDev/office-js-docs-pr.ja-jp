---
title: マニフェスト ファイルの SupportUrl 要素
description: SupportUrl 要素は、アドインのサポート情報を提供するページの URL を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: be516fe5848d775dacb0d424a92be02d59f85512
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641411"
---
# <a name="supporturl-element"></a><span data-ttu-id="4e880-103">SupportUrl 要素</span><span class="sxs-lookup"><span data-stu-id="4e880-103">SupportUrl element</span></span>

<span data-ttu-id="4e880-104">アドインのサポート情報を提供するページの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="4e880-104">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="4e880-105">構文</span><span class="sxs-lookup"><span data-stu-id="4e880-105">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="4e880-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="4e880-106">Contained in</span></span>

[<span data-ttu-id="4e880-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="4e880-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="4e880-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="4e880-108">Can contain</span></span>

|  <span data-ttu-id="4e880-109">要素</span><span class="sxs-lookup"><span data-stu-id="4e880-109">Element</span></span> | <span data-ttu-id="4e880-110">必須</span><span class="sxs-lookup"><span data-stu-id="4e880-110">Required</span></span> | <span data-ttu-id="4e880-111">説明</span><span class="sxs-lookup"><span data-stu-id="4e880-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4e880-112">Override</span><span class="sxs-lookup"><span data-stu-id="4e880-112">Override</span></span>](override.md)   | <span data-ttu-id="4e880-113">なし</span><span class="sxs-lookup"><span data-stu-id="4e880-113">No</span></span> | <span data-ttu-id="4e880-114">追加のロケール URL の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="4e880-114">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="4e880-115">属性</span><span class="sxs-lookup"><span data-stu-id="4e880-115">Attributes</span></span>

|<span data-ttu-id="4e880-116">属性</span><span class="sxs-lookup"><span data-stu-id="4e880-116">Attribute</span></span>|<span data-ttu-id="4e880-117">型</span><span class="sxs-lookup"><span data-stu-id="4e880-117">Type</span></span>|<span data-ttu-id="4e880-118">必須</span><span class="sxs-lookup"><span data-stu-id="4e880-118">Required</span></span>|<span data-ttu-id="4e880-119">説明</span><span class="sxs-lookup"><span data-stu-id="4e880-119">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="4e880-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="4e880-120">DefaultValue</span></span>|<span data-ttu-id="4e880-121">URL</span><span class="sxs-lookup"><span data-stu-id="4e880-121">URL</span></span>|<span data-ttu-id="4e880-122">必須</span><span class="sxs-lookup"><span data-stu-id="4e880-122">required</span></span>|<span data-ttu-id="4e880-123">この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。</span><span class="sxs-lookup"><span data-stu-id="4e880-123">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
