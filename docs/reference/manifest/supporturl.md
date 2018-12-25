---
title: マニフェスト ファイルの SupportUrl 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 00234ef9fe8960b9956e6a2595e2e2e71bfb97c6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432670"
---
# <a name="supporturl-element"></a><span data-ttu-id="d2b59-102">SupportUrl 要素</span><span class="sxs-lookup"><span data-stu-id="d2b59-102">SupportUrl element</span></span>

<span data-ttu-id="d2b59-103">アドインのサポート情報を提供するページの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="d2b59-103">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="d2b59-104">構文</span><span class="sxs-lookup"><span data-stu-id="d2b59-104">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="d2b59-105">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="d2b59-105">Contained in</span></span>

[<span data-ttu-id="d2b59-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="d2b59-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="d2b59-107">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="d2b59-107">Can contain</span></span>

|  <span data-ttu-id="d2b59-108">要素</span><span class="sxs-lookup"><span data-stu-id="d2b59-108">Element</span></span> | <span data-ttu-id="d2b59-109">必須</span><span class="sxs-lookup"><span data-stu-id="d2b59-109">Required</span></span> | <span data-ttu-id="d2b59-110">説明</span><span class="sxs-lookup"><span data-stu-id="d2b59-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d2b59-111">Override</span><span class="sxs-lookup"><span data-stu-id="d2b59-111">Override</span></span>](override.md)   | <span data-ttu-id="d2b59-112">なし</span><span class="sxs-lookup"><span data-stu-id="d2b59-112">No</span></span> | <span data-ttu-id="d2b59-113">追加のロケール URL の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="d2b59-113">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="d2b59-114">属性</span><span class="sxs-lookup"><span data-stu-id="d2b59-114">Attributes</span></span>

|<span data-ttu-id="d2b59-115">**属性**</span><span class="sxs-lookup"><span data-stu-id="d2b59-115">**Attribute**</span></span>|<span data-ttu-id="d2b59-116">**型**</span><span class="sxs-lookup"><span data-stu-id="d2b59-116">**Type**</span></span>|<span data-ttu-id="d2b59-117">**必須**</span><span class="sxs-lookup"><span data-stu-id="d2b59-117">**Required**</span></span>|<span data-ttu-id="d2b59-118">**説明**</span><span class="sxs-lookup"><span data-stu-id="d2b59-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d2b59-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="d2b59-119">DefaultValue</span></span>|<span data-ttu-id="d2b59-120">URL</span><span class="sxs-lookup"><span data-stu-id="d2b59-120">URL</span></span>|<span data-ttu-id="d2b59-121">必須</span><span class="sxs-lookup"><span data-stu-id="d2b59-121">required</span></span>|<span data-ttu-id="d2b59-122">この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。</span><span class="sxs-lookup"><span data-stu-id="d2b59-122">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
