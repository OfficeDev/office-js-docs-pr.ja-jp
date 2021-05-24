---
title: マニフェスト ファイルの SourceLocation 要素
description: SourceLocation 要素は、アドインのソース ファイルOffice指定します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 4dcd093db2f23220eaa34c0c81300c4994c1a697
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590898"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="29ff6-103">SourceLocation 要素</span><span class="sxs-lookup"><span data-stu-id="29ff6-103">SourceLocation element</span></span>

<span data-ttu-id="29ff6-104">1 ~ 2018 文字の URL として、Officeアドインのソース ファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="29ff6-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="29ff6-105">ソースの場所はファイル パスではなく、HTTPS アドレスにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="29ff6-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="29ff6-106">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="29ff6-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="29ff6-107">構文</span><span class="sxs-lookup"><span data-stu-id="29ff6-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="29ff6-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="29ff6-108">Contained in</span></span>

- <span data-ttu-id="29ff6-109">[DefaultSettings](defaultsettings.md) (コンテンツ アドインおよび作業ウィンドウ アドイン)</span><span class="sxs-lookup"><span data-stu-id="29ff6-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="29ff6-110">[FormSettings](formsettings.md) (メール アドイン)</span><span class="sxs-lookup"><span data-stu-id="29ff6-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="29ff6-111">[ExtensionPoint](extensionpoint.md) (コンテキスト メール アドインと LaunchEvent メール アドイン)</span><span class="sxs-lookup"><span data-stu-id="29ff6-111">[ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="29ff6-112">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="29ff6-112">Can contain</span></span>

[<span data-ttu-id="29ff6-113">Override</span><span class="sxs-lookup"><span data-stu-id="29ff6-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="29ff6-114">属性</span><span class="sxs-lookup"><span data-stu-id="29ff6-114">Attributes</span></span>

|<span data-ttu-id="29ff6-115">属性</span><span class="sxs-lookup"><span data-stu-id="29ff6-115">Attribute</span></span>|<span data-ttu-id="29ff6-116">型</span><span class="sxs-lookup"><span data-stu-id="29ff6-116">Type</span></span>|<span data-ttu-id="29ff6-117">必須</span><span class="sxs-lookup"><span data-stu-id="29ff6-117">Required</span></span>|<span data-ttu-id="29ff6-118">説明</span><span class="sxs-lookup"><span data-stu-id="29ff6-118">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="29ff6-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="29ff6-119">DefaultValue</span></span>|<span data-ttu-id="29ff6-120">URL</span><span class="sxs-lookup"><span data-stu-id="29ff6-120">URL</span></span>|<span data-ttu-id="29ff6-121">必須</span><span class="sxs-lookup"><span data-stu-id="29ff6-121">required</span></span>|<span data-ttu-id="29ff6-122">[DefaultLocale](defaultlocale.md) 要素に指定されるロケール用に、この設定の既定値を指定します。</span><span class="sxs-lookup"><span data-stu-id="29ff6-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
