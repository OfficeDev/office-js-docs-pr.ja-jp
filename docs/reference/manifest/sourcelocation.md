---
title: マニフェスト ファイルの SourceLocation 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dc432ebb9482e8e9b8be5d90a838357ccf519ad3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433517"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="f0c16-102">SourceLocation 要素</span><span class="sxs-lookup"><span data-stu-id="f0c16-102">SourceLocation element</span></span>

<span data-ttu-id="f0c16-p101">Office アドインのソース ファイルの場所を、1 から 2018 文字までの長さの URL として指定します。ソースの場所はファイル パスではなく、HTTPS アドレスにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f0c16-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="f0c16-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="f0c16-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f0c16-106">構文</span><span class="sxs-lookup"><span data-stu-id="f0c16-106">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="f0c16-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="f0c16-107">Contained in</span></span>

- <span data-ttu-id="f0c16-108">[DefaultSettings](defaultsettings.md) (コンテンツ アドインおよび作業ウィンドウ アドイン)</span><span class="sxs-lookup"><span data-stu-id="f0c16-108">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="f0c16-109">[FormSettings](formsettings.md) (メール アドイン)</span><span class="sxs-lookup"><span data-stu-id="f0c16-109">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="f0c16-110">[ExtensionPoint](extensionpoint.md) (コンテキスト メール アドイン)</span><span class="sxs-lookup"><span data-stu-id="f0c16-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="f0c16-111">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="f0c16-111">Can contain</span></span>

[<span data-ttu-id="f0c16-112">Override</span><span class="sxs-lookup"><span data-stu-id="f0c16-112">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="f0c16-113">属性</span><span class="sxs-lookup"><span data-stu-id="f0c16-113">Attributes</span></span>

|<span data-ttu-id="f0c16-114">**属性**</span><span class="sxs-lookup"><span data-stu-id="f0c16-114">**Attribute**</span></span>|<span data-ttu-id="f0c16-115">**型**</span><span class="sxs-lookup"><span data-stu-id="f0c16-115">**Type**</span></span>|<span data-ttu-id="f0c16-116">**必須**</span><span class="sxs-lookup"><span data-stu-id="f0c16-116">**Required**</span></span>|<span data-ttu-id="f0c16-117">**説明**</span><span class="sxs-lookup"><span data-stu-id="f0c16-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f0c16-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="f0c16-118">DefaultValue</span></span>|<span data-ttu-id="f0c16-119">URL</span><span class="sxs-lookup"><span data-stu-id="f0c16-119">URL</span></span>|<span data-ttu-id="f0c16-120">必須</span><span class="sxs-lookup"><span data-stu-id="f0c16-120">required</span></span>|<span data-ttu-id="f0c16-121">[DefaultLocale](defaultlocale.md) 要素に指定されるロケール用に、この設定の既定値を指定します。</span><span class="sxs-lookup"><span data-stu-id="f0c16-121">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
