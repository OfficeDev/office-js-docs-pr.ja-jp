---
title: マニフェスト ファイルの SourceLocation 要素
description: SourceLocation 要素は、Office アドインのソースファイルの場所を指定します。
ms.date: 05/12/2020
localization_priority: Normal
ms.openlocfilehash: 447adb7df7d0c59305fe5046357959fcd7824735
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641404"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="f7a53-103">SourceLocation 要素</span><span class="sxs-lookup"><span data-stu-id="f7a53-103">SourceLocation element</span></span>

<span data-ttu-id="f7a53-104">Office アドインのソースファイルの場所を、1 ~ 2018 文字の長さの URL として指定します。</span><span class="sxs-lookup"><span data-stu-id="f7a53-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="f7a53-105">ソースの場所はファイル パスではなく、HTTPS アドレスにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f7a53-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="f7a53-106">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="f7a53-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f7a53-107">構文</span><span class="sxs-lookup"><span data-stu-id="f7a53-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="f7a53-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="f7a53-108">Contained in</span></span>

- <span data-ttu-id="f7a53-109">[DefaultSettings](defaultsettings.md) (コンテンツ アドインおよび作業ウィンドウ アドイン)</span><span class="sxs-lookup"><span data-stu-id="f7a53-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="f7a53-110">[FormSettings](formsettings.md) (メール アドイン)</span><span class="sxs-lookup"><span data-stu-id="f7a53-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="f7a53-111">[Extensionpoint](extensionpoint.md) (コンテキストおよび launchevent (プレビュー) メールアドイン)</span><span class="sxs-lookup"><span data-stu-id="f7a53-111">[ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent (preview) mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="f7a53-112">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="f7a53-112">Can contain</span></span>

[<span data-ttu-id="f7a53-113">Override</span><span class="sxs-lookup"><span data-stu-id="f7a53-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="f7a53-114">属性</span><span class="sxs-lookup"><span data-stu-id="f7a53-114">Attributes</span></span>

|<span data-ttu-id="f7a53-115">属性</span><span class="sxs-lookup"><span data-stu-id="f7a53-115">Attribute</span></span>|<span data-ttu-id="f7a53-116">型</span><span class="sxs-lookup"><span data-stu-id="f7a53-116">Type</span></span>|<span data-ttu-id="f7a53-117">必須</span><span class="sxs-lookup"><span data-stu-id="f7a53-117">Required</span></span>|<span data-ttu-id="f7a53-118">説明</span><span class="sxs-lookup"><span data-stu-id="f7a53-118">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f7a53-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="f7a53-119">DefaultValue</span></span>|<span data-ttu-id="f7a53-120">URL</span><span class="sxs-lookup"><span data-stu-id="f7a53-120">URL</span></span>|<span data-ttu-id="f7a53-121">必須</span><span class="sxs-lookup"><span data-stu-id="f7a53-121">required</span></span>|<span data-ttu-id="f7a53-122">[DefaultLocale](defaultlocale.md) 要素に指定されるロケール用に、この設定の既定値を指定します。</span><span class="sxs-lookup"><span data-stu-id="f7a53-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
