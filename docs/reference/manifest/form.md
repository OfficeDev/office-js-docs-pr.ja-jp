---
title: マニフェスト ファイルの Form 要素
description: メール アドインが特定のデバイス (デスクトップ、タブレット、または電話) で実行されているときに使用するフォームの UX の設定。
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 9b1696b2fecf6b07ee2a3c0a31611d4f2ad1f291
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718210"
---
# <a name="form-element"></a><span data-ttu-id="0bdc7-103">Form 要素</span><span class="sxs-lookup"><span data-stu-id="0bdc7-103">Form element</span></span>

<span data-ttu-id="0bdc7-104">メール アドインが特定のデバイス (デスクトップ、タブレット、または電話) で実行されているときに使用するフォームの UX の設定。</span><span class="sxs-lookup"><span data-stu-id="0bdc7-104">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0bdc7-105">、、および`PhoneSettings`要素は、従来の web 上の outlook (通常は社内の Exchange server の古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。 `DesktopSettings` `TabletSettings`</span><span class="sxs-lookup"><span data-stu-id="0bdc7-105">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="0bdc7-106">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="0bdc7-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0bdc7-107">構文</span><span class="sxs-lookup"><span data-stu-id="0bdc7-107">Syntax</span></span>

```XML
<Form xsi:type="ItemRead">
   <!--website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </DesktopSettings>
   <TabletSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a><span data-ttu-id="0bdc7-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="0bdc7-108">Contained in</span></span>

[<span data-ttu-id="0bdc7-109">FormSettings</span><span class="sxs-lookup"><span data-stu-id="0bdc7-109">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="0bdc7-110">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="0bdc7-110">Can contain</span></span>

|<span data-ttu-id="0bdc7-111">**Element**</span><span class="sxs-lookup"><span data-stu-id="0bdc7-111">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="0bdc7-112">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="0bdc7-112">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="0bdc7-113">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="0bdc7-113">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="0bdc7-114">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="0bdc7-114">PhoneSettings</span></span>](phonesettings.md)|
