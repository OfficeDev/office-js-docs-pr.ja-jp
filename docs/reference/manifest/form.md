---
title: マニフェスト ファイルの Form 要素
description: メール アドインが特定のデバイス (デスクトップ、タブレット、または電話) で実行されているときに使用するフォームの UX の設定。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 3e8d60c13a72a50090075d7cd16a0719498c4982
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215069"
---
# <a name="form-element"></a><span data-ttu-id="116c7-103">Form 要素</span><span class="sxs-lookup"><span data-stu-id="116c7-103">Form element</span></span>

<span data-ttu-id="116c7-104">メール アドインが特定のデバイス (デスクトップ、タブレット、または電話) で実行されているときに使用するフォームの UX の設定。</span><span class="sxs-lookup"><span data-stu-id="116c7-104">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="116c7-105">、、および`PhoneSettings`要素は、従来の web 上の outlook (通常は社内の Exchange server の古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。 `DesktopSettings` `TabletSettings`</span><span class="sxs-lookup"><span data-stu-id="116c7-105">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="116c7-106">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="116c7-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="116c7-107">構文</span><span class="sxs-lookup"><span data-stu-id="116c7-107">Syntax</span></span>

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a><span data-ttu-id="116c7-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="116c7-108">Contained in</span></span>

[<span data-ttu-id="116c7-109">FormSettings</span><span class="sxs-lookup"><span data-stu-id="116c7-109">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="116c7-110">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="116c7-110">Can contain</span></span>

|<span data-ttu-id="116c7-111">**Element**</span><span class="sxs-lookup"><span data-stu-id="116c7-111">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="116c7-112">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="116c7-112">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="116c7-113">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="116c7-113">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="116c7-114">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="116c7-114">PhoneSettings</span></span>](phonesettings.md)|
