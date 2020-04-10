---
title: マニフェスト ファイルの PhoneSettings 要素
description: Phone の Settings 要素は、メールアドインが電話で使用されるときに適用されるソースの場所と制御の設定を指定します。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: cfdf8efc262c1a75142eb06dd71383579598a86f
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215048"
---
# <a name="phonesettings-element"></a><span data-ttu-id="7cd87-103">PhoneSettings 要素</span><span class="sxs-lookup"><span data-stu-id="7cd87-103">PhoneSettings element</span></span>

<span data-ttu-id="7cd87-104">メール アドインが電話で使用されるときに適用されるソースの場所と制御の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="7cd87-104">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7cd87-105">この`PhoneSettings`要素は、web 上の従来の Outlook (社内 Exchange サーバーの古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="7cd87-105">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="7cd87-106">Android および iOS で Outlook をサポートするには、「 [Outlook Mobile 用のアドイン](../../outlook/outlook-mobile-addins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7cd87-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="7cd87-107">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="7cd87-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="7cd87-108">構文</span><span class="sxs-lookup"><span data-stu-id="7cd87-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="7cd87-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="7cd87-109">Contained in</span></span>

[<span data-ttu-id="7cd87-110">Form</span><span class="sxs-lookup"><span data-stu-id="7cd87-110">Form</span></span>](form.md)

