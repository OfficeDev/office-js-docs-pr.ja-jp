---
title: マニフェスト ファイルの PhoneSettings 要素
description: Phone の Settings 要素は、メールアドインが電話で使用されるときに適用されるソースの場所と制御の設定を指定します。
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 581a3ae71a58cd05aac52129a6f4395a60c20cef
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720478"
---
# <a name="phonesettings-element"></a><span data-ttu-id="deec8-103">PhoneSettings 要素</span><span class="sxs-lookup"><span data-stu-id="deec8-103">PhoneSettings element</span></span>

<span data-ttu-id="deec8-104">メール アドインが電話で使用されるときに適用されるソースの場所と制御の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="deec8-104">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="deec8-105">この`PhoneSettings`要素は、web 上の従来の Outlook (社内 Exchange サーバーの古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="deec8-105">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="deec8-106">Android および iOS で Outlook をサポートするには、「 [Outlook Mobile 用のアドイン](../../outlook/outlook-mobile-addins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="deec8-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="deec8-107">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="deec8-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="deec8-108">構文</span><span class="sxs-lookup"><span data-stu-id="deec8-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="deec8-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="deec8-109">Contained in</span></span>

[<span data-ttu-id="deec8-110">Form</span><span class="sxs-lookup"><span data-stu-id="deec8-110">Form</span></span>](form.md)

