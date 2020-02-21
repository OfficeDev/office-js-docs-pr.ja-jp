---
title: マニフェスト ファイルの PhoneSettings 要素
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 4614c86af865e5242657f47e21e6786545a616b6
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165539"
---
# <a name="phonesettings-element"></a><span data-ttu-id="b6201-102">PhoneSettings 要素</span><span class="sxs-lookup"><span data-stu-id="b6201-102">PhoneSettings element</span></span>

<span data-ttu-id="b6201-103">メール アドインが電話で使用されるときに適用されるソースの場所と制御の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="b6201-103">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b6201-104">この`PhoneSettings`要素は、web 上の従来の Outlook (社内 Exchange サーバーの古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="b6201-104">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="b6201-105">Android および iOS で Outlook をサポートするには、「 [Outlook Mobile 用のアドイン](../../outlook/outlook-mobile-addins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b6201-105">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="b6201-106">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="b6201-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b6201-107">構文</span><span class="sxs-lookup"><span data-stu-id="b6201-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="b6201-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="b6201-108">Contained in</span></span>

[<span data-ttu-id="b6201-109">Form</span><span class="sxs-lookup"><span data-stu-id="b6201-109">Form</span></span>](form.md)

