---
title: マニフェスト ファイルの TabletSettings 要素
description: TabletSettings 要素は、メールアドインがタブレットで使用されるときに適用する制御の設定を指定します。
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 2b8b372d27274d89d3aed4b5bacb9faa4893fda5
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717860"
---
# <a name="tabletsettings-element"></a><span data-ttu-id="20170-103">TabletSettings 要素</span><span class="sxs-lookup"><span data-stu-id="20170-103">TabletSettings element</span></span>

<span data-ttu-id="20170-104">メール アドインがタブレットで使用されるときに適用される制御の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="20170-104">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="20170-105">この`TabletSettings`要素は、web 上の従来の Outlook (社内 Exchange サーバーの古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="20170-105">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="20170-106">Android および iOS で Outlook をサポートするには、「 [Outlook Mobile 用のアドイン](../../outlook/outlook-mobile-addins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="20170-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="20170-107">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="20170-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="20170-108">構文</span><span class="sxs-lookup"><span data-stu-id="20170-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="20170-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="20170-109">Contained in</span></span>

[<span data-ttu-id="20170-110">Form</span><span class="sxs-lookup"><span data-stu-id="20170-110">Form</span></span>](form.md)

