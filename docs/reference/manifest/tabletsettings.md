---
title: マニフェスト ファイルの TabletSettings 要素
description: TabletSettings 要素は、メールアドインがタブレットで使用されるときに適用する制御の設定を指定します。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: b5a74db4f9fb43df10a08ab43b59507f6e0d7952
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608699"
---
# <a name="tabletsettings-element"></a><span data-ttu-id="fede6-103">TabletSettings 要素</span><span class="sxs-lookup"><span data-stu-id="fede6-103">TabletSettings element</span></span>

<span data-ttu-id="fede6-104">メール アドインがタブレットで使用されるときに適用される制御の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="fede6-104">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fede6-105">この `TabletSettings` 要素は、web 上の従来の Outlook (社内 Exchange サーバーの古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="fede6-105">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="fede6-106">Android および iOS で Outlook をサポートするには、「 [Outlook Mobile 用のアドイン](../../outlook/outlook-mobile-addins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fede6-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="fede6-107">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="fede6-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="fede6-108">構文</span><span class="sxs-lookup"><span data-stu-id="fede6-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="fede6-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="fede6-109">Contained in</span></span>

[<span data-ttu-id="fede6-110">Form</span><span class="sxs-lookup"><span data-stu-id="fede6-110">Form</span></span>](form.md)
