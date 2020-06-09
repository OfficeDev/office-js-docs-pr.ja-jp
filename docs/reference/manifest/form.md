---
title: マニフェスト ファイルの Form 要素
description: メール アドインが特定のデバイス (デスクトップ、タブレット、または電話) で実行されているときに使用するフォームの UX の設定。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: c9cd1d9104fc51edc84149ef677c4308dfb1a9f5
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611856"
---
# <a name="form-element"></a><span data-ttu-id="bd362-103">Form 要素</span><span class="sxs-lookup"><span data-stu-id="bd362-103">Form element</span></span>

<span data-ttu-id="bd362-104">メール アドインが特定のデバイス (デスクトップ、タブレット、または電話) で実行されているときに使用するフォームの UX の設定。</span><span class="sxs-lookup"><span data-stu-id="bd362-104">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bd362-105">、 `DesktopSettings` 、 `TabletSettings` および要素は、 `PhoneSettings` 従来の Web 上の outlook (通常は社内の Exchange server の古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="bd362-105">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="bd362-106">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="bd362-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bd362-107">構文</span><span class="sxs-lookup"><span data-stu-id="bd362-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="bd362-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="bd362-108">Contained in</span></span>

[<span data-ttu-id="bd362-109">FormSettings</span><span class="sxs-lookup"><span data-stu-id="bd362-109">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="bd362-110">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="bd362-110">Can contain</span></span>

|<span data-ttu-id="bd362-111">**Element**</span><span class="sxs-lookup"><span data-stu-id="bd362-111">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="bd362-112">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="bd362-112">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="bd362-113">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="bd362-113">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="bd362-114">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="bd362-114">PhoneSettings</span></span>](phonesettings.md)|
