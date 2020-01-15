---
title: マニフェスト ファイルの Form 要素
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: d545d471e007f0077a8310b0b847bbbf99a8f7ac
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120650"
---
# <a name="form-element"></a><span data-ttu-id="cba68-102">Form 要素</span><span class="sxs-lookup"><span data-stu-id="cba68-102">Form element</span></span>

<span data-ttu-id="cba68-103">メール アドインが特定のデバイス (デスクトップ、タブレット、または電話) で実行されているときに使用するフォームの UX の設定。</span><span class="sxs-lookup"><span data-stu-id="cba68-103">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cba68-104">、、および`PhoneSettings`要素は、従来の web 上の outlook (通常は社内の Exchange server の古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。 `DesktopSettings` `TabletSettings`</span><span class="sxs-lookup"><span data-stu-id="cba68-104">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="cba68-105">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="cba68-105">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cba68-106">構文</span><span class="sxs-lookup"><span data-stu-id="cba68-106">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="cba68-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="cba68-107">Contained in</span></span>

[<span data-ttu-id="cba68-108">FormSettings</span><span class="sxs-lookup"><span data-stu-id="cba68-108">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="cba68-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="cba68-109">Can contain</span></span>

|<span data-ttu-id="cba68-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="cba68-110">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="cba68-111">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="cba68-111">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="cba68-112">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="cba68-112">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="cba68-113">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="cba68-113">PhoneSettings</span></span>](phonesettings.md)|
