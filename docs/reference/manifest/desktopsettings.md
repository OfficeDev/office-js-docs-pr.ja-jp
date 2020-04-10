---
title: マニフェスト ファイルの DesktopSettings 要素
description: メール アドインがデスクトップ コンピューターで使用されるときに適用されるソースの場所と制御の設定を指定します。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 574e04ec577f831e17184cf4f801dae22441bca2
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215076"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="a340c-103">DesktopSettings 要素</span><span class="sxs-lookup"><span data-stu-id="a340c-103">DesktopSettings element</span></span>

<span data-ttu-id="a340c-104">メール アドインがデスクトップ コンピューターで使用されるときに適用されるソースの場所と制御の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="a340c-104">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a340c-105">この`DesktopSettings`要素は、web 上の従来の Outlook (社内 Exchange サーバーの古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="a340c-105">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="a340c-106">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="a340c-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a340c-107">構文</span><span class="sxs-lookup"><span data-stu-id="a340c-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="a340c-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="a340c-108">Contained in</span></span>

[<span data-ttu-id="a340c-109">Form</span><span class="sxs-lookup"><span data-stu-id="a340c-109">Form</span></span>](form.md)
