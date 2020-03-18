---
title: マニフェスト ファイルの DesktopSettings 要素
description: メール アドインがデスクトップ コンピューターで使用されるときに適用されるソースの場所と制御の設定を指定します。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d48532482fc71fec2a96133ee8e813cae798613f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718357"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="67fbb-103">DesktopSettings 要素</span><span class="sxs-lookup"><span data-stu-id="67fbb-103">DesktopSettings element</span></span>

<span data-ttu-id="67fbb-104">メール アドインがデスクトップ コンピューターで使用されるときに適用されるソースの場所と制御の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="67fbb-104">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="67fbb-105">この`DesktopSettings`要素は、web 上の従来の Outlook (社内 Exchange サーバーの古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="67fbb-105">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="67fbb-106">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="67fbb-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="67fbb-107">構文</span><span class="sxs-lookup"><span data-stu-id="67fbb-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="67fbb-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="67fbb-108">Contained in</span></span>

[<span data-ttu-id="67fbb-109">Form</span><span class="sxs-lookup"><span data-stu-id="67fbb-109">Form</span></span>](form.md)
