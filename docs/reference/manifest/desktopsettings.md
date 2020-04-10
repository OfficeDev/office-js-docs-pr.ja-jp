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
# <a name="desktopsettings-element"></a>DesktopSettings 要素

メール アドインがデスクトップ コンピューターで使用されるときに適用されるソースの場所と制御の設定を指定します。

> [!IMPORTANT]
> この`DesktopSettings`要素は、web 上の従来の Outlook (社内 Exchange サーバーの古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。

**アドインの種類:** メール

## <a name="syntax"></a>構文

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

## <a name="contained-in"></a>含まれる場所

[Form](form.md)
