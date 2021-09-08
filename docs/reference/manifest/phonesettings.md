---
title: マニフェスト ファイルの PhoneSettings 要素
description: PhoneSettings 要素は、メール アドインが電話で使用される場合に適用されるソースの場所と制御設定を指定します。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: d7957e23a77a0f837366e5cedc0e0f350b5635c8
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936351"
---
# <a name="phonesettings-element"></a>PhoneSettings 要素

メール アドインが電話で使用されるときに適用されるソースの場所と制御の設定を指定します。

> [!IMPORTANT]
> この要素は、従来の Outlook on the web (通常はオンプレミスの Exchange サーバーの古いバージョンに接続されている) および Outlook `PhoneSettings` 2013 Windows。 Android と iOS Outlookをサポートするには[、「Add-ins for Outlook Mobile 」を参照してください](../../outlook/outlook-mobile-addins.md)。

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

