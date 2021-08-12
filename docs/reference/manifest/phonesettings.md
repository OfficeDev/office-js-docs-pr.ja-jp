---
title: マニフェスト ファイルの PhoneSettings 要素
description: PhoneSettings 要素は、メール アドインが電話で使用される場合に適用されるソースの場所と制御設定を指定します。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 883242dc290384f9f0b72736338bdd78a2d23ffeee6cf3aee46d5acd970654ab
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57085730"
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

