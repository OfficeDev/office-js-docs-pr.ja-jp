---
title: マニフェスト ファイルの TabletSettings 要素
description: TabletSettings 要素は、メール アドインをタブレットで使用するときに適用されるコントロール設定を指定します。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: b5a74db4f9fb43df10a08ab43b59507f6e0d7952
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936932"
---
# <a name="tabletsettings-element"></a>TabletSettings 要素

メール アドインがタブレットで使用されるときに適用される制御の設定を指定します。

> [!IMPORTANT]
> この要素は、従来の Outlook on the web (通常はオンプレミスの Exchange サーバーの古いバージョンに接続されている) および Outlook `TabletSettings` 2013 Windows。 Android と iOS Outlookをサポートするには[、「Add-ins for Outlook Mobile 」を参照してください](../../outlook/outlook-mobile-addins.md)。

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
