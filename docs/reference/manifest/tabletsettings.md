---
title: マニフェスト ファイルの TabletSettings 要素
description: TabletSettings 要素は、メール アドインをタブレットで使用するときに適用されるコントロール設定を指定します。
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 3d7ace7fe9258ee32f3f5507d35b35ae026ef5eb
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151213"
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
