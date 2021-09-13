---
title: マニフェスト ファイルの Form 要素
description: メール アドインが特定のデバイス (デスクトップ、タブレット、または電話) で実行されているときに使用するフォームの UX の設定。
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: ecea978cd87a9feed08ba621c9377525acaf4e3a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150235"
---
# <a name="form-element"></a>Form 要素

メール アドインが特定のデバイス (デスクトップ、タブレット、または電話) で実行されているときに使用するフォームの UX の設定。

> [!IMPORTANT]
> 、および要素は、従来の Outlook on the web (通常はオンプレミスの Exchange サーバーの古いバージョンに接続されている) および Windows の `DesktopSettings` `TabletSettings` Outlook `PhoneSettings` でしか使用できません。

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

[FormSettings](formsettings.md)


## <a name="can-contain"></a>含めることができるもの

|**Element**|
|:-----|
|[DesktopSettings](desktopsettings.md)|
|[TabletSettings](tabletsettings.md)|
|[PhoneSettings](phonesettings.md)|
