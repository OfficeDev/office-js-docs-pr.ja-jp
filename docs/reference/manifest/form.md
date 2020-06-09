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
# <a name="form-element"></a>Form 要素

メール アドインが特定のデバイス (デスクトップ、タブレット、または電話) で実行されているときに使用するフォームの UX の設定。

> [!IMPORTANT]
> 、 `DesktopSettings` 、 `TabletSettings` および要素は、 `PhoneSettings` 従来の Web 上の outlook (通常は社内の Exchange server の古いバージョンに接続されている) と Windows の outlook 2013 でのみ使用できます。

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
