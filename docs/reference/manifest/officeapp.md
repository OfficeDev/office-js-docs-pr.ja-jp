---
title: マニフェスト ファイルの OfficeApp 要素
description: OfficeApp 要素は、アドイン マニフェストのOffice要素です。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 39ab9285720f7a9a7b5eede1cd5883e2d42602f9be86b7fe756713e0b98e9218
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089045"
---
# <a name="officeapp-element"></a>OfficeApp 要素

Office アドインのマニフェストのルート要素。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a>次に含まれる

 _none_

## <a name="must-contain"></a>含める必要があるもの

|要素|コンテンツ|メール|TaskPane|
|:-----|:-----|:-----|:-----|
|[Id](id.md)|x|x|x|
|[バージョン](version.md)|x|x|x|
|[ProviderName](providername.md)|x|x|x|
|[DefaultLocale](defaultlocale.md)|x|x|x|
|[DefaultSettings](defaultsettings.md)|x||x|
|[DisplayName](displayname.md)|x|x|x|
|[説明](description.md)|x|x|x|
|[FormSettings](formsettings.md)||x||
|[アクセス許可](permissions.md)|x||x|
|[Rule](rule.md)||x||

## <a name="can-contain"></a>含めることができるもの

|要素|コンテンツ|メール|TaskPane|
|:-----|:-----|:-----|:-----|
|[AlternateId](alternateid.md)|x|x|x|
|[IconUrl](iconurl.md)|x|x|x|
|[HighResolutionIconUrl](highresolutioniconurl.md)|x|x|x|
|[SupportUrl](supporturl.md)|x|x|x|
|[AppDomains](appdomains.md)|x|x|x|
|[Hosts](hosts.md)|x|x|x|
|[Requirements](requirements.md)|x|x|x|
|[AllowSnapshot](allowsnapshot.md)|x|||
|[アクセス許可](permissions.md)||x||
|[DisableEntityHighlighting](disableentityhighlighting.md)||x||
|[Dictionary](dictionary.md)|||x|
|[VersionOverrides](versionoverrides.md)|x|x|x|
|[ExtendedOverrides](extendedoverrides.md)|||x|

## <a name="attributes"></a>属性

|属性|説明|
|:-----|:-----|
|xmlns|Office アドイン マニフェストの名前空間とスキーマ バージョンを定義します。この属性は常に `"http://schemas.microsoft.com/office/appforoffice/1.1"` に設定する必要があります。|
|xmlns:xsi|XMLSchema インスタンスを定義します。この属性は常に `"http://www.w3.org/2001/XMLSchema-instance"` に設定する必要があります。|
|xsi:type|Office アドインの種類を定義します。この属性は、`"ContentApp"`、`"MailApp"`、または `"TaskPaneApp"` のいずれかに設定する必要があります。|
