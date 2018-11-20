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

|**要素**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Id](id.md)|x|x|x|
|[Version](version.md)|x|x|x|
|[ProviderName](providername.md)|x|x|x|
|[DefaultLocale](defaultlocale.md)|x|x|x|
|[DefaultSettings](defaultsettings.md)|x||x|
|[DisplayName](displayname.md)|x|x|x|
|[Description](description.md)|x|x|x|
|[FormSettings](formsettings.md)||x||
|[Permissions](permissions.md)|x||x|
|[Rule](rule.md)||x||

## <a name="can-contain"></a>含めることができるもの

|**要素**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](alternateid.md)|x|x|x|
|[IconUrl](iconurl.md)|x|x|x|
|[HighResolutionIconUrl](highresolutioniconurl.md)|x|x|x|
|[SupportUrl](supporturl.md)|x|x|x|
|[AppDomains](appdomains.md)|x|x|x|
|[Hosts](hosts.md)|x|x|x|
|[Requirements](requirements.md)|x|x|x|
|[AllowSnapshot](allowsnapshot.md)|x|||
|[Permissions](permissions.md)||x||
|[DisableEntityHighlighting](disableentityhighlighting.md)||x||
|[Dictionary](dictionary.md)|||x|
|[VersionOverrides](versionoverrides.md)||x||

## <a name="attributes"></a>属性

|||
|:-----|:-----|
|xmlns|Office アドイン マニフェストの名前空間とスキーマ バージョンを定義します。この属性は常に `"http://schemas.microsoft.com/office/appforoffice/1.1"` に設定する必要があります。|
|xmlns:xsi|XMLSchema インスタンスを定義します。この属性は常に `"http://www.w3.org/2001/XMLSchema-instance"` に設定する必要があります。|
|xsi:type|Office アドインの種類を定義します。この属性は、`"ContentApp"`、`"MailApp"`、または `"TaskPaneApp"` のいずれかに設定する必要があります。|
