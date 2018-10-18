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

## <a name="contained-in"></a>次に含まれる:

 _なし_

## <a name="must-contain"></a>含める必要があるもの:

|**要素**|**コンテンツ**|**メール**|**作業ウィンドウ**|
|:-----|:-----|:-----|:-----|
|[ID](id.md)|x|x|x|
|[バージョン](version.md)|x|x|x|
|[ProviderName](providername.md)|x|x|x|
|[DefaultLocale](defaultlocale.md)|x|x|x|
|[DefaultSettings](defaultsettings.md)|x||x|
|[DisplayName](displayname.md)|x|x|x|
|[説明](description.md)|x|x|x|
|[FormSettings](formsettings.md)||x||
|[アクセス許可](permissions.md)|x||x|
|[ルール](rule.md)||x||

## <a name="can-contain"></a>含めることができるもの:

|**要素**|**コンテンツ**|**メール**|**作業ウィンドウ**|
|:-----|:-----|:-----|:-----|
|[AlternateId](alternateid.md)|x|x|x|
|[IconUrl](iconurl.md)|x|x|x|
|[HighResolutionIconUrl](highresolutioniconurl.md)|x|x|x|
|[SupportUrl](supporturl.md)|x|x|x|
|[AppDomains](appdomains.md)|x|x|x|
|[ホスト](hosts.md)|x|x|x|
|[要件](requirements.md)|x|x|x|
|[AllowSnapshot](allowsnapshot.md)|x|||
|[アクセス許可](permissions.md)||x||
|[DisableEntityHighlighting](disableentityhighlighting.md)||x||
|[辞典](dictionary.md)|||x|
|[VersionOverrides](versionoverrides.md)|X|X|X|

## <a name="attributes"></a>属性

|||
|:-----|:-----|
|xmlns||||UNTRANSLATED_CONTENT_START|||Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to|||UNTRANSLATED_CONTENT_END|||  `"http://schemas.microsoft.com/office/appforoffice/1.1"`|
|xmlns:xsi||||UNTRANSLATED_CONTENT_START|||Defines the XMLSchema instance. This attribute should always be set to|||UNTRANSLATED_CONTENT_END|||  `"http://www.w3.org/2001/XMLSchema-instance"`|
|xsi:type||||UNTRANSLATED_CONTENT_START|||Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or|||UNTRANSLATED_CONTENT_END|||  `"TaskPaneApp"`|
