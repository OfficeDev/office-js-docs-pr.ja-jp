# <a name="supporturl-element"></a>SupportUrl 要素

アドインのサポート情報を提供するページの URL を指定します。

## <a name="syntax"></a>構文

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a>次に含まれる:

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができるもの:

|  要素 | 必須 | 説明  |
|:-----|:-----|:-----|
|  [オーバーライド](override.md)   | いいえ | 追加のロケール URL の設定を指定します。 |

## <a name="attributes"></a>属性

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必須|この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。|
