# <a name="appdomains-element"></a>AppDomains 要素

Office アドイン でページを読み込むのに使う SourceLocation 要素で指定されたドメインの他に、任意のドメインを一覧表示します。追加の各ドメインに、AppDomain 要素を指定します。

 **アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a>次に含まれる:

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができるもの:

[AppDomain](appdomain.md)

## <a name="remarks"></a>注釈

アドインは、既定では **SourceLocation** 要素で指定されたのと同じ場所のドメインのページを読み込みます。アドインと同じドメインにないページを読み込む場合は、**AppDomains** 要素と **AppDomain** 要素を使ってドメインを指定します。この要素は空にすることはできません。 
