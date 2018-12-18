# <a name="appdomain-element"></a>AppDomain 要素

アドイン ウィンドウにページを読み込むために使用される追加のドメインを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain<AppDomain>`) が含まれている必要があります。

## <a name="contained-in"></a>含まれる場所

[AppDomains](appdomains.md)

## <a name="remarks"></a>解説

**AppDomain** 要素は、[SourceLocation](sourcelocation.md) 要素で指定したドメイン以外のものを追加指定するために使用する必要があります。 詳細については、「[Office アドイン XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)」を参照してください。
