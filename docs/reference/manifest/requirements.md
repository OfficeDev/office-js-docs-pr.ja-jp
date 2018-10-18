# <a name="requirements-element"></a>Requirements 要素

Office アドインをアクティブにするために必要な JavaScript API for Office の最小要件セット ([要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>次に含まれる:

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができるもの:

|**要素**|**コンテンツ**|**メール**|**作業ウィンドウ**|
|:-----|:-----|:-----|:-----|
|[セット](sets.md)|x|x|x|
|[メソッド](methods.md)|x||x|

## <a name="remarks"></a>注釈

要求セットの詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。

