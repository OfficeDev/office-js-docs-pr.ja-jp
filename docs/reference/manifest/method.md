# <a name="method-element"></a>Method 要素

Office アドインをアクティブにするために必要な JavaScript API for Office の個別のメソッドを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>次に含まれる:

[メソッド](methods.md)

## <a name="attributes"></a>属性

|**属性**|**種類**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|名前|文字列|必須|必要なメソッドの名前をその親オブジェクトで修飾して指定します。たとえば、**getSelectedDataAsync** メソッドを指定するには、`"Document.getSelectedDataAsync"` と指定する必要があります。|

## <a name="remarks"></a>注釈

メールのアドインでは、 **メソッド** および **メソッド** の要素はサポートされていません。要求セットの詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。

> [!IMPORTANT] 
> 個々 のメソッドのバージョンの最小要件を指定する方法がないため、実行時にメソッドが必ず利用可能であるようにするには、アドインのスクリプトにおけるそのメソッドを呼び出すときに **if**  ステートメントも使用する必要があります。 これを行う方法の詳細については、 [Office 用の  JavaScript API for Office を理解する](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)を参照してください。

