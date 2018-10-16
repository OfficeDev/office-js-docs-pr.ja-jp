# <a name="set-element"></a>セット要素

Office アドインをアクティブ化するために必要な JavaScript API for Office の要件セットを指定します。

**アドインの種類 : **コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>次に含まれる :

[セット](sets.md)

## <a name="attributes"></a>属性

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|名前|文字列|必須|[要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)の名前。|
|MinVersion|文字列|省略可能|アドインに必要な API セットの最小バージョンを指定します。** DefaultMinVersion** の値が親の [Set](sets.md) 要素に指定されている場合は、その値を上書きします。|

## <a name="remarks"></a>注釈

要件要求セットの詳細情報については、「[ Office のバージョンおよび要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets) 」を参照してください。

**Set **要素の **MinVersion** 属性と**Set **要素の **DefaultMinVersion** 属性の詳細については、「[マニフェストで Requirements 要素を設定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」をご覧ください。

> [!IMPORTANT] 
> メール アドインの場合、`"Mailbox"`連絡可能な要件を 1 つのみ設定します。 この要件セットには、 Outlook 向けのメール アドインでサポートされている API の全体のサブセットが含まれ、`"Mailbox"`メール アドインのマニフェストの要件設定を指定しなければなりません ( コンテンツ アドインと作業ウィンドウ アドインの場合とは異なり、オプションではありません ) 。 また、メールのアドインの特定のメソッドのサポートを宣言することはできません。
