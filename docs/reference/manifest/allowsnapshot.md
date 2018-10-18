# <a name="allowsnapshot-element"></a>AllowSnapshot 要素

ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。

**アドインの種類:** コンテンツ

## <a name="syntax"></a>構文

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>次に含まれる:

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注釈

 > [!IMPORTANT]
 > **AllowSnapshot** の規定値は `true` です。 この場合、Office アドインをサポートしていないバージョンのホスト アプリケーションでドキュメントを開くユーザーがアドインのイメージを表示できるようになったり、ホスト アプリケーションがアドインをホストするサーバーに接続できない場合にアドインの静的イメージが提供されたりします。 しかしこれは、アドインがホストされるドキュメントから、アドインに表示される機密性の高い情報に直接アクセスできるということでもあります。

