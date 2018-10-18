# <a name="requestedheight-element"></a>RequestedHeight 要素

コンテンツのアドインまたはメール アドインの初期の高さ (ピクセル単位で) を指定します。 

**アドインの種類:** コンテンツ

## <a name="syntax"></a>構文

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>次に含まれる:

- [DefaultSettings](defaultsettings.md) (アドインのコンテンツ) の値は、32 から 1000 の間にすることができます
- [DesktopSettings](desktopsettings.md)および[TabletSettings](tabletsettings.md) (メール アドインの場合) の値は、32 ~ 450 を指定できます。
- [ExtensionPoint](extensionpoint.md) (コンテキスト メール アドインの場合) の値は 140 と **DetectedEntity** の拡張点の 450 の間で、32 と **CustomPane** の拡張点の 450 の間で、