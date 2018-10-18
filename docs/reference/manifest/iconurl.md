# <a name="iconurl-element"></a>IconUrl 要素

挿入 UX と Office ストアの Office アドインを表すために使用されるイメージの URL を指定します。

**アドインの種類 : **コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>含めることができるもの:

[上書き](override.md)

## <a name="attributes"></a>属性

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|文字列|必須|この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。|

## <a name="remarks"></a>注釈

メール アドインの場合、アイコンは、**[ファイル]**  >  **[アドインの管理]** UI (Outlook) または **[設定]**  >  **[アドインの管理]** UI (Outlook Web App) に表示されます。コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]**  >  **[アドイン]** UI に表示されます。どのアドインの種類についても、アドインを Office ストアに公開すると、アイコンは Office ストア サイトでも使用されます。

イメージは、GIF、JPG、PNG、EXIF、BMP、または TIFF のいずれかのファイル形式である必要があります。 コンテンツと作業ウィンドウ アプリでは、指定したイメージは 32 x 32 ピクセルである必要があります。 メール アプリでは、イメージは 64 × 64 ピクセルである必要があります。 また、[HighResolutionIconUrl](highresolutioniconurl.md) 要素を使用して、高 DPI 画面で実行されている Office ホスト アプリケーションで使用するためのアイコンも指定しなければなりません。 詳細については、「_AppSource および Office 内で効果的な一覧を作成する_」の 「[アプリに一貫性のあるビジュアル ID を作成する](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)」セクションを参照してください。
