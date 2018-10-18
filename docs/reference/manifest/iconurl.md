# <a name="iconurl-element"></a><span data-ttu-id="5cc2a-101">IconUrl 要素</span><span class="sxs-lookup"><span data-stu-id="5cc2a-101">IconUrl element</span></span>

<span data-ttu-id="5cc2a-102">挿入 UX と Office ストアの Office アドインを表すために使用されるイメージの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="5cc2a-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="5cc2a-103">\*\*アドインの種類 : \*\*コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="5cc2a-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="5cc2a-104">構文</span><span class="sxs-lookup"><span data-stu-id="5cc2a-104">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="5cc2a-105">含めることができるもの:</span><span class="sxs-lookup"><span data-stu-id="5cc2a-105">Can contain:</span></span>

[<span data-ttu-id="5cc2a-106">上書き</span><span class="sxs-lookup"><span data-stu-id="5cc2a-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="5cc2a-107">属性</span><span class="sxs-lookup"><span data-stu-id="5cc2a-107">Attributes</span></span>

|<span data-ttu-id="5cc2a-108">**属性**</span><span class="sxs-lookup"><span data-stu-id="5cc2a-108">**Attribute**</span></span>|<span data-ttu-id="5cc2a-109">**型**</span><span class="sxs-lookup"><span data-stu-id="5cc2a-109">**Type**</span></span>|<span data-ttu-id="5cc2a-110">**必須**</span><span class="sxs-lookup"><span data-stu-id="5cc2a-110">**Required**</span></span>|<span data-ttu-id="5cc2a-111">**説明**</span><span class="sxs-lookup"><span data-stu-id="5cc2a-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5cc2a-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="5cc2a-112">DefaultValue</span></span>|<span data-ttu-id="5cc2a-113">文字列</span><span class="sxs-lookup"><span data-stu-id="5cc2a-113">string</span></span>|<span data-ttu-id="5cc2a-114">必須</span><span class="sxs-lookup"><span data-stu-id="5cc2a-114">required</span></span>|<span data-ttu-id="5cc2a-115">この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。</span><span class="sxs-lookup"><span data-stu-id="5cc2a-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="5cc2a-116">注釈</span><span class="sxs-lookup"><span data-stu-id="5cc2a-116">Remarks</span></span>

<span data-ttu-id="5cc2a-p101">メール アドインの場合、アイコンは、**[ファイル]**  >  **[アドインの管理]** UI (Outlook) または **[設定]**  >  **[アドインの管理]** UI (Outlook Web App) に表示されます。コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]**  >  **[アドイン]** UI に表示されます。どのアドインの種類についても、アドインを Office ストアに公開すると、アイコンは Office ストア サイトでも使用されます。</span><span class="sxs-lookup"><span data-stu-id="5cc2a-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook Web App). For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.</span></span>

<span data-ttu-id="5cc2a-120">イメージは、GIF、JPG、PNG、EXIF、BMP、または TIFF のいずれかのファイル形式である必要があります。</span><span class="sxs-lookup"><span data-stu-id="5cc2a-120">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="5cc2a-121">コンテンツと作業ウィンドウ アプリでは、指定したイメージは 32 x 32 ピクセルである必要があります。</span><span class="sxs-lookup"><span data-stu-id="5cc2a-121">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="5cc2a-122">メール アプリでは、イメージは 64 × 64 ピクセルである必要があります。</span><span class="sxs-lookup"><span data-stu-id="5cc2a-122">For mail apps, the image must be 64 x 64 pixels.</span></span> <span data-ttu-id="5cc2a-123">また、[HighResolutionIconUrl](highresolutioniconurl.md) 要素を使用して、高 DPI 画面で実行されている Office ホスト アプリケーションで使用するためのアイコンも指定しなければなりません。</span><span class="sxs-lookup"><span data-stu-id="5cc2a-123">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="5cc2a-124">詳細については、「_AppSource および Office 内で効果的な一覧を作成する_」の 「[アプリに一貫性のあるビジュアル ID を作成する](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="5cc2a-124">For more information, see the section  Create a consistent visual identity for your app in Create effective Office Store apps and add-ins.</span></span>
