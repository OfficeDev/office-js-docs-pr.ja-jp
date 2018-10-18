# <a name="override-element"></a><span data-ttu-id="bb3cd-101">Override 要素</span><span class="sxs-lookup"><span data-stu-id="bb3cd-101">Override element</span></span>

<span data-ttu-id="bb3cd-102">追加ロケールの設定の値を指定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-102">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="bb3cd-103">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="bb3cd-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bb3cd-104">構文</span><span class="sxs-lookup"><span data-stu-id="bb3cd-104">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="bb3cd-105">次に含まれる:</span><span class="sxs-lookup"><span data-stu-id="bb3cd-105">Contained in:</span></span>

|<span data-ttu-id="bb3cd-106">**要素**</span><span class="sxs-lookup"><span data-stu-id="bb3cd-106">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="bb3cd-107">CitationText</span><span class="sxs-lookup"><span data-stu-id="bb3cd-107">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="bb3cd-108">説明</span><span class="sxs-lookup"><span data-stu-id="bb3cd-108">Description</span></span>](description.md)|
|[<span data-ttu-id="bb3cd-109">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="bb3cd-109">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="bb3cd-110">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="bb3cd-110">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="bb3cd-111">DisplayName</span><span class="sxs-lookup"><span data-stu-id="bb3cd-111">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="bb3cd-112">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="bb3cd-112">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="bb3cd-113">IconUrl</span><span class="sxs-lookup"><span data-stu-id="bb3cd-113">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="bb3cd-114">QueryUri</span><span class="sxs-lookup"><span data-stu-id="bb3cd-114">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="bb3cd-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="bb3cd-115">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="bb3cd-116">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="bb3cd-116">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="bb3cd-117">属性</span><span class="sxs-lookup"><span data-stu-id="bb3cd-117">Attributes</span></span>

|<span data-ttu-id="bb3cd-118">**属性**</span><span class="sxs-lookup"><span data-stu-id="bb3cd-118">**Attribute**</span></span>|<span data-ttu-id="bb3cd-119">**型**</span><span class="sxs-lookup"><span data-stu-id="bb3cd-119">**Type**</span></span>|<span data-ttu-id="bb3cd-120">**必須**</span><span class="sxs-lookup"><span data-stu-id="bb3cd-120">**Required**</span></span>|<span data-ttu-id="bb3cd-121">**説明**</span><span class="sxs-lookup"><span data-stu-id="bb3cd-121">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bb3cd-122">言語</span><span class="sxs-lookup"><span data-stu-id="bb3cd-122">Locale</span></span>|<span data-ttu-id="bb3cd-123">文字列</span><span class="sxs-lookup"><span data-stu-id="bb3cd-123">string</span></span>|<span data-ttu-id="bb3cd-124">必須</span><span class="sxs-lookup"><span data-stu-id="bb3cd-124">required</span></span>|<span data-ttu-id="bb3cd-125">|||UNTRANSLATED_CONTENT_START|||Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="bb3cd-125">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="bb3cd-126">値</span><span class="sxs-lookup"><span data-stu-id="bb3cd-126">Value</span></span>|<span data-ttu-id="bb3cd-127">文字列</span><span class="sxs-lookup"><span data-stu-id="bb3cd-127">string</span></span>|<span data-ttu-id="bb3cd-128">必須</span><span class="sxs-lookup"><span data-stu-id="bb3cd-128">required</span></span>|<span data-ttu-id="bb3cd-129">指定のロケールに対して表される設定の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-129">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="bb3cd-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="bb3cd-130">See also</span></span>

- [<span data-ttu-id="bb3cd-131">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="bb3cd-131">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
