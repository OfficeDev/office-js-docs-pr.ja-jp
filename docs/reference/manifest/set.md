# <a name="set-element"></a><span data-ttu-id="5a64f-101">セット要素</span><span class="sxs-lookup"><span data-stu-id="5a64f-101">Set element</span></span>

<span data-ttu-id="5a64f-102">Office アドインをアクティブ化するために必要な JavaScript API for Office の要件セットを指定します。</span><span class="sxs-lookup"><span data-stu-id="5a64f-102">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="5a64f-103">\*\*アドインの種類 : \*\*コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="5a64f-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="5a64f-104">構文</span><span class="sxs-lookup"><span data-stu-id="5a64f-104">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="5a64f-105">次に含まれる :</span><span class="sxs-lookup"><span data-stu-id="5a64f-105">Contained in:</span></span>

[<span data-ttu-id="5a64f-106">セット</span><span class="sxs-lookup"><span data-stu-id="5a64f-106">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="5a64f-107">属性</span><span class="sxs-lookup"><span data-stu-id="5a64f-107">Attributes</span></span>

|<span data-ttu-id="5a64f-108">**属性**</span><span class="sxs-lookup"><span data-stu-id="5a64f-108">**Attribute**</span></span>|<span data-ttu-id="5a64f-109">**型**</span><span class="sxs-lookup"><span data-stu-id="5a64f-109">**Type**</span></span>|<span data-ttu-id="5a64f-110">**必須**</span><span class="sxs-lookup"><span data-stu-id="5a64f-110">**Required**</span></span>|<span data-ttu-id="5a64f-111">**説明**</span><span class="sxs-lookup"><span data-stu-id="5a64f-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5a64f-112">名前</span><span class="sxs-lookup"><span data-stu-id="5a64f-112">Name</span></span>|<span data-ttu-id="5a64f-113">文字列</span><span class="sxs-lookup"><span data-stu-id="5a64f-113">string</span></span>|<span data-ttu-id="5a64f-114">必須</span><span class="sxs-lookup"><span data-stu-id="5a64f-114">required</span></span>|<span data-ttu-id="5a64f-115">[要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)の名前。</span><span class="sxs-lookup"><span data-stu-id="5a64f-115">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="5a64f-116">MinVersion</span><span class="sxs-lookup"><span data-stu-id="5a64f-116">MinVersion</span></span>|<span data-ttu-id="5a64f-117">文字列</span><span class="sxs-lookup"><span data-stu-id="5a64f-117">string</span></span>|<span data-ttu-id="5a64f-118">省略可能</span><span class="sxs-lookup"><span data-stu-id="5a64f-118">optional</span></span>|<span data-ttu-id="5a64f-p101">アドインに必要な API セットの最小バージョンを指定します。\*\* DefaultMinVersion\*\* の値が親の [Set](sets.md) 要素に指定されている場合は、その値を上書きします。</span><span class="sxs-lookup"><span data-stu-id="5a64f-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="5a64f-121">注釈</span><span class="sxs-lookup"><span data-stu-id="5a64f-121">Remarks</span></span>

<span data-ttu-id="5a64f-122">要件要求セットの詳細情報については、「[ Office のバージョンおよび要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5a64f-122">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="5a64f-123">\*\*Set **要素の **MinVersion** 属性と**Set \*\*要素の **DefaultMinVersion** 属性の詳細については、「[マニフェストで Requirements 要素を設定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="5a64f-123">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="5a64f-124">メール アドインの場合、`"Mailbox"`連絡可能な要件を 1 つのみ設定します。</span><span class="sxs-lookup"><span data-stu-id="5a64f-124">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="5a64f-125">この要件セットには、 Outlook 向けのメール アドインでサポートされている API の全体のサブセットが含まれ、`"Mailbox"`メール アドインのマニフェストの要件設定を指定しなければなりません ( コンテンツ アドインと作業ウィンドウ アドインの場合とは異なり、オプションではありません ) 。</span><span class="sxs-lookup"><span data-stu-id="5a64f-125">Important  For mail add-ins, there is only one   requirement set available. This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins). Also, you can't declare support for specific methods in mail add-ins.</span></span> <span data-ttu-id="5a64f-126">また、メールのアドインの特定のメソッドのサポートを宣言することはできません。</span><span class="sxs-lookup"><span data-stu-id="5a64f-126">Also, you can't declare support for specific methods in mail add-ins.</span></span>
