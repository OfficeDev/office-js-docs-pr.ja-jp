# <a name="appdomains-element"></a><span data-ttu-id="a6c9b-101">AppDomains 要素</span><span class="sxs-lookup"><span data-stu-id="a6c9b-101">AppDomains element</span></span>

<span data-ttu-id="a6c9b-p101">Office アドイン でページを読み込むのに使う SourceLocation 要素で指定されたドメインの他に、任意のドメインを一覧表示します。追加の各ドメインに、AppDomain 要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="a6c9b-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="a6c9b-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="a6c9b-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a6c9b-105">構文</span><span class="sxs-lookup"><span data-stu-id="a6c9b-105">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="a6c9b-106">すべての **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain<AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="a6c9b-106">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="a6c9b-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="a6c9b-107">Contained in</span></span>

[<span data-ttu-id="a6c9b-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="a6c9b-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="a6c9b-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="a6c9b-109">Can contain</span></span>

[<span data-ttu-id="a6c9b-110">AppDomain</span><span class="sxs-lookup"><span data-stu-id="a6c9b-110">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="a6c9b-111">解説</span><span class="sxs-lookup"><span data-stu-id="a6c9b-111">Remarks</span></span>

<span data-ttu-id="a6c9b-112">アドインは、既定では [SourceLocation](sourcelocation.md) 要素で指定されたものと同じ場所のドメインのページを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="a6c9b-112">By default, your add-in can load any page that is in the same domain as the location specified in the SourceLocation element. To load pages that are not in the same domain as the add-in, specify the domains by using the AppDomains and AppDomain elements. This element can't be empty.</span></span> <span data-ttu-id="a6c9b-113">アドインと同じドメインにないページを読み込む場合は、**AppDomains** 要素と **AppDomain** 要素を使用してドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="a6c9b-113">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="a6c9b-114">この要素は空にできません。</span><span class="sxs-lookup"><span data-stu-id="a6c9b-114">This element can't be empty.</span></span>
