# <a name="appdomains-element"></a><span data-ttu-id="1cefd-101">AppDomains 要素</span><span class="sxs-lookup"><span data-stu-id="1cefd-101">AppDomains element</span></span>

<span data-ttu-id="1cefd-p101">Office アドイン でページを読み込むのに使う SourceLocation 要素で指定されたドメインの他に、任意のドメインを一覧表示します。追加の各ドメインに、AppDomain 要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="1cefd-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="1cefd-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="1cefd-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1cefd-105">構文</span><span class="sxs-lookup"><span data-stu-id="1cefd-105">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a><span data-ttu-id="1cefd-106">次に含まれる:</span><span class="sxs-lookup"><span data-stu-id="1cefd-106">Contained in:</span></span>

[<span data-ttu-id="1cefd-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="1cefd-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="1cefd-108">含めることができるもの:</span><span class="sxs-lookup"><span data-stu-id="1cefd-108">Can contain:</span></span>

[<span data-ttu-id="1cefd-109">AppDomain</span><span class="sxs-lookup"><span data-stu-id="1cefd-109">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="1cefd-110">注釈</span><span class="sxs-lookup"><span data-stu-id="1cefd-110">Remarks</span></span>

<span data-ttu-id="1cefd-p102">アドインは、既定では **SourceLocation** 要素で指定されたのと同じ場所のドメインのページを読み込みます。アドインと同じドメインにないページを読み込む場合は、**AppDomains** 要素と **AppDomain** 要素を使ってドメインを指定します。この要素は空にすることはできません。</span><span class="sxs-lookup"><span data-stu-id="1cefd-p102">By default, your add-in can load any page that is in the same domain as the location specified in the **SourceLocation** element. To load pages that are not in the same domain as the add-in, specify the domains by using the **AppDomains** and **AppDomain** elements. This element can't be empty.</span></span> 
