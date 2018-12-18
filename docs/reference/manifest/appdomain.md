# <a name="appdomain-element"></a><span data-ttu-id="b44c6-101">AppDomain 要素</span><span class="sxs-lookup"><span data-stu-id="b44c6-101">AppDomain element</span></span>

<span data-ttu-id="b44c6-102">アドイン ウィンドウにページを読み込むために使用される追加のドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="b44c6-102">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="b44c6-103">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="b44c6-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b44c6-104">構文</span><span class="sxs-lookup"><span data-stu-id="b44c6-104">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> <span data-ttu-id="b44c6-105">**AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain<AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="b44c6-105">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="b44c6-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="b44c6-106">Contained in</span></span>

[<span data-ttu-id="b44c6-107">AppDomains</span><span class="sxs-lookup"><span data-stu-id="b44c6-107">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="b44c6-108">解説</span><span class="sxs-lookup"><span data-stu-id="b44c6-108">Remarks</span></span>

<span data-ttu-id="b44c6-109">**AppDomain** 要素は、[SourceLocation](sourcelocation.md) 要素で指定したドメイン以外のものを追加指定するために使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b44c6-109">The  AppDomains and **AppDomain** elements are used to specify any additional domains other than the one specified in the [SourceLocation element. For more information, see Office Add-ins XML manifest](sourcelocation.md).</span></span> <span data-ttu-id="b44c6-110">詳細については、「[Office アドイン XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b44c6-110">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
