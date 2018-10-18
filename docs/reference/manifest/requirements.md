# <a name="requirements-element"></a><span data-ttu-id="e509a-101">Requirements 要素</span><span class="sxs-lookup"><span data-stu-id="e509a-101">Requirements element</span></span>

<span data-ttu-id="e509a-102">Office アドインをアクティブにするために必要な JavaScript API for Office の最小要件セット ([要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。</span><span class="sxs-lookup"><span data-stu-id="e509a-102">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="e509a-103">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="e509a-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e509a-104">構文</span><span class="sxs-lookup"><span data-stu-id="e509a-104">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="e509a-105">次に含まれる:</span><span class="sxs-lookup"><span data-stu-id="e509a-105">Contained in:</span></span>

[<span data-ttu-id="e509a-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="e509a-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="e509a-107">含めることができるもの:</span><span class="sxs-lookup"><span data-stu-id="e509a-107">Can contain:</span></span>

|<span data-ttu-id="e509a-108">**要素**</span><span class="sxs-lookup"><span data-stu-id="e509a-108">**Element**</span></span>|<span data-ttu-id="e509a-109">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="e509a-109">**Content**</span></span>|<span data-ttu-id="e509a-110">**メール**</span><span class="sxs-lookup"><span data-stu-id="e509a-110">**Mail**</span></span>|<span data-ttu-id="e509a-111">**作業ウィンドウ**</span><span class="sxs-lookup"><span data-stu-id="e509a-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="e509a-112">セット</span><span class="sxs-lookup"><span data-stu-id="e509a-112">Sets</span></span>](sets.md)|<span data-ttu-id="e509a-113">x</span><span class="sxs-lookup"><span data-stu-id="e509a-113">x</span></span>|<span data-ttu-id="e509a-114">x</span><span class="sxs-lookup"><span data-stu-id="e509a-114">x</span></span>|<span data-ttu-id="e509a-115">x</span><span class="sxs-lookup"><span data-stu-id="e509a-115">x</span></span>|
|[<span data-ttu-id="e509a-116">メソッド</span><span class="sxs-lookup"><span data-stu-id="e509a-116">Methods</span></span>](methods.md)|<span data-ttu-id="e509a-117">x</span><span class="sxs-lookup"><span data-stu-id="e509a-117">x</span></span>||<span data-ttu-id="e509a-118">x</span><span class="sxs-lookup"><span data-stu-id="e509a-118">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="e509a-119">注釈</span><span class="sxs-lookup"><span data-stu-id="e509a-119">Remarks</span></span>

<span data-ttu-id="e509a-120">要求セットの詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e509a-120">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

