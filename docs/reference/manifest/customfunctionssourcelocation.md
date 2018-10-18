# <a name="sourcelocation-element"></a><span data-ttu-id="25c64-101">SourceLocation 要素</span><span class="sxs-lookup"><span data-stu-id="25c64-101">SourceLocation element</span></span>

<span data-ttu-id="25c64-102">Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="25c64-102">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="25c64-103">属性</span><span class="sxs-lookup"><span data-stu-id="25c64-103">Attributes</span></span>

| <span data-ttu-id="25c64-104">**属性**</span><span class="sxs-lookup"><span data-stu-id="25c64-104">**Attribute**</span></span> | <span data-ttu-id="25c64-105">**必須**</span><span class="sxs-lookup"><span data-stu-id="25c64-105">**Required**</span></span> | <span data-ttu-id="25c64-106">**説明**</span><span class="sxs-lookup"><span data-stu-id="25c64-106">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="25c64-107">resid</span><span class="sxs-lookup"><span data-stu-id="25c64-107">resid</span></span>         | <span data-ttu-id="25c64-108">はい</span><span class="sxs-lookup"><span data-stu-id="25c64-108">Yes</span></span>          | <span data-ttu-id="25c64-109">マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。</span><span class="sxs-lookup"><span data-stu-id="25c64-109">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="25c64-110">子要素</span><span class="sxs-lookup"><span data-stu-id="25c64-110">Child elements</span></span>

<span data-ttu-id="25c64-111">なし</span><span class="sxs-lookup"><span data-stu-id="25c64-111">None</span></span>

## <a name="example"></a><span data-ttu-id="25c64-112">例</span><span class="sxs-lookup"><span data-stu-id="25c64-112">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```