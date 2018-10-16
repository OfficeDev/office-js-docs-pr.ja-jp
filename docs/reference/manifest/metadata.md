# <a name="metadata-element"></a><span data-ttu-id="f8f42-101">MetaData 要素</span><span class="sxs-lookup"><span data-stu-id="f8f42-101">MetaData element</span></span>

<span data-ttu-id="f8f42-102">Excel でユーザー定義関数によって使用されるメタデータの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="f8f42-102">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="f8f42-103">属性</span><span class="sxs-lookup"><span data-stu-id="f8f42-103">Attributes</span></span>

<span data-ttu-id="f8f42-104">なし</span><span class="sxs-lookup"><span data-stu-id="f8f42-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="f8f42-105">子要素</span><span class="sxs-lookup"><span data-stu-id="f8f42-105">Child elements</span></span>

|  <span data-ttu-id="f8f42-106">要素</span><span class="sxs-lookup"><span data-stu-id="f8f42-106">Element</span></span>  |  <span data-ttu-id="f8f42-107">必須</span><span class="sxs-lookup"><span data-stu-id="f8f42-107">Required</span></span>  |  <span data-ttu-id="f8f42-108">説明</span><span class="sxs-lookup"><span data-stu-id="f8f42-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f8f42-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f8f42-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="f8f42-110">はい</span><span class="sxs-lookup"><span data-stu-id="f8f42-110">Yes</span></span>  | <span data-ttu-id="f8f42-111">カスタム関数によって使用される JSON  ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="f8f42-111">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="f8f42-112">例</span><span class="sxs-lookup"><span data-stu-id="f8f42-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
