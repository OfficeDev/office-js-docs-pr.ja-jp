# <a name="allformfactors-element"></a><span data-ttu-id="2da16-101">AllFormFactors 要素</span><span class="sxs-lookup"><span data-stu-id="2da16-101">AllFormFactors element</span></span>

<span data-ttu-id="2da16-102">すべてのフォーム ファクターについてアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="2da16-102">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="2da16-103">現在、\*\* AllFormFactors\*\*  を使用する機能はカスタム関数のみです。</span><span class="sxs-lookup"><span data-stu-id="2da16-103">Currently, the only feature using AllFormFactors is custom functions.</span></span> <span data-ttu-id="2da16-104">\*\* AllFormFactors\*\*  は、カスタム関数を使用するときの必須要素です。</span><span class="sxs-lookup"><span data-stu-id="2da16-104">AllFormFactors is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="2da16-105">子要素</span><span class="sxs-lookup"><span data-stu-id="2da16-105">Child elements</span></span>

|  <span data-ttu-id="2da16-106">要素</span><span class="sxs-lookup"><span data-stu-id="2da16-106">Element</span></span> |  <span data-ttu-id="2da16-107">必須</span><span class="sxs-lookup"><span data-stu-id="2da16-107">Required</span></span>  |  <span data-ttu-id="2da16-108">説明</span><span class="sxs-lookup"><span data-stu-id="2da16-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="2da16-109">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="2da16-109">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="2da16-110">はい</span><span class="sxs-lookup"><span data-stu-id="2da16-110">Yes</span></span> |  <span data-ttu-id="2da16-111">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="2da16-111">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="2da16-112">AllFormFactors の例</span><span class="sxs-lookup"><span data-stu-id="2da16-112">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
