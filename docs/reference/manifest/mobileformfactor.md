# <a name="mobileformfactor-element"></a><span data-ttu-id="45a29-101">MobileFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="45a29-101">MobileFormFactor element</span></span>

<span data-ttu-id="45a29-p101">モバイル フォーム ファクターについてアドインの設定を指定します。**Resources** ノードを除くモバイル フォーム ファクターのアドイン情報をすべて含みます。</span><span class="sxs-lookup"><span data-stu-id="45a29-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="45a29-p102">各 **MobileFormFactor** の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="45a29-p102">Each **MobileFormFactor** definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="45a29-p103">**MobileFormFactor** 要素は、VersionOverrides のスキーマ 1.1 で定義されています。これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="45a29-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="45a29-108">子要素</span><span class="sxs-lookup"><span data-stu-id="45a29-108">Child elements</span></span>

| <span data-ttu-id="45a29-109">要素</span><span class="sxs-lookup"><span data-stu-id="45a29-109">Element</span></span>                               | <span data-ttu-id="45a29-110">必須</span><span class="sxs-lookup"><span data-stu-id="45a29-110">Required</span></span> | <span data-ttu-id="45a29-111">説明</span><span class="sxs-lookup"><span data-stu-id="45a29-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="45a29-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="45a29-112">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="45a29-113">はい</span><span class="sxs-lookup"><span data-stu-id="45a29-113">Yes</span></span>      | <span data-ttu-id="45a29-114">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="45a29-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="45a29-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="45a29-115">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="45a29-116">はい</span><span class="sxs-lookup"><span data-stu-id="45a29-116">Yes</span></span>      | <span data-ttu-id="45a29-117">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="45a29-117">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="45a29-118">MobileFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="45a29-118">MobileFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
