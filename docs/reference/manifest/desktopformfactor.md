# <a name="desktopformfactor-element"></a><span data-ttu-id="40843-101">DesktopFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="40843-101">DesktopFormFactor element</span></span>

<span data-ttu-id="40843-p101">デスクトップ フォーム ファクターについてアドインの設定を指定します。デスクトップ フォーム ファクターには、Office for Windows、Office for Mac、Office Online が含まれています。**Resources** ノードを除くデスクトップ フォーム ファクターのアドイン情報をすべて含みます。</span><span class="sxs-lookup"><span data-stu-id="40843-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="40843-p102">各 DesktopFormFactor の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="40843-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="40843-107">子要素</span><span class="sxs-lookup"><span data-stu-id="40843-107">Child elements</span></span>

| <span data-ttu-id="40843-108">要素</span><span class="sxs-lookup"><span data-stu-id="40843-108">Element</span></span>                               | <span data-ttu-id="40843-109">必須</span><span class="sxs-lookup"><span data-stu-id="40843-109">Required</span></span> | <span data-ttu-id="40843-110">説明</span><span class="sxs-lookup"><span data-stu-id="40843-110">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="40843-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="40843-111">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="40843-112">はい</span><span class="sxs-lookup"><span data-stu-id="40843-112">Yes</span></span>      | <span data-ttu-id="40843-113">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="40843-113">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="40843-114">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="40843-114">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="40843-115">はい</span><span class="sxs-lookup"><span data-stu-id="40843-115">Yes</span></span>      | <span data-ttu-id="40843-116">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="40843-116">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="40843-117">GetStarted</span><span class="sxs-lookup"><span data-stu-id="40843-117">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="40843-118">いいえ</span><span class="sxs-lookup"><span data-stu-id="40843-118">No</span></span>       | <span data-ttu-id="40843-119">Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。</span><span class="sxs-lookup"><span data-stu-id="40843-119">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="40843-120">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="40843-120">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="40843-121">いいえ</span><span class="sxs-lookup"><span data-stu-id="40843-121">No</span></span> | <span data-ttu-id="40843-122">Outlook アドインが委任のシナリオでは使用可能か、既定で *false* に設定されているかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="40843-122">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="40843-123">**重要**: この要素は、Exchange Online への Outlook アドイン プレビュー要求セットでのみ使用可能です。</span><span class="sxs-lookup"><span data-stu-id="40843-123">**Important**: This element is only available in the Outlook add-ins Preview requirement set against Exchange Online.</span></span> <span data-ttu-id="40843-124">この要素を使用するアドインは、AppSource に発行または一元展開で展開できません。</span><span class="sxs-lookup"><span data-stu-id="40843-124">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="40843-125">DesktopFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="40843-125">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
