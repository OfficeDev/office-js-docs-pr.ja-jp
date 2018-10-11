# <a name="officemenu-element"></a><span data-ttu-id="3f744-101">OfficeMenu 要素</span><span class="sxs-lookup"><span data-stu-id="3f744-101">OfficeMenu element</span></span>

<span data-ttu-id="3f744-p101">Office のコンテキスト メニューに追加されるコントロールのコレクションを定義します。Word、Excel、PowerPoint、OneNote アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="3f744-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="3f744-104">属性</span><span class="sxs-lookup"><span data-stu-id="3f744-104">Attributes</span></span>

| <span data-ttu-id="3f744-105">属性</span><span class="sxs-lookup"><span data-stu-id="3f744-105">Attribute</span></span>            | <span data-ttu-id="3f744-106">必須</span><span class="sxs-lookup"><span data-stu-id="3f744-106">Required</span></span> | <span data-ttu-id="3f744-107">説明</span><span class="sxs-lookup"><span data-stu-id="3f744-107">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="3f744-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="3f744-108">xsi:type</span></span>](#xsitype) | <span data-ttu-id="3f744-109">はい</span><span class="sxs-lookup"><span data-stu-id="3f744-109">Yes</span></span>      | <span data-ttu-id="3f744-110">定義されている OfficeMenu の種類。</span><span class="sxs-lookup"><span data-stu-id="3f744-110">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="3f744-111">子要素</span><span class="sxs-lookup"><span data-stu-id="3f744-111">Child elements</span></span>

|  <span data-ttu-id="3f744-112">要素</span><span class="sxs-lookup"><span data-stu-id="3f744-112">Element</span></span> |  <span data-ttu-id="3f744-113">必須</span><span class="sxs-lookup"><span data-stu-id="3f744-113">Required</span></span>  |  <span data-ttu-id="3f744-114">説明</span><span class="sxs-lookup"><span data-stu-id="3f744-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3f744-115">コントロール</span><span class="sxs-lookup"><span data-stu-id="3f744-115">Control</span></span>](#control)    | <span data-ttu-id="3f744-116">はい</span><span class="sxs-lookup"><span data-stu-id="3f744-116">Yes</span></span> |  <span data-ttu-id="3f744-117">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="3f744-117">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="3f744-118">xsi:type</span><span class="sxs-lookup"><span data-stu-id="3f744-118">xsi:type</span></span>

<span data-ttu-id="3f744-119">この Office アドインを追加する Office クライアント アプリケーションの組み込みメニューを指定します。</span><span class="sxs-lookup"><span data-stu-id="3f744-119">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="3f744-p102">`ContextMenuText` - テキストが選択され、その選択されたテキストのコンテキスト メニューをユーザーが開いたときに (右クリック)、コンテキスト メニューに項目が表示されます。Word、Excel、PowerPoint、OneNote に適用されます。</span><span class="sxs-lookup"><span data-stu-id="3f744-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="3f744-p103">`ContextMenuCell` - ユーザーがスプレッドシートのセルのコンテキスト メニューを開くと (右クリック)、コンテキスト メニューに項目が表示されます。Excel に適用されます。</span><span class="sxs-lookup"><span data-stu-id="3f744-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="3f744-124">コントロール</span><span class="sxs-lookup"><span data-stu-id="3f744-124">Control</span></span>

<span data-ttu-id="3f744-125">各 **OfficeMenu** 要素には、1 つ以上の[メニュー](control.md#menu-dropdown-button-controls)コントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="3f744-125">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="3f744-126">例</span><span class="sxs-lookup"><span data-stu-id="3f744-126">Example</span></span>

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
