# <a name="customtab-element"></a><span data-ttu-id="abbba-101">CustomTab 要素</span><span class="sxs-lookup"><span data-stu-id="abbba-101">CustomTab element</span></span>

<span data-ttu-id="abbba-p101">リボン上で、アドイン コマンドに使用するタブとグループを指定します。これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。</span><span class="sxs-lookup"><span data-stu-id="abbba-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="abbba-p102">カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、一つのカスタム タブに制限されています。</span><span class="sxs-lookup"><span data-stu-id="abbba-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="abbba-107">**id** 属性はマニフェスト内で一意でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="abbba-107">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="abbba-108">子要素</span><span class="sxs-lookup"><span data-stu-id="abbba-108">Child elements</span></span>

|  <span data-ttu-id="abbba-109">要素</span><span class="sxs-lookup"><span data-stu-id="abbba-109">Element</span></span> |  <span data-ttu-id="abbba-110">必須</span><span class="sxs-lookup"><span data-stu-id="abbba-110">Required</span></span>  |  <span data-ttu-id="abbba-111">説明</span><span class="sxs-lookup"><span data-stu-id="abbba-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="abbba-112">グループ</span><span class="sxs-lookup"><span data-stu-id="abbba-112">Group</span></span>](group.md)      | <span data-ttu-id="abbba-113">はい</span><span class="sxs-lookup"><span data-stu-id="abbba-113">Yes</span></span> |  <span data-ttu-id="abbba-114">コマンドのグループを定義します。</span><span class="sxs-lookup"><span data-stu-id="abbba-114">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="abbba-115">ラベル</span><span class="sxs-lookup"><span data-stu-id="abbba-115">Label</span></span>](#label-tab)      | <span data-ttu-id="abbba-116">はい</span><span class="sxs-lookup"><span data-stu-id="abbba-116">Yes</span></span> |  <span data-ttu-id="abbba-117">CustomTab または Group のラベル。</span><span class="sxs-lookup"><span data-stu-id="abbba-117">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="abbba-118">コントロール</span><span class="sxs-lookup"><span data-stu-id="abbba-118">Control</span></span>](control.md)    | <span data-ttu-id="abbba-119">はい</span><span class="sxs-lookup"><span data-stu-id="abbba-119">Yes</span></span> |  <span data-ttu-id="abbba-120">一つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="abbba-120">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="abbba-121">グループ</span><span class="sxs-lookup"><span data-stu-id="abbba-121">Group</span></span>

<span data-ttu-id="abbba-p103">必須です。[Group 要素](group.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="abbba-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="abbba-124">ラベル(タブ)</span><span class="sxs-lookup"><span data-stu-id="abbba-124">Label (Tab)</span></span>

<span data-ttu-id="abbba-p104">必須。カスタム タブのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="abbba-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="abbba-127">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="abbba-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```