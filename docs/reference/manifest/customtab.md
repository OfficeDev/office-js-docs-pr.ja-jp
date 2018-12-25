---
title: マニフェスト ファイルの CustomTab 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 7d609ad216ba5e8e7358bbc741f7b6c992bc97e2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433608"
---
# <a name="customtab-element"></a><span data-ttu-id="30009-102">CustomTab 要素</span><span class="sxs-lookup"><span data-stu-id="30009-102">CustomTab element</span></span>

<span data-ttu-id="30009-p101">リボン上で、アドイン コマンドに使用するタブとグループを指定します。これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。</span><span class="sxs-lookup"><span data-stu-id="30009-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="30009-p102">カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="30009-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="30009-108">**id** 属性はマニフェスト内で一意でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="30009-108">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="30009-109">子要素</span><span class="sxs-lookup"><span data-stu-id="30009-109">Child elements</span></span>

|  <span data-ttu-id="30009-110">要素</span><span class="sxs-lookup"><span data-stu-id="30009-110">Element</span></span> |  <span data-ttu-id="30009-111">必須</span><span class="sxs-lookup"><span data-stu-id="30009-111">Required</span></span>  |  <span data-ttu-id="30009-112">説明</span><span class="sxs-lookup"><span data-stu-id="30009-112">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="30009-113">Group</span><span class="sxs-lookup"><span data-stu-id="30009-113">Group</span></span>](group.md)      | <span data-ttu-id="30009-114">はい</span><span class="sxs-lookup"><span data-stu-id="30009-114">Yes</span></span> |  <span data-ttu-id="30009-115">コマンドのグループを定義します。</span><span class="sxs-lookup"><span data-stu-id="30009-115">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="30009-116">Label</span><span class="sxs-lookup"><span data-stu-id="30009-116">Label</span></span>](#label-tab)      | <span data-ttu-id="30009-117">はい</span><span class="sxs-lookup"><span data-stu-id="30009-117">Yes</span></span> |  <span data-ttu-id="30009-118">CustomTab または Group のラベル。</span><span class="sxs-lookup"><span data-stu-id="30009-118">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="30009-119">Control</span><span class="sxs-lookup"><span data-stu-id="30009-119">Control</span></span>](control.md)    | <span data-ttu-id="30009-120">はい</span><span class="sxs-lookup"><span data-stu-id="30009-120">Yes</span></span> |  <span data-ttu-id="30009-121">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="30009-121">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="30009-122">Group</span><span class="sxs-lookup"><span data-stu-id="30009-122">Group</span></span>

<span data-ttu-id="30009-p103">必須です。[Group 要素](group.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="30009-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="30009-125">Label (タブ)</span><span class="sxs-lookup"><span data-stu-id="30009-125">Label (Tab)</span></span>

<span data-ttu-id="30009-p104">必須。カスタム タブのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="30009-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="30009-128">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="30009-128">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```