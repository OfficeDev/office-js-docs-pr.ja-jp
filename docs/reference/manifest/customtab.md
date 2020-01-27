---
title: マニフェスト ファイルの CustomTab 要素
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: c48e526534a3c1295e9c3f0c6fc626df94a874d3
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554014"
---
# <a name="customtab-element"></a><span data-ttu-id="7aee2-102">CustomTab 要素</span><span class="sxs-lookup"><span data-stu-id="7aee2-102">CustomTab element</span></span>

<span data-ttu-id="7aee2-p101">リボン上で、アドイン コマンドに使用するタブとグループを指定します。これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。</span><span class="sxs-lookup"><span data-stu-id="7aee2-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="7aee2-p102">カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="7aee2-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="7aee2-108">**id** 属性はマニフェスト内で一意でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="7aee2-108">The  **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7aee2-109">Mac 上の Outlook では`CustomTab` 、要素を使用できないため、代わりに[[officetab タブ](officetab.md)を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7aee2-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="7aee2-110">子要素</span><span class="sxs-lookup"><span data-stu-id="7aee2-110">Child elements</span></span>

|  <span data-ttu-id="7aee2-111">要素</span><span class="sxs-lookup"><span data-stu-id="7aee2-111">Element</span></span> |  <span data-ttu-id="7aee2-112">必須</span><span class="sxs-lookup"><span data-stu-id="7aee2-112">Required</span></span>  |  <span data-ttu-id="7aee2-113">説明</span><span class="sxs-lookup"><span data-stu-id="7aee2-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7aee2-114">Group</span><span class="sxs-lookup"><span data-stu-id="7aee2-114">Group</span></span>](group.md)      | <span data-ttu-id="7aee2-115">はい</span><span class="sxs-lookup"><span data-stu-id="7aee2-115">Yes</span></span> |  <span data-ttu-id="7aee2-116">コマンドのグループを定義します。</span><span class="sxs-lookup"><span data-stu-id="7aee2-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="7aee2-117">Label</span><span class="sxs-lookup"><span data-stu-id="7aee2-117">Label</span></span>](#label-tab)      | <span data-ttu-id="7aee2-118">はい</span><span class="sxs-lookup"><span data-stu-id="7aee2-118">Yes</span></span> |  <span data-ttu-id="7aee2-119">CustomTab または Group のラベル。</span><span class="sxs-lookup"><span data-stu-id="7aee2-119">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="7aee2-120">Group</span><span class="sxs-lookup"><span data-stu-id="7aee2-120">Group</span></span>

<span data-ttu-id="7aee2-p103">必須です。 [Group 要素](group.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7aee2-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="7aee2-123">Label (タブ)</span><span class="sxs-lookup"><span data-stu-id="7aee2-123">Label (Tab)</span></span>

<span data-ttu-id="7aee2-p104">必須。カスタム タブのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7aee2-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="7aee2-126">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="7aee2-126">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
