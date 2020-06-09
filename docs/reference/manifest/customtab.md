---
title: マニフェスト ファイルの CustomTab 要素
description: リボン上で、アドイン コマンドに使用するタブとグループを指定します。
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: a81b64a17eeeb463d55024e189b09048b2eb96ac
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612305"
---
# <a name="customtab-element"></a><span data-ttu-id="6749c-103">CustomTab 要素</span><span class="sxs-lookup"><span data-stu-id="6749c-103">CustomTab element</span></span>

<span data-ttu-id="6749c-104">リボン上で、アドイン コマンドに使用するタブとグループを指定します。</span><span class="sxs-lookup"><span data-stu-id="6749c-104">On the ribbon, you specify which tab and group for their add-in commands.</span></span> <span data-ttu-id="6749c-105">これは既定のタブ ([**ホーム**]、[**メッセージ**]、[**会議**] のいずれか)、またはアドインで定義されたカスタム タブになります。</span><span class="sxs-lookup"><span data-stu-id="6749c-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="6749c-p102">カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="6749c-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="6749c-109">**Id**属性はマニフェスト内で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="6749c-109">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6749c-110">Mac 上の Outlook では、要素を使用でき `CustomTab` ないため、代わりに[[officetab タブ](officetab.md)を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6749c-110">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="6749c-111">子要素</span><span class="sxs-lookup"><span data-stu-id="6749c-111">Child elements</span></span>

|  <span data-ttu-id="6749c-112">要素</span><span class="sxs-lookup"><span data-stu-id="6749c-112">Element</span></span> |  <span data-ttu-id="6749c-113">必須</span><span class="sxs-lookup"><span data-stu-id="6749c-113">Required</span></span>  |  <span data-ttu-id="6749c-114">説明</span><span class="sxs-lookup"><span data-stu-id="6749c-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6749c-115">Group</span><span class="sxs-lookup"><span data-stu-id="6749c-115">Group</span></span>](group.md)      | <span data-ttu-id="6749c-116">はい</span><span class="sxs-lookup"><span data-stu-id="6749c-116">Yes</span></span> |  <span data-ttu-id="6749c-117">コマンドのグループを定義します。</span><span class="sxs-lookup"><span data-stu-id="6749c-117">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="6749c-118">Label</span><span class="sxs-lookup"><span data-stu-id="6749c-118">Label</span></span>](#label-tab)      | <span data-ttu-id="6749c-119">はい</span><span class="sxs-lookup"><span data-stu-id="6749c-119">Yes</span></span> |  <span data-ttu-id="6749c-120">CustomTab または Group のラベル。</span><span class="sxs-lookup"><span data-stu-id="6749c-120">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="6749c-121">Group</span><span class="sxs-lookup"><span data-stu-id="6749c-121">Group</span></span>

<span data-ttu-id="6749c-p103">必須です。 [Group 要素](group.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6749c-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="6749c-124">Label (タブ)</span><span class="sxs-lookup"><span data-stu-id="6749c-124">Label (Tab)</span></span>

<span data-ttu-id="6749c-125">必須です。</span><span class="sxs-lookup"><span data-stu-id="6749c-125">Required.</span></span> <span data-ttu-id="6749c-126">ユーザー設定のタブのラベルを示します。**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6749c-126">The label of the custom tab. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="6749c-127">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="6749c-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
