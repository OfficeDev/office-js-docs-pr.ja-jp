---
title: マニフェスト ファイルの CustomTab 要素
description: リボン上で、アドイン コマンドに使用するタブとグループを指定します。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6a9540fd7e98464681a90021a36f7a7529186f7f
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340114"
---
# <a name="customtab-element"></a>CustomTab 要素

リボンのカスタム タブOfficeします。 アドインのリボン コントロールとグループを、いずれかのビルド イン Office タブまたは独自のカスタム タブに追加します。**CustomTab** 要素を使用して、カスタム タブをリボンに追加します。 カスタム タブでは、アドインにカスタム グループまたは組み込みグループを設定できます。 アドインは、カスタム タブ 1 つに制限されています。

> [!IMPORTANT]
> Mac Outlookでは **、CustomTab** 要素は使用できませんが、組み込みの [OfficeTabs](officetab.md) の  1 つにカスタム コントロール グループを置くことができます。 組み込 *みのグループは*、任意のプラットフォームの組み込みOutlookに置く必要があります。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

> [!NOTE]
> 一部の子要素は、メール スキーマで無効です。 「 [Child 要素」を参照してください](#child-elements)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)
- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md). 一部の子要素で必要です。 「 [Child 要素」を参照してください](#child-elements)。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  はい  | カスタム タブの一意の ID。|

### <a name="id-attribute"></a>id 属性

必須です。 カスタム タブの一意の識別子。これは、最大 125 文字の文字列です。 これはマニフェスト内で一意である必要があります。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Group](group.md)      | いいえ |  コマンドのグループを定義します。  |
|  [OfficeGroup](#officegroup)      | いいえ |  組み込みのコントロール グループOfficeします。 **重要**: このサイトではOutlook。 |
|  [Label](#label-tab)      | はい |  CustomTab のラベル。  |
|  [InsertAfter](#insertafter)      | いいえ |  指定した組み込みタブの直後にカスタム タブをOfficeします。**重要**: カスタム タブでのみPowerPoint。 |
|  [InsertBefore](#insertbefore)      | いいえ |  カスタム タブを指定した組み込みタブの直前Office指定します。**重要**: カスタム タブでのみPowerPoint。 |

### <a name="group"></a>グループ

省略可能ですが、存在しない場合は、少なくとも 1 つの **OfficeGroup 要素が必要** です。 [Group 要素](group.md)を参照してください。 マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。複数の要素がある場合は、これらの要素を混同できますが、すべてが Label 要素の上にある **必要** があります。

### <a name="officegroup"></a>OfficeGroup

省略可能ですが、存在しない場合は、少なくとも 1 つの Group 要素が **必要** です。 組み込みのコントロール グループOfficeします。 **id 属性** は、組み込みのグループの ID をOfficeします。 組み込みグループの ID を見つけるには、「コントロールとコントロール グループの ID を検索 [する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。複数の要素がある場合は、これらの要素を混同できますが、すべてが Label 要素の上にある **必要** があります。

> [!IMPORTANT]
> **OfficeGroup 要素** は、このプロパティではOutlook。 このPowerPoint、Mac および Windows のプレビュー中ですが、PowerPoint on the web の実稼働アドインで使用できます。

**アドインの種類:** 作業ウィンドウ

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="label-tab"></a>Label (タブ)

必須です。 カスタム タブのラベル。**resid 属性** は 32 文字以内で、Resources 要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="insertafter"></a>InsertAfter

省略可能。 指定した組み込みタブの直後にカスタム タブをOfficeします。要素の値は、 などの組み込みタブの `TabHome` ID です`TabReview`。  組み込みのタブの一覧については、「 [OfficeTab」を参照してください](officetab.md)。 存在する場合は、Label 要素の後に **指定する必要** があります。 **InsertAfter** と **InsertBefore の両方を使用することはできません**。

> [!IMPORTANT]
> **InsertAfter** 要素は、次のページでのみPowerPoint。

**アドインの種類:** 作業ウィンドウ

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="insertbefore"></a>InsertBefore

省略可能。 指定した組み込みタブの直前にカスタム タブを指定Officeします。要素の値は、 などの組み込みタブの `TabHome` ID です`TabReview`。 要素の値は、 などの組み込みタブの `TabHome` ID です `TabReview`。  組み込みのタブの一覧については、「 [OfficeTab」を参照してください](officetab.md)。 存在する場合は、Label 要素の後に **指定する必要** があります。 **InsertAfter** と **InsertBefore の両方を使用することはできません**。

> [!IMPORTANT]
> **InsertBefore** 要素は、次のPowerPoint。

**アドインの種類:** 作業ウィンドウ

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)


## <a name="examples"></a>例

次のマークアップ例は、Office段落コントロール グループをカスタム タブに追加し、カスタム グループの直後に表示する位置を設定します。

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.TabCustom1.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

次のマークアップの例では、Office Superscript コントロールをカスタム グループに追加し、カスタム ボタンの直後に表示する位置を設定します。

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group2">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```
