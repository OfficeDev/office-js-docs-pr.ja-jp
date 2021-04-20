---
title: マニフェスト ファイルの CustomTab 要素
description: リボン上で、アドイン コマンドに使用するタブとグループを指定します。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: d74859d1326d29517b5a8226a86f901322957933
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173928"
---
# <a name="customtab-element"></a>CustomTab 要素

リボンで、アドイン コマンドのタブとグループを指定します。 これは既定のタブ ([**ホーム**]、[**メッセージ**]、[**会議**] のいずれか)、またはアドインで定義されたカスタム タブになります。

カスタム タブでは、アドインにカスタム グループまたは組み込みグループを含めできます。 アドインは、カスタム タブ 1 つに制限されています。

**id 属性** はマニフェスト内で一意である必要があります。

> [!IMPORTANT]
> Outlook on Mac では、この要素 `CustomTab` は使用できないので、代わりに [OfficeTab を使用する](officetab.md) 必要があります。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Group](group.md)      | いいえ |  コマンドのグループを定義します。  |
|  [OfficeGroup](#officegroup)      | いいえ |  組み込みのコントロール グループOffice表します。 **重要**: Outlook では使用できません。 |
|  [Label](#label-tab)      | はい |  CustomTab または Group のラベル。  |
|  [InsertAfter](#insertafter)      | いいえ |  ユーザー設定のタブを、指定した組み込みのタブの直後に表示Office指定します。 **重要**: Outlook では使用できません。 |
|  [InsertBefore](#insertbefore)      | いいえ |  ユーザー設定のタブを、指定した組み込みのタブの直前にOffice指定します。 **重要**: Outlook では使用できません。 |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | いいえ |  カスタム タブを、カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームの組み合わせに表示するかどうかを指定します。 **重要**: Outlook では使用できません。 |

### <a name="group"></a>グループ

省略可能ですが、存在しない場合は、少なくとも 1 つの **OfficeGroup 要素が必要** です。 [Group 要素](group.md)を参照してください。 マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。要素が複数ある場合は、これらの要素が不確定になる可能性がありますが、すべてが Label 要素の上にある **必要** があります。

### <a name="officegroup"></a>OfficeGroup

省略可能ですが、存在しない場合は、少なくとも 1 つの Group 要素が **必要** です。 組み込みのコントロール グループOffice表します。 **id 属性** は、グループに組み込Officeします。 組み込みのグループの ID を検索するには、「コントロールとコントロール グループの ID を検索する」 [を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。要素が複数ある場合は、これらの要素が不確定になる可能性がありますが、すべてが Label 要素の上にある **必要** があります。

> [!IMPORTANT]
> この `OfficeGroup` 要素は Outlook では使用できません。

### <a name="label-tab"></a>Label (タブ)

必須です。 カスタム タブのラベルを指定します。**resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。

### <a name="insertafter"></a>InsertAfter

省略可能。 指定した組み込みのタブの直後にカスタム タブをOfficeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。 (「 [コントロールとコントロール グループの ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups))。存在する場合は、Label 要素の後 **に配置する必要** があります。 InsertAfter と **InsertBefore の両方を指定することはできません**。

> [!IMPORTANT]
> この `InsertAfter` 要素は Outlook では使用できません。

### <a name="insertbefore"></a>InsertBefore

省略可能。 ユーザー設定タブを指定した組み込みタブの直前にOfficeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。 (「 [コントロールとコントロール グループの ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups))。 存在する場合は、Label 要素の後 **に配置する必要** があります。 InsertAfter と **InsertBefore の両方を指定することはできません**。

> [!IMPORTANT]
> この `InsertBefore` 要素は Outlook では使用できません。

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

省略可能 (ブール値)。 カスタム コンテキスト タブをリボンに実行時にインストールする API をサポートするアプリケーションとプラットフォームの組み合わせで **CustomTab** を非表示にするかどうかを指定します。 既定値 (存在しない場合) は次の値です `false` 。 使用する場合 **、OverriddenByRibbonApi は** CustomTab の最初 *の子***である必要があります**。 詳細については [、「OverriddenByRibbonApi」を参照してください](overriddenbyribbonapi.md)。

> [!IMPORTANT]
> この `OverriddenByRibbonApi` 要素は Outlook では使用できません。

## <a name="customtab-example"></a>CustomTab の例

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
