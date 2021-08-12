---
title: マニフェスト ファイルの CustomTab 要素
description: リボン上で、アドイン コマンドに使用するタブとグループを指定します。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: de6233966abea4de423f255bda3c9e6e38ff5037c760c90cae7c8a1c7ca6ab2e
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57085055"
---
# <a name="customtab-element"></a>CustomTab 要素

リボンで、アドイン コマンドのタブとグループを指定します。 これは既定のタブ ([**ホーム**]、[**メッセージ**]、[**会議**] のいずれか)、またはアドインで定義されたカスタム タブになります。

カスタム タブでは、アドインにカスタム グループまたは組み込みグループを設定できます。 アドインは、カスタム タブ 1 つに制限されています。

**id 属性は** マニフェスト内で一意である必要があります。

> [!IMPORTANT]
> Mac Outlookでは、要素は使用できないので、代わりに `CustomTab` [OfficeTab を使用する](officetab.md)必要があります。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Group](group.md)      | いいえ |  コマンドのグループを定義します。  |
|  [OfficeGroup](#officegroup)      | いいえ |  組み込みのコントロール グループOfficeします。 **重要**: このサイトではOutlook。 |
|  [Label](#label-tab)      | はい |  CustomTab または Group のラベル。  |
|  [InsertAfter](#insertafter)      | いいえ |  指定した組み込みタブの直後にカスタム タブOffice指定します。**重要**: Outlook。 |
|  [InsertBefore](#insertbefore)      | いいえ |  指定した組み込みタブの直前にカスタム タブOffice指定します。**重要**: Outlook。 |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | いいえ |  カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームの組み合わせにカスタム タブを表示するかどうかを指定します。 **重要**: このサイトではOutlook。 |

### <a name="group"></a>グループ

省略可能ですが、存在しない場合は、少なくとも 1 つの **OfficeGroup 要素が必要** です。 [Group 要素](group.md)を参照してください。 マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。複数の要素がある場合は、これらの要素を混同できますが、すべてが Label 要素の上にある **必要** があります。

### <a name="officegroup"></a>OfficeGroup

省略可能ですが、存在しない場合は、少なくとも 1 つの Group 要素が **必要** です。 組み込みのコントロール グループOfficeします。 **id 属性** は、組み込みのグループの ID Officeします。 組み込みグループの ID を見つけるには、「コントロールとコントロール グループの ID を検索 [する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。複数の要素がある場合は、これらの要素を混同できますが、すべてが Label 要素の上にある **必要** があります。

> [!IMPORTANT]
> 要素 `OfficeGroup` は、このプロパティではOutlook。

### <a name="label-tab"></a>Label (タブ)

必須です。 カスタム タブのラベル。**resid 属性** は 32 文字以内で、Resources 要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。

### <a name="insertafter"></a>InsertAfter

省略可能。 指定した組み込みタブの直後にカスタム タブを指定Officeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。 (「 [コントロールとコントロール グループの ID を検索する」を参照](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)してください。存在する場合は、Label 要素の後に **指定する必要** があります。 **InsertAfter** と **InsertBefore の両方を使用することはできません**。

> [!IMPORTANT]
> 要素 `InsertAfter` は、このプロパティではOutlook。

### <a name="insertbefore"></a>InsertBefore

省略可能。 指定した組み込みタブの直前にカスタム タブを指定Officeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。 (「 [コントロールとコントロール グループの ID を検索する」を参照](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)してください。 存在する場合は、Label 要素の後に **指定する必要** があります。 **InsertAfter** と **InsertBefore の両方を使用することはできません**。

> [!IMPORTANT]
> 要素 `InsertBefore` は、このプロパティではOutlook。

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

省略可能 (ブール型)。 カスタム コンテキスト タブを実行時にリボンにインストールする API をサポートするアプリケーションとプラットフォームの組み合わせで **CustomTab** を非表示にするかどうかを指定します。 既定値 (存在しない場合) は、 です `false` 。 使用する場合 **、OverriddenByRibbonApi は** CustomTab の *最初* の子 **である必要があります**。 詳細については [、「OverriddenByRibbonApi」を参照してください](overriddenbyribbonapi.md)。

> [!IMPORTANT]
> 要素 `OverriddenByRibbonApi` は、このプロパティではOutlook。

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
