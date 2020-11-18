---
title: マニフェスト ファイルの CustomTab 要素
description: リボン上で、アドイン コマンドに使用するタブとグループを指定します。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 99670b27d963060a008899a8808ca967cfd710a6
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087939"
---
# <a name="customtab-element"></a>CustomTab 要素

リボンで、アドインコマンドのタブとグループを指定します。 これは既定のタブ ([**ホーム**]、[**メッセージ**]、[**会議**] のいずれか)、またはアドインで定義されたカスタム タブになります。

カスタムタブでは、アドインにカスタムグループまたは組み込みグループを含めることができます。 アドインは、カスタム タブ 1 つに制限されています。

**Id** 属性はマニフェスト内で一意である必要があります。

> [!IMPORTANT]
> Mac 上の Outlook では、要素を使用でき `CustomTab` ないため、代わりに [[officetab タブ](officetab.md) を使用する必要があります。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Group](group.md)      | いいえ |  コマンドのグループを定義します。  |
|  [OfficeGroup](#officegroup)      | いいえ |  組み込みの Office コントロールグループを表します。  |
|  [Label](#label-tab)      | はい |  CustomTab または Group のラベル。  |
|  [InsertAfter](#insertafter)      | いいえ |  指定した組み込みの Office タブの直後にカスタムタブを作成するように指定します。  |
|  [InsertBefore](#insertbefore)      | いいえ |  指定した組み込みの Office タブの直前にカスタムタブを表示するように指定します。  |

### <a name="group"></a>Group

省略可能。ただし、指定されていない場合は、少なくとも1つの **Officegroup** 要素が存在している必要があります。 [Group 要素](group.md)を参照してください。 マニフェスト内の **グループ** と **officegroup** の順序は、[ユーザー設定] タブに表示する順序にする必要があります。複数の要素がある場合には混在させることができますが、 **Label** 要素の上にある必要があります。

### <a name="officegroup"></a>OfficeGroup

省略可能。ただし、指定されていない場合は、少なくとも1つの **Group** 要素が存在している必要があります。 組み込みの Office コントロールグループを表します。 **Id** 属性は、組み込みの Office グループの id を指定します。 組み込みグループの ID を検索するには、「 [コントロールおよびコントロールグループの id を検索](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)する」を参照してください。 マニフェスト内の **グループ** と **officegroup** の順序は、[ユーザー設定] タブに表示する順序にする必要があります。複数の要素がある場合には混在させることができますが、 **Label** 要素の上にある必要があります。

### <a name="label-tab"></a>Label (タブ)

必須です。 ユーザー設定のタブのラベルを示します。**Resid** 属性は、 [Resources](resources.md)要素の Short **strings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。

### <a name="insertafter"></a>InsertAfter

省略可能。 指定した組み込みの Office タブの直後にカスタムタブを作成するように指定します。要素の値は、組み込みタブの ID ("TabHome"、"Tabhome" など) です。 (「 [コントロールおよびコントロールグループの id を検索する」を](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)参照してください)。指定する場合は、 **Label** 要素の後にする必要があります。 **InsertAfter** と **insertbefore** の両方を使用することはできません。

### <a name="insertbefore"></a>InsertBefore

省略可能。 指定した組み込みの Office タブの直前にカスタムタブを表示するように指定します。要素の値は、組み込みタブの ID ("TabHome"、"Tabhome" など) です。 (「 [コントロールおよびコントロールグループの id を検索する」を](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)参照してください)。 指定する場合は、 **Label** 要素の後にする必要があります。 **InsertAfter** と **insertbefore** の両方を使用することはできません。

## <a name="customtab-example"></a>CustomTab の例

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
