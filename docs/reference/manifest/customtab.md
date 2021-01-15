---
title: マニフェスト ファイルの CustomTab 要素
description: リボン上で、アドイン コマンドに使用するタブとグループを指定します。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 642222af02431814e4e64141504911c67ca829fa
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771327"
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
|  [OfficeGroup](#officegroup)      | いいえ |  組み込みのコントロール グループOffice表します。  |
|  [Label](#label-tab)      | はい |  CustomTab または Group のラベル。  |
|  [InsertAfter](#insertafter)      | いいえ |  指定した組み込みのタブの直後にカスタム タブをOfficeします。  |
|  [InsertBefore](#insertbefore)      | いいえ |  ユーザー設定タブを指定した組み込みタブの直前にOfficeします。  |

### <a name="group"></a>グループ

省略可能ですが、存在しない場合は、少なくとも 1 つの **OfficeGroup 要素が必要** です。 [Group 要素](group.md)を参照してください。 マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。要素が複数ある場合は、これらの要素が不確定になる可能性がありますが、すべて Label 要素の上に配置する **必要** があります。

### <a name="officegroup"></a>OfficeGroup

省略可能ですが、存在しない場合は、少なくとも 1 つの Group 要素が **必要** です。 組み込みのコントロール グループOffice表します。 **id 属性** は、グループに組み込Officeします。 組み込みのグループの ID を検索するには、「コントロールとコントロール グループの ID を検索する」 [を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。要素が複数ある場合は、これらの要素が不確定になる可能性がありますが、すべて Label 要素の上に配置する **必要** があります。

### <a name="label-tab"></a>Label (タブ)

Required. カスタム タブのラベルを指定します。**resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。

### <a name="insertafter"></a>InsertAfter

省略可能です。 指定した組み込みのタブの直後にカスタム タブをOfficeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。 (「 [コントロールとコントロール グループの ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups))。存在する場合は、Label 要素の後 **に配置する必要** があります。 InsertAfter と **InsertBefore の両方を指定することはできません**。

### <a name="insertbefore"></a>InsertBefore

省略可能です。 ユーザー設定タブを指定した組み込みタブの直前にOfficeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。 (「 [コントロールとコントロール グループの ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups))。 存在する場合は、Label 要素の後 **に配置する必要** があります。 InsertAfter と **InsertBefore の両方を指定することはできません**。

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
