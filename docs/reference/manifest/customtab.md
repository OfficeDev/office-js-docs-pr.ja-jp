---
title: マニフェスト ファイルの CustomTab 要素
description: リボン上で、アドイン コマンドに使用するタブとグループを指定します。
ms.date: 09/02/2021
localization_priority: Normal
ms.openlocfilehash: 642b6eabaa9885041dd122b179ee2baa3e772977
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937839"
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
|  [InsertAfter](#insertafter)      | いいえ |  カスタム タブを指定した組み込みタブの直後Office指定します。**重要**: PowerPoint でのみ使用できます。 |
|  [InsertBefore](#insertbefore)      | いいえ |  カスタム タブを指定した組み込みタブの直前Office指定します。**重要**: PowerPoint でのみ使用できます。 |

### <a name="group"></a>グループ

省略可能ですが、存在しない場合は、少なくとも 1 つの **OfficeGroup 要素が必要** です。 [Group 要素](group.md)を参照してください。 マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。複数の要素がある場合は、これらの要素を混同できますが、すべてが Label 要素の上にある **必要** があります。

### <a name="officegroup"></a>OfficeGroup

省略可能ですが、存在しない場合は、少なくとも 1 つの Group 要素が **必要** です。 組み込みのコントロール グループOfficeします。 **id 属性** は、組み込みのグループの ID Officeします。 組み込みグループの ID を見つけるには、「コントロールとコントロール グループの ID を検索 [する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。複数の要素がある場合は、これらの要素を混同できますが、すべてが Label 要素の上にある **必要** があります。

> [!IMPORTANT]
> 要素 `OfficeGroup` は、このプロパティではOutlook。

### <a name="label-tab"></a>Label (タブ)

必須です。 カスタム タブのラベル。**resid 属性** は 32 文字以内で、Resources 要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。

### <a name="insertafter"></a>InsertAfter

省略可能です。 指定した組み込みタブの直後にカスタム タブを指定Officeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。 (「 [コントロールとコントロール グループの ID を検索する」を参照](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)してください。存在する場合は、Label 要素の後に **指定する必要** があります。 **InsertAfter** と **InsertBefore の両方を使用することはできません**。

> [!IMPORTANT]
> 要素 `InsertAfter` は、次のPowerPoint。

### <a name="insertbefore"></a>InsertBefore

省略可能。 指定した組み込みタブの直前にカスタム タブを指定Officeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。 (「 [コントロールとコントロール グループの ID を検索する」を参照](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)してください。 存在する場合は、Label 要素の後に **指定する必要** があります。 **InsertAfter** と **InsertBefore の両方を使用することはできません**。

> [!IMPORTANT]
> 要素 `InsertBefore` は、次のPowerPoint。
