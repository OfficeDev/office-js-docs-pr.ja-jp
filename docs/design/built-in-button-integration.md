---
title: 組み込みの Office ボタンをカスタム コントロール グループとタブに統合する
description: Office リボンのカスタム コマンド グループとタブに組み込みの Office ボタンを含める方法について説明します。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4dc706fcd0b049647847a73f7c40144dba9df0e2
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659788"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>組み込みの Office ボタンをカスタム コントロール グループとタブに統合する

アドインのマニフェストのマークアップを使用して、Office リボンのカスタム コントロール グループに組み込みの Office ボタンを挿入できます。 (組み込みの Office グループにカスタム アドイン コマンドを挿入することはできません。また、組み込みの Office コントロール グループ全体をカスタム リボン タブに挿入することもできます。

> [!NOTE]
> この記事では、 [アドイン コマンドの基本的な概念](add-in-commands.md)に関する記事を理解していることを前提としています。 最近行っていない場合は、確認してください。

> [!IMPORTANT]
>
> - この記事で説明するアドイン機能とマークアップは *、PowerPoint on the webでのみ使用できます*。
> - この記事で説明するマークアップは、要件セット **AddinCommands 1.3** をサポートするプラットフォームでのみ機能します。 後のセクション「 [サポートされていないプラットフォームでの動作」](#behavior-on-unsupported-platforms)を参照してください。

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>組み込みのコントロール グループをカスタム タブに挿入する

組み込みの Office コントロール グループをタブに挿入するには、親 **\<CustomTab\>** 要素の子要素として [OfficeGroup](/javascript/api/manifest/customtab#officegroup) 要素を追加します。 要素の **\<OfficeGroup\>** 属性は`id`、組み込みグループの ID に設定されます。 [「コントロールとコントロール グループの ID を検索する](#find-the-ids-of-controls-and-control-groups)」を参照してください。

次のマークアップ例では、Office Paragraph コントロール グループをカスタム タブに追加し、カスタム グループの直後に表示するように配置します。

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a>組み込みコントロールをカスタム グループに挿入する

組み込みの Office コントロールをカスタム グループに挿入するには、親 **\<Group\>** 要素の子要素として [OfficeControl](/javascript/api/manifest/group#officecontrol) 要素を追加します。 要素の **\<OfficeControl\>** 属性は`id`、組み込みコントロールの ID に設定されます。 [「コントロールとコントロール グループの ID を検索する](#find-the-ids-of-controls-and-control-groups)」を参照してください。

次のマークアップ例では、Office 上付きコントロールをカスタム グループに追加し、カスタム ボタンの直後に表示するように配置します。

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button1">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

> [!NOTE]
> ユーザーは、Office アプリケーションでリボンをカスタマイズできます。 すべてのユーザーのカスタマイズは、マニフェスト設定をオーバーライドします。 たとえば、ユーザーは任意のグループからボタンを削除し、タブから任意のグループを削除できます。

## <a name="find-the-ids-of-controls-and-control-groups"></a>コントロールとコントロール グループの ID を検索する

サポートされているコントロールとコントロール グループの ID は、 [リポジトリの Office コントロール ID](https://github.com/OfficeDev/office-control-ids) 内のファイルにあります。 そのリポジトリの ReadMe ファイルの指示に従います。

## <a name="behavior-on-unsupported-platforms"></a>サポートされていないプラットフォームでの動作

アドインが [要件セット AddinCommands 1.3](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets) をサポートしていないプラットフォームにインストールされている場合、この記事で説明されているマークアップは無視され、組み込みの Office コントロール/グループはカスタム グループ/タブに表示されません。 マークアップをサポートしていないプラットフォームにアドインがインストールされないようにするには、マニフェストのセクションで要件セットへの参照を **\<Requirements\>** 追加します。 手順については、「アドインを [ホストできる Office バージョンとプラットフォームを指定する」を参照してください](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)。 または、「代替エクスペリエンスの設計」で説明されているように、 **AddinCommands 1.3** がサポートされていない場合にエクスペリエンスを持つアドイン [を設計します](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences)。 たとえば、アドインに、組み込みボタンがカスタム グループにあると想定する手順が含まれている場合は、組み込みボタンが通常の場所にあると想定するバージョンを設計できます。
