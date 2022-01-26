---
title: 組み込みのコントロール Officeカスタム コントロール グループとタブに統合する
description: カスタム コマンド グループとタブに組み込Officeボタンをリボンに含めるOfficeします。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 81765f470d95a43e597e06f976ad2bfa2a7b66c8
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222130"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>組み込みのコントロール Officeカスタム コントロール グループとタブに統合する

アドインのマニフェストでマークアップOffice使用して、Office リボンのカスタム コントロール グループに組み込みのコントロール ボタンを挿入できます。 (カスタム アドイン コマンドを組み込みのアドイン グループにOfficeできます)。また、組み込みのコントロール グループ全体Officeカスタム リボン タブに挿入することもできます。

> [!NOTE]
> この記事では、アドイン コマンドの基本的な概念に [精通している必要があります](add-in-commands.md)。 最近行っていない場合は、確認してください。

> [!IMPORTANT]
>
> - この記事で説明するアドイン機能とマークアップは、この記事でのみ *PowerPoint on the web。*
> - この記事で説明するマークアップは、要件セット **AddinCommands 1.3** をサポートするプラットフォームでのみ機能します。 後のセクション「 [サポートされていないプラットフォームでの動作」を参照してください](#behavior-on-unsupported-platforms)。

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>組み込みのコントロール グループをカスタム タブに挿入する

組み込みのコントロール グループOfficeタブに挿入するには、[親 CustomTab](../reference/manifest/customtab.md#officegroup)要素に子要素として **OfficeGroup 要素を追加** します。 `id` **OfficeGroup** 要素の属性は、組み込みグループの ID に設定されます。 「 [コントロールとコントロール グループの ID を検索する」を参照してください](#find-the-ids-of-controls-and-control-groups)。

次のマークアップ例は、Office段落コントロール グループをカスタム タブに追加し、カスタム グループの直後に表示する位置を設定します。

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a>組み込みのコントロールをカスタム グループに挿入する

カスタム グループに組み込Officeコントロールを挿入するには、親 Group 要素に子要素として [OfficeControl](../reference/manifest/group.md#officecontrol)要素を **追加** します。 `id` **OfficeControl** 要素の属性は、組み込みコントロールの ID に設定されます。 「 [コントロールとコントロール グループの ID を検索する」を参照してください](#find-the-ids-of-controls-and-control-groups)。

次のマークアップ例は、superscript コントロールOfficeカスタム グループに追加し、カスタム ボタンの直後に表示する位置を設定します。

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
> ユーザーは、アプリケーションでリボンをOfficeできます。 ユーザーのカスタマイズは、マニフェスト設定を上書きします。 たとえば、ユーザーは任意のグループからボタンを削除し、タブから任意のグループを削除できます。

## <a name="find-the-ids-of-controls-and-control-groups"></a>コントロールとコントロール グループの ID を検索する

サポートされているコントロールとコントロール グループの ID は、repo コントロールのファイル内Office[に含まれます](https://github.com/OfficeDev/office-control-ids)。 そのレポの ReadMe ファイルの指示に従います。

## <a name="behavior-on-unsupported-platforms"></a>サポートされていないプラットフォームでの動作

アドインが要件セット[AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)をサポートしていないプラットフォームにインストールされている場合、この記事で説明するマークアップは無視され、組み込みの Office コントロール/グループはカスタム グループ/タブに表示されません。 マークアップをサポートしないプラットフォームにアドインがインストールされるのを防ぐには、マニフェストの [要件] セクションで要件セットへの参照を追加します。 手順については、「アドインをホストOfficeバージョンとプラットフォームを指定[する」を参照してください](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)。 または、「代替エクスペリエンス用のデザイン」で説明するように、アドインを設計して **AddinCommands 1.3** がサポートされていない場合にエクスペリエンス [を提供するようにします](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences)。 たとえば、組み込みのボタンがカスタム グループ内にあると仮定する手順がアドインに含まれている場合は、組み込みボタンが通常の場所にのみ含まれると想定するバージョンを設計できます。
