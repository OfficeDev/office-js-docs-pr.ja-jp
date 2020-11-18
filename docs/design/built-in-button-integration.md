---
title: 組み込みの Office ボタンをカスタムコントロールグループとタブに統合する
description: Office リボンのカスタムコマンドグループとタブに組み込みの Office ボタンを含める方法について説明します。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: e04107893b3c0dd453c84d38fdd5623e308b70e3
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/17/2020
ms.locfileid: "49088176"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs-preview"></a>組み込みの Office ボタンをカスタムコントロールグループとタブに統合する (プレビュー)

アドインのマニフェストでマークアップを使用すると、office リボンのカスタムコントロールグループに組み込みの Office ボタンを挿入できます。 (組み込みの Office グループにカスタムアドインコマンドを挿入することはできません。)組み込みの Office コントロールグループのすべてをカスタムのリボンタブに挿入することもできます。

> [!NOTE]
> この記事では、 [アドインコマンドの基本的な概念](add-in-commands.md)について理解していることを前提としています。 まだ行っていない場合は、確認してください。

> [!IMPORTANT]
>
> - この記事に記載されているアドイン機能とマークアップはプレビュー段階であり、 *PowerPoint on the web でのみ利用でき* ます。 テストおよび開発環境でマークアップを試すことをお勧めします。 運用環境または業務上重要なドキュメント内では、プレビューマークアップを使用しないでください。
> - この記事に記載されているマークアップは、要件セット **Addincommands 1.3** をサポートするプラットフォームでのみ機能します。 サポートされてい [ないプラットフォームで](#behavior-on-unsupported-platforms)は、後のセクションの動作を参照してください。

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>組み込みのコントロールグループをカスタムタブに挿入する

組み込みの Office コントロールグループをタブに挿入するには、 [Officegroup](../reference/manifest/customtab.md#officegroup) 要素を親要素の子要素として追加し `<CustomTab>` ます。 `id`要素の属性は、 `<OfficeGroup>` 組み込みのグループの ID に設定されます。 「 [コントロールおよびコントロールグループの id を検索する」を](#find-the-ids-of-controls-and-control-groups)参照してください。

次のマークアップの例では、ユーザー設定のタブに Office 段落コントロールグループを追加し、ユーザー設定のグループの直後に表示されるように配置します。

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a>組み込みのコントロールをカスタムグループに挿入する

組み込みの Office コントロールをカスタムグループに挿入するには、親要素の子要素として、 [Officeecontrol](../reference/manifest/group.md#officecontrol) 要素を追加し `<Group>` ます。 `id`要素の属性 `<OfficeControl>` は、組み込みのコントロールの ID に設定されます。 「 [コントロールおよびコントロールグループの id を検索する」を](#find-the-ids-of-controls-and-control-groups)参照してください。

次のマークアップの例では、ユーザー設定のグループに Office の上付きコントロールを追加し、ユーザー設定のボタンの直後に表示されるように配置します。

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.grp1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Button2">
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
> ユーザーは、Office アプリケーションでリボンをカスタマイズできます。 ユーザーのカスタマイズは、マニフェストの設定よりも優先されます。 たとえば、ユーザーは任意のグループからボタンを削除したり、タブから任意のグループを削除したりできます。

## <a name="find-the-ids-of-controls-and-control-groups"></a>コントロールおよびコントロールグループの Id を検索する

サポートされているコントロールおよびコントロールグループの Id は、リポジトリの [Office コントロール id](https://github.com/OfficeDev/office-control-ids)のファイルにあります。 そのリポジトリの ReadMe ファイルの手順に従います。

## <a name="behavior-on-unsupported-platforms"></a>サポートされていないプラットフォームでの動作

[要件セット AddinCommands コマンド 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)をサポートしていないプラットフォームにアドインがインストールされている場合、この記事に記載されているマークアップは無視され、組み込みの Office コントロール/グループはカスタムグループ/タブに表示されません。 マークアップをサポートしていないプラットフォームにアドインがインストールされないようにするには、マニフェストのセクションにある要件セットへの参照を追加し `<Requirements>` ます。 手順については、「 [マニフェストの要件要素を設定する](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)」を参照してください。 または、「 [JavaScript コードでランタイムチェックを使用](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)する」で説明されているように、 **addincommands 1.3** がサポートされていない場合に、アドインの代替操作を実行するようにアドインを設計することもできます。 たとえば、アドインに組み込みのボタンがカスタムグループにあると想定される命令が含まれている場合は、その組み込みボタンが通常の場所にしかないことを前提とした代替バージョンを使用できます。
