---
title: ユーザー設定のタブをリボンに配置する
description: ユーザー設定のタブを Office リボンに表示する方法、および既定でフォーカスがあるかどうかを制御する方法について説明します。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 2c1e2ae66805212e78868cf7c07a0e5c14cd4025
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/17/2020
ms.locfileid: "49088181"
---
# <a name="position-a-custom-tab-on-the-ribbon-preview"></a>リボンにカスタムタブを配置する (プレビュー)

アドインのマニフェストでマークアップを使用することによって、アドインのカスタムタブを Office アプリケーションのリボンに表示する場所を指定できます。

> [!NOTE]
> この記事では、 [アドインコマンドの基本的な概念](add-in-commands.md)について理解していることを前提としています。 最近実行していない場合は、確認してください。

> [!IMPORTANT]
>
> - この記事に記載されているアドイン機能とマークアップはプレビュー段階であり、 *PowerPoint on the web でのみ利用でき* ます。 テストおよび開発環境でマークアップを試すことをお勧めします。 運用環境または業務上重要なドキュメント内では、プレビューマークアップを使用しないでください。
> - この記事に記載されているマークアップは、要件セット **Addincommands 1.3** をサポートするプラットフォームでのみ機能します。 以下のサポートされて [いないプラットフォームでの動作を](#behavior-on-unsupported-platforms) 参照してください。

ユーザー設定のタブを表示する場所を指定するには、そのタブを配置する組み込みの Office タブを特定し、組み込みタブの左側と右側のどちらに配置するかを指定します。これらの指定を行うには、アドインのマニフェストの[Customtab](../reference/manifest/customtab.md)要素に[insertbefore](../reference/manifest/customtab.md#insertbefore) (左側) 要素または[InsertAfter](../reference/manifest/customtab.md#insertafter) (right) 要素のいずれかを含めます。 (両方の要素を使用することはできません。)

次の例では、ユーザー設定のタブが [**校閲**] タブの *すぐ後* に表示されるように構成されています。要素の値は、組み込みの Office タブの ID であることに注意 `<InsertAfter>` してください。 

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

次の点に注意してください。

- `<InsertBefore>`および `<InsertAfter>` 要素はオプションです。 どちらも使用しない場合は、カスタムタブがリボンの右端のタブとして表示されます。
- `<InsertBefore>`および `<InsertAfter>` 要素は相互に排他的です。 両方を使用することはできません。
- ユーザーがカスタムタブを同じ場所に構成した複数のアドインをインストールすると、[ **校閲** ] タブの後に、最後にインストールされたアドインのタブが配置されます。 以前にインストールしたアドインのタブは、1つの場所に移動されます。 たとえば、ユーザーがアドイン A、B、および C をその順序でインストールし、[ **校閲** ] タブの後にタブを挿入するようにすべてが構成されている場合、タブはこの順序で表示されます。 **review**、 **addインシデントタブ**、 **addinbtab**、 **addinbtab**。
- ユーザーは、Office アプリケーションでリボンをカスタマイズできます。 たとえば、ユーザーはアドインのタブを移動したり、非表示にしたりできます。この問題を回避したり、発生したことを検出したりすることはできません。
- ユーザーが組み込みのタブのいずれかを移動すると、Office は `<InsertBefore>` `<InsertAfter>` *組み込みタブの既定の場所* に関して要素と要素を解釈します。たとえば、ユーザーが [**校閲**] タブをリボンの右端に移動すると、Office は上記の例のマークアップを解釈します。 "ユーザー設定のタブは、[ ***校閲**] タブの既定値* の右側に配置します。" という意味があります。

## <a name="specifying-which-tab-has-focus-when-the-document-opens"></a>ドキュメントを開くときにフォーカスがあるタブを指定する

Office は、常に、[ **ファイル** ] タブのすぐ右のタブに既定のフォーカスを与えます。既定では、これは [ **ホーム** ] タブです。ユーザー設定のタブを [ **ホーム** ] タブの前に構成した場合、では `<InsertBefore>TabHome</InsertBefore>` 、ドキュメントを開いたときにユーザー設定のタブにフォーカスが置かれます。

> [!IMPORTANT]
> アドイン inconveniences と annoys ユーザーおよび管理者に過度の prominence を付与します。 ユーザーがドキュメントを操作するための主要な方法でない限り、[ **ホーム** ] タブの前にカスタムタブを配置しないようにします。

## <a name="behavior-on-unsupported-platforms"></a>サポートされていないプラットフォームでの動作

[要件セット AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)をサポートしていないプラットフォームにアドインがインストールされている場合、この記事で説明しているマークアップは無視され、カスタムタブがリボンの右端のタブとして表示されます。 マークアップをサポートしていないプラットフォームにアドインがインストールされないようにするには、マニフェストのセクションにある要件セットへの参照を追加し `<Requirements>` ます。 手順については、「 [マニフェストの要件要素を設定する](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)」を参照してください。 または、「 [JavaScript コードでランタイムチェックを使用](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)する」で説明されているように、 **addincommands 1.3** がサポートされていない場合に、アドインの代替操作を実行するようにアドインを設計することもできます。 たとえば、カスタムタブが目的の場所に配置されていると仮定した場合、アドインには、タブが右端であると想定される代替バージョンがあります。
