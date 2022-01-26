---
title: カスタム タブをリボンに配置する
description: カスタム タブがリボンに表示される場所と、既定Officeフォーカスが設定されているかどうかを制御する方法について説明します。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: bced5bf5506d0366b29d8e2ad6803b0ddfaad80b
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222095"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a>カスタム タブをリボンに配置する

アドインのマニフェストでマークアップを使用して、Office アプリケーションのリボンにアドインのカスタム タブを表示する場所を指定できます。

> [!NOTE]
> この記事では、アドイン コマンドの基本的な概念に [精通している必要があります](add-in-commands.md)。 最近行ったことがない場合は、確認してください。

> [!IMPORTANT]
>
> - この記事で説明するアドイン機能とマークアップは、この記事でのみ *PowerPoint on the web。*
> - この記事で説明するマークアップは、要件セット **AddinCommands 1.3** をサポートするプラットフォームでのみ機能します。 以下の [「サポートされていないプラットフォームでの動作」を参照](#behavior-on-unsupported-platforms) してください。

カスタム タブを表示する場所を指定するには、どの組み込みの Office タブを隣に配置するか、組み込みタブの左側または右側に配置する必要があるかどうかを指定します。アドインのマニフェストの[CustomTab](../reference/manifest/customtab.md)要素に[InsertBefore](../reference/manifest/customtab.md#insertbefore) (左) 要素または[InsertAfter](../reference/manifest/customtab.md#insertafter) (右) 要素を含めて、これらの仕様を指定します。 (両方の要素を持つ必要があります)。

次の例では、カスタム タブが [レビュー] タブの直後 *に表示* するように **構成** されています。**InsertAfter** 要素の値は、組み込みの [プロパティ] タブOffice注意してください。 

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group2">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

以下の点を念頭に置いておきます。

- **InsertBefore 要素と** **InsertAfter** 要素はオプションです。 どちらも使用しない場合は、カスタム タブがリボンの右端のタブとして表示されます。
- **InsertBefore 要素と** **InsertAfter** 要素は相互に排他的です。 両方を使用することはできません。
- ユーザーが複数のアドインをインストールし、そのユーザー設定タブが同じ場所に構成されている場合は、[確認]タブの後に、最近インストールしたアドインのタブがその場所に配置されます。 以前にインストールしたアドインのタブは、1 か所に移動されます。 たとえば、ユーザーはその順序でアドイン A、B、C をインストールし、すべて [レビュー] タブの後にタブを挿入するように構成され、タブは次の順序で表示されます。[レビュー] **、AddinCTab** **、AddinBTab、AddinATab** の順にタブが **表示** されます。
- ユーザーは、アプリケーションでリボンをOfficeできます。 たとえば、ユーザーはアドインのタブを移動または非表示にできます。これを防止したり、発生したことを検出したりすることはできません。
- ユーザーが組み込みタブの 1 つを移動すると、Office は、組み込みタブの既定の場所に関して **InsertBefore** 要素と **InsertAfter** 要素を *解釈します*。たとえば、ユーザーが [レビュー]タブをリボンの右側に移動した場合、Office は前の例のマークアップを「既定で [レビュー] タブが表示される場所の右側にカスタム タブを置く」という意味と解釈します。 **

## <a name="specify-which-tab-has-focus-when-the-document-opens"></a>ドキュメントを開く際にフォーカスがあるタブを指定する

Office、[ファイル] タブの右側にあるタブに既定のフォーカスが常に **表示** されます。既定では、[ホーム]**タブ** です。[ホーム] タブの前にカスタムタブを構成すると、ドキュメントが開くと、カスタム タブ `<InsertBefore>TabHome</InsertBefore>` にフォーカスが設定されます。

> [!IMPORTANT]
> アドインの不便さを過度に目立たせ、ユーザーや管理者を悩ませます。 ユーザーがドキュメントを操作する主な方法がアドインではない限り、ユーザー設定タブを [ホーム] タブの前に配置しない。

## <a name="behavior-on-unsupported-platforms"></a>サポートされていないプラットフォームでの動作

アドインが要件セット [AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)をサポートしないプラットフォームにインストールされている場合、この記事で説明するマークアップは無視され、カスタム タブはリボンの右端のタブとして表示されます。 マークアップをサポートしないプラットフォームにアドインがインストールされるのを防ぐには、マニフェストの [要件] セクションで要件セットへの参照を追加します。 手順については、「アドインをホストOfficeバージョンとプラットフォームを指定[する」を参照してください](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)。 または、「代替エクスペリエンスの設計」で説明するように、 **アドインを設計して、AddinCommands 1.3** がサポートされていない場合に別のエクスペリエンス [を提供するようにします](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences)。 たとえば、カスタム タブが必要な場所を想定した手順がアドインに含まれている場合は、タブが右端にあると仮定する別のバージョンを使用できます。
