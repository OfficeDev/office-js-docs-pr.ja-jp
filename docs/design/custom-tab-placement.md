---
title: カスタム タブをリボンに配置する
description: Office リボンにカスタム タブが表示される場所と、既定でフォーカスがあるかどうかを制御する方法について説明します。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 42445898623e082c3c85e756625307dc5a237c28
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659816"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a>カスタム タブをリボンに配置する

アドインのマニフェストのマークアップを使用して、アドインのカスタム タブを Office アプリケーションのリボンに表示する場所を指定できます。

> [!NOTE]
> この記事では、 [アドイン コマンドの基本的な概念](add-in-commands.md)に関する記事を理解していることを前提としています。 最近行っていない場合は、確認してください。

> [!IMPORTANT]
>
> - この記事で説明するアドイン機能とマークアップは *、PowerPoint on the webでのみ使用できます*。
> - この記事で説明するマークアップは、要件セット **AddinCommands 1.3** をサポートするプラットフォームでのみ機能します。 以下 [の「サポートされていないプラットフォームでの動作」を](#behavior-on-unsupported-platforms) 参照してください。

カスタム タブを表示する場所を指定するには、その横に表示する組み込み Office タブを特定し、組み込みタブの左側または右側に配置するかどうかを指定します。アドインのマニフェストの [CustomTab](/javascript/api/manifest/customtab) 要素に [InsertBefore](/javascript/api/manifest/customtab#insertbefore) (左) または [InsertAfter](/javascript/api/manifest/customtab#insertafter) (右) 要素を含めることで、これらの仕様を作成します。 (両方の要素を持つことはできません。)

次の例では、カスタム タブが **[校閲**] タブ *のすぐ後* に表示されるように構成されています。要素の **\<InsertAfter\>** 値は、組み込みの Office タブの ID であることに注意してください。 

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

次の点に留意してください。

- 要素と **\<InsertAfter\>** 要素は **\<InsertBefore\>** 省略可能です。 どちらも使用しないと、リボンの右端のタブとしてカスタム タブが表示されます。
- 要素と **\<InsertAfter\>** 要素は **\<InsertBefore\>** 相互に排他的です。 両方を使用することはできません。
- ユーザーが同じ場所に対してカスタム タブが構成されている複数のアドインをインストールした場合 ( **たとえば、[校閲** ] タブの後) は、最後にインストールされたアドインのタブがその場所に配置されます。 以前にインストールしたアドインのタブは、1 か所に移動されます。 たとえば、ユーザーはその順序でアドイン A、B、C をインストールし、すべて **[校閲** ] タブの後にタブを挿入するように構成されます。その後、タブは次の順序で表示されます: **Review**、 **AddinCTab**、 **AddinBTab**、 **AddinATab**。
- ユーザーは、Office アプリケーションでリボンをカスタマイズできます。 たとえば、ユーザーはアドインのタブを移動または非表示にすることができます。これを防止したり、発生したことを検出したりすることはできません。
- ユーザーが組み込みタブの 1 つを移動した場合、Office は *組み込みタブの既定の場所* の観点から要素と **\<InsertAfter\>** 要素を解釈 **\<InsertBefore\>** します。たとえば、ユーザーが **[校閲**] タブをリボンの右端に移動した場合、Office は前の例のマークアップを"***[校閲**] タブの既定の場所* の右側にカスタム タブを配置する" という意味として解釈します。

## <a name="specify-which-tab-has-focus-when-the-document-opens"></a>ドキュメントが開いたときにフォーカスがあるタブを指定する

Office では常に、[ **ファイル** ] タブの右側にあるタブに既定のフォーカスが設定されます。既定では、[ **ホーム** ] タブです。[ **ホーム** ] タブの前にカスタム タブを設定すると、 `<InsertBefore>TabHome</InsertBefore>`ドキュメントが開いたときにカスタム タブにフォーカスが設定されます。

> [!IMPORTANT]
> アドインの不便さを過度に目立たせ、ユーザーや管理者を悩ませます。 アドインがユーザーがドキュメントを操作する主な方法でない限り、カスタム タブを **[ホーム** ] タブの前に配置しないでください。

## <a name="behavior-on-unsupported-platforms"></a>サポートされていないプラットフォームでの動作

アドインが [要件セット AddinCommands 1.3](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets) をサポートしていないプラットフォームにインストールされている場合、この記事で説明されているマークアップは無視され、カスタム タブはリボンの右端のタブとして表示されます。 マークアップをサポートしていないプラットフォームにアドインがインストールされないようにするには、マニフェストのセクションで要件セットへの参照を **\<Requirements\>** 追加します。 手順については、「アドインを [ホストできる Office バージョンとプラットフォームを指定する」を参照してください](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)。 または、「代替エクスペリエンスの設計」で説明されているように、 **AddinCommands 1.3** がサポートされていない場合に、別のエクスペリエンスを持つアドイン [を設計します](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences)。 たとえば、カスタム タブが目的の場所であると想定する手順がアドインに含まれている場合は、タブが最も右側であることを前提とした別のバージョンを使用できます。
