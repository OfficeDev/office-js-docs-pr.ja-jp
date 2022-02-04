---
title: メール アドインのマニフェスト ファイル内の VersionOverrides 1.0 要素
description: アドイン マニフェスト (XML) ファイルOffice VersionOverrides 要素 (mail) のリファレンス ドキュメント。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: b433a52ad922fb3d397993a3861038f2f82ff165
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042195"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-mail-add-in"></a>メール アドインのマニフェスト ファイル内の VersionOverrides 1.0 要素

この要素には、基本マニフェストでサポートされていない機能の情報が含まれています。

> [!NOTE]
> この記事では、要素の属性とバリエーションに関する重要な情報を含む [VersionOverrides](versionoverrides.md)要素の概要を理解している必要があります。

**アドインの種類:** メール

**次の VersionOverrides スキーマでのみ有効です**。

- メール 1.0

詳細については、「マニフェストの [バージョンオーバーライド」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [Mailbox 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)
- 一部の子要素は、追加の要件セットに関連付けられる場合があります。

## <a name="child-elements"></a>子要素

次の表は **、VersionOverrides** 要素のバージョン 1.0 にのみ適用され、メール アドインにのみ適用されます。

> [!NOTE]
> iOS では、サポート `<WebApplicationInfo>` されているのは唯一です。 **VersionOverrides** の他のすべての子要素は無視されます。

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [説明](#description)    |  いいえ   |  アドインについての説明。 |
|  [Requirements](requirements.md)  |  いいえ   |  親のマークアップを有効にするためにサポートする必要がある最小要件セット `VersionOverrides` を指定します。 これは、マニフェスト *の基本* 部分の要素よりも常に `Requirements` 制限が厳しい必要があります。|
|  [Hosts](hosts.md)                |  はい  |  アプリケーションのコレクションをOfficeします。 子 Hosts 要素は、マニフェストの親部分にある Hosts 要素をオーバーライドします。  |
|  [Resources](resources.md)    |  はい  | マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。|
|  **VersionOverrides**    |  いいえ  | より新しいスキーマ バージョンでアドイン コマンドを定義します。詳細については、「[複数のバージョンを実装する](#implementing-multiple-versions)」を参照してください。 |
|  [WebApplicationInfo](webapplicationinfo.md)    |  いいえ  | セキュリティで保護されたトークン発行者とのアドインの登録に関する詳細 (V2.0 などAzure Active Directory指定します。 |

### <a name="description"></a>説明

アドインの説明です。 これは、マニフェスト内の任意の親部分の `Description` 要素を上書きします。 説明のテキストは、**Resources** 要素の [LongString](resources.md) 要素の子要素に含まれています。 Description 要素の属性は 32 文字以内で、Resources 要素に含まれる ShortString 要素の子要素の属性の値と一致 `resid`  `id` [する必要](resources.md)があります。 

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「マニフェストの [バージョンオーバーライド」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) 親が `<VersionOverrides>` Taskpane 1.0 型の場合。
- [親が Mail 1.0](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) と入力されている場合のメールボックス `<VersionOverrides>` 1.3。
- [親が Mail 1.1](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) と入力されている場合のメールボックス `<VersionOverrides>` 1.5。

## <a name="example"></a>例

次に簡単な例を示します。 詳細な例については、アドイン コード サンプルのサンプル アドインOffice[を参照してください](https://github.com/OfficeDev/PnP-OfficeAddins)。

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a>複数のバージョンを実装する

1 つのマニフェストで、複数のバージョンの `VersionOverrides` 要素を実装することで、異なるバージョンの VersionOverrides スキーマをサポートできます。これは、新しいスキーマの新機能をオプションでサポートしながら、新機能をサポートしていない古いクライアントもサポートすることで実現できます。

複数のバージョンを実装するために、新しいバージョンの `VersionOverrides` 要素は、古いバージョンの `VersionOverrides` 要素の子にする必要があります。 子の `VersionOverrides` 要素は、どの値も親から継承しません。

VersionOverrides v1.0 スキーマと v1.1 スキーマの両方を実装するには、マニフェストは次の例のようになります。

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
  </VersionOverrides>
...
</OfficeApp>
```