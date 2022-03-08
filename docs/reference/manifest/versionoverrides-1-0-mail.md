---
title: メール アドインのマニフェスト ファイル内の VersionOverrides 1.0 要素
description: アドイン マニフェスト (XML) ファイルOffice VersionOverrides 要素 (mail) のリファレンス ドキュメント。
ms.date: 02/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5288c085c94ff6fc8ab8fc31711c5c8fa142e946
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340674"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-mail-add-in"></a>メール アドインのマニフェスト ファイル内の VersionOverrides 1.0 要素

この要素には、基本マニフェストでサポートされていない機能の情報が含まれています。

> [!NOTE]
> この記事では、要素の属性とバリエーションに関する重要な情報を含む [VersionOverrides](versionoverrides.md) 要素の概要を理解している必要があります。

**アドインの種類:** メール

**次の VersionOverrides スキーマでのみ有効です**。

- メール 1.0

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [Mailbox 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)
- 一部の子要素は、追加の要件セットに関連付けられる場合があります。

## <a name="child-elements"></a>子要素

次の表は、 **VersionOverrides** 要素のバージョン 1.0 にのみ適用され、メール アドインにのみ適用されます。

> [!NOTE]
> iOS では、 **WebApplicationInfo だけが** サポートされています。 **VersionOverrides の他のすべての子要素は** 無視されます。

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [説明](#description)    |  いいえ   |  アドインについての説明。 |
|  [Requirements](requirements.md)  |  いいえ   |  親 **VersionOverrides** のマークアップを有効にするためにサポートする必要がある最小要件セットを指定します。 これは、マニフェストの *基本* 部分の **Requirements** 要素よりも常に制限が厳しい必要があります。|
|  [Hosts](hosts.md)                |  はい  |  アプリケーションのコレクションをOfficeします。 子の **Hosts** 要素は、マニフェストの親部分の **Hosts** 要素を上書きします。  |
|  [Resources](resources.md)    |  はい  | マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。|
|  **VersionOverrides**    |  いいえ  | より新しいスキーマ バージョンでアドイン コマンドを定義します。詳細については、「[複数のバージョンを実装する](#implementing-multiple-versions)」を参照してください。 |
|  [WebApplicationInfo](webapplicationinfo.md)    |  いいえ  | セキュリティで保護されたトークン発行者とのアドインの登録に関する詳細 (V2.0 などAzure Active Directory指定します。 |

### <a name="description"></a>説明

アドインの説明です。 マニフェストのすべての親部分の **Description** 要素を上書きします。 説明のテキストは、**Resources** 要素の [LongString](resources.md) 要素の子要素に含まれています。 Description `resid` 要素の属性は 32 `id` 文字以内で、Resources 要素に含まれる **ShortString** 要素の子要素の属性の値と一致 [する必要](resources.md)があります。 

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) 親 **VersionOverrides が** Taskpane 1.0 と入力されている場合。
- 親 **VersionOverrides が Mail** 1.0 と入力されている場合のメールボックス [1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)。
- 親 **VersionOverrides が Mail** 1.1 と入力されている場合のメールボックス [1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)。

## <a name="example"></a>例

次に簡単な例を示します。 より複雑な例については、アドイン コード サンプルのサンプル アドインOffice[を参照してください](https://github.com/OfficeDev/PnP-OfficeAddins)。

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

マニフェストは、VersionOverrides スキーマの異なるバージョンをサポートする **VersionOverrides** 要素の複数のバージョンを実装できます。 これは、必要に応じて新しいスキーマの新機能をサポートしながら、新しい機能をサポートしない古いクライアントをサポートするために実行できます。

複数のバージョンを実装するには、新しいバージョンの **VersionOverrides** `VersionOverrides` 要素が、以前のバージョンの要素の子である必要があります。 子 **VersionOverrides 要素** は、親から値を継承しない。

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
