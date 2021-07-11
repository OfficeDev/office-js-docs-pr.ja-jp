---
title: マニフェスト ファイルの VersionOverrides 要素
description: アドイン マニフェスト (XML) ファイルOffice VersionOverrides 要素のリファレンス ドキュメント。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 787ba8e7d90900cc72d6c5e9370d68ced0faee2f
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348658"
---
# <a name="versionoverrides-element"></a>VersionOverrides 要素

アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。**VersionOverrides** は、マニフェスト内の [OfficeApp](officeapp.md) 要素の子要素です。この要素は、マニフェスト スキーマ v1.1 以降でサポートされていますが、VersionOverrides v1.0 または v1.1 スキーマで定義されています。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **xmlns**       |  はい  |  VersionOverrides スキーマ名前空間。 許可される値は、この要素の `<VersionOverrides>` **xsi:type** 値と親要素の **xsi:type** 値によって異 `<OfficeApp>` なります。 以下の [名前空間の値を参照](#namespace-values) してください。|
|  **xsi:type**  |  はい  | スキーマのバージョン。現時点では、`VersionOverridesV1_0` および `VersionOverridesV1_1` のみが有効な値になります。 |

### <a name="namespace-values"></a>名前空間の値

親要素の **xsi:type** 値に応じて **、xmlns** 値の必要な値を次に示 `<OfficeApp>` します。

- **TaskPaneApp は** VersionOverrides のバージョン 1.0 のみをサポートし **、xmlns は** `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .
- **ContentApp** は VersionOverrides のバージョン 1.0 のみをサポートし **、xmlns は** `http://schemas.microsoft.com/office/contentappversionoverrides` .
- **MailApp** は VersionOverrides のバージョン 1.0 と 1.1 をサポートしています。 **したがって、xmlns** の値は、この要素の `<VersionOverrides>` **xsi:type** 値によって異なります。
    - **xsi:type がである** 場合 `VersionOverridesV1_0` は **、xmlns を** 指定する必要があります `http://schemas.microsoft.com/office/mailappversionoverrides` 。
    - **xsi:type がである** 場合 `VersionOverridesV1_1` は **、xmlns を** 指定する必要があります `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` 。

> [!NOTE]
> 現在のところ、Outlook 2016以降は VersionOverrides v1.1 スキーマと型をサポート `VersionOverridesV1_1` しています。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **説明**    |  いいえ   |  アドインについての説明。 これは、マニフェスト内の任意の親部分の `Description` 要素を上書きします。 説明のテキストは、**Resources** 要素の [LongString](resources.md) 要素の子要素に含まれています。 `resid`Description 要素の **属性** は 32 文字以内で、テキストを含む要素の属性の値 `id` `String` に設定されます。|
|  **Requirements**  |  いいえ   |  アドインに必要な最小の Office.js のセットおよびバージョンを指定します。これは、マニフェストの親部分の `Requirements` 要素を上書きします。|
|  [Hosts](hosts.md)                |  はい  |  アプリケーションのコレクションをOfficeします。 子 Hosts 要素は、マニフェストの親部分にある Hosts 要素をオーバーライドします。  |
|  [Resources](resources.md)    |  はい  | マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。|
|  [EquivalentAddins](equivalentaddins.md)    |  いいえ  | Web アドインと同等のネイティブ (COM/XLL) アドインを指定します。 同等のネイティブ アドインがインストールされている場合、Web アドインはアクティブ化されません。|
|  **VersionOverrides**    |  いいえ  | より新しいスキーマ バージョンでアドイン コマンドを定義します。詳細については、「[複数のバージョンを実装する](#implementing-multiple-versions)」を参照してください。 |
|  [WebApplicationInfo](webapplicationinfo.md)    |  いいえ  | セキュリティで保護されたトークン発行者とのアドインの登録に関する詳細 (V2.0 などAzure Active Directory指定します。 |
|  [ExtendedPermissions](extendedpermissions.md) |  いいえ  |  拡張アクセス許可のコレクションを指定します。 |

### <a name="versionoverrides-example"></a>VersionOverrides の例

次に示すのは、一般的な要素の例です。一部の子要素は必須ではなく、通常 `<VersionOverrides>` は使用されます。

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
