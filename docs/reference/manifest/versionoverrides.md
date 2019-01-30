---
title: マニフェスト ファイルの VersionOverrides 要素
description: ''
ms.date: 01/29/2019
localization_priority: Normal
ms.openlocfilehash: 897c2203ef6ae84911b7f269ee8a2c88aec36bd0
ms.sourcegitcommit: 2e4b97f0252ff3dd908a3aa7a9720f0cb50b855d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/30/2019
ms.locfileid: "29635910"
---
# <a name="versionoverrides-element"></a>VersionOverrides 要素

アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。**VersionOverrides** は、マニフェスト内の [OfficeApp](./officeapp.md) 要素の子要素です。この要素は、マニフェスト スキーマ v1.1 以降でサポートされていますが、VersionOverrides v1.0 または v1.1 スキーマで定義されています。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **xmlns**       |  はい  |  スキーマの場所。`xsi:type` が `VersionOverridesV1_0` の場合は `http://schemas.microsoft.com/office/mailappversionoverrides` にする必要があり、`xsi:type` が `VersionOverridesV1_1` の場合は `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` にする必要があります。|
|  **xsi:type**  |  はい  | スキーマのバージョン。現時点では、`VersionOverridesV1_0` および `VersionOverridesV1_1` のみが有効な値になります。 |

> [!NOTE]
> 2016 またはそれ以降、現在は Outlook には、VersionOverrides v1.1 のスキーマがサポートされていると、`VersionOverridesV1_1`型です。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Description**    |  いいえ   |  アドインについての説明。これは、マニフェスト内の任意の親部分の `Description` 要素を上書きします。説明のテキストは、[Resources](./resources.md) 要素の **LongString** 要素の子要素に含まれています。**Description** 要素の `resid` の属性は、テキストを含む `String` 要素の `id` 属性の値に設定されています。|
|  **Requirements**  |  いいえ   |  アドインに必要な最小の Office.js のセットおよびバージョンを指定します。これは、マニフェストの親部分の `Requirements` 要素を上書きします。|
|  [Hosts](./hosts.md)                |  はい  |  Office ホストのコレクションを指定します。子の Host 要素は、マニフェストの親部分の Host 要素を上書きします。  |
|  [Resources](./resources.md)    |  はい  | マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。|
|  **VersionOverrides**    |  いいえ  | より新しいスキーマ バージョンでアドイン コマンドを定義します。詳細については、「[複数のバージョンを実装する](#implementing-multiple-versions)」を参照してください。 |
|  **WebApplicationInfo**    |  いいえ  | アドインの関連 Web アプリケーションについての詳細を指定します。 |

### <a name="versionoverrides-example"></a>VersionOverrides の例

次の一般的な例では`<VersionOverrides>`、必須ではありませんが、通常使用されるいくつかの子要素を含む要素です。

```xml
<OfficeApp>
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

VersionOverrides v1.0 と v1.1 の両方のスキーマを実装するためのマニフェストは、次に示す例のようになります。

```xml
<OfficeApp>
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
