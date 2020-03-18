---
title: マニフェスト ファイルの VersionOverrides 要素
description: Office アドインのマニフェスト (XML) ファイルの VersionOverrides 要素の参照ドキュメント。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 055a796d269ffb230293639a7f69b2dde4e36eee
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717853"
---
# <a name="versionoverrides-element"></a>VersionOverrides 要素

アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。**VersionOverrides** は、マニフェスト内の [OfficeApp](./officeapp.md) 要素の子要素です。この要素は、マニフェスト スキーマ v1.1 以降でサポートされていますが、VersionOverrides v1.0 または v1.1 スキーマで定義されています。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **xmlns**       |  はい  |  VersionOverrides スキーマ名前空間。 指定できる値は、 `<VersionOverrides>`この要素の**xsi: type**値と親`<OfficeApp>`要素の**xsi: type**値によって異なります。 以下の[名前空間の値](#namespace-values)を参照してください。|
|  **xsi:type**  |  はい  | スキーマのバージョン。現時点では、`VersionOverridesV1_0` および `VersionOverridesV1_1` のみが有効な値になります。 |

### <a name="namespace-values"></a>名前空間の値

次に、親`<OfficeApp>`要素の**xsi: type**値に応じて、 **xmlns**値に必要な値を示します。

- **Task区画アプリ**は、バージョン1.0 の versionoverrides のみをサポート**xmlns**し、xmlns `http://schemas.microsoft.com/office/taskpaneappversionoverrides`はにする必要があります。
- **Contentapp**はバージョン1.0 の versionoverrides のみをサポートし、 **xmlns**は`http://schemas.microsoft.com/office/contentappversionoverrides`である必要があります。
- **Mailapp**はバージョン1.0 および1.1 の versionoverrides をサポートしているため、 **xmlns**の値`<VersionOverrides>`は次の要素の**xsi: type**値に応じて異なります。
    - **Xsi: type**が`VersionOverridesV1_0`の場合、 **xmlns**は`http://schemas.microsoft.com/office/mailappversionoverrides`でなければなりません。
    - **Xsi: type**が`VersionOverridesV1_1`の場合、 **xmlns**は`http://schemas.microsoft.com/office/mailappversionoverrides/1.1`でなければなりません。

> [!NOTE]
> 現在、Outlook 2016 以降では、VersionOverrides v1.1 スキーマと`VersionOverridesV1_1`種類をサポートしています。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **説明**    |  No/しない   |  アドインについての説明。これは、マニフェスト内の任意の親部分の `Description` 要素を上書きします。説明のテキストは、**Resources** 要素の [LongString](resources.md) 要素の子要素に含まれています。`resid` 要素の **** の属性は、テキストを含む `id` 要素の `String` 属性の値に設定されています。|
|  **Requirements**  |  いいえ   |  アドインに必要な最小の Office.js のセットおよびバージョンを指定します。これは、マニフェストの親部分の `Requirements` 要素を上書きします。|
|  [Hosts](hosts.md)                |  はい  |  Office ホストのコレクションを指定します。子の Host 要素は、マニフェストの親部分の Host 要素を上書きします。  |
|  [Resources](resources.md)    |  はい  | マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。|
|  [EquivalentAddins](equivalentaddins.md)    |  いいえ  | Web アドインと同等のネイティブ (COM/XLL) アドインを指定します。 同等のネイティブアドインがインストールされている場合、web アドインはアクティブ化されません。|
|  **VersionOverrides**    |  いいえ  | より新しいスキーマ バージョンでアドイン コマンドを定義します。詳細については、「[複数のバージョンを実装する](#implementing-multiple-versions)」を参照してください。 |
|  [WebApplicationInfo](webapplicationinfo.md)    |  いいえ  | Azure Active Directory v2.0 など、セキュリティで保護されたトークン発行者によるアドインの登録に関する詳細を指定します。 |
|  [ExtendedPermissions](extendedpermissions.md) |  いいえ  |  拡張アクセス許可のコレクションを指定します。<br><br>**重要**: [Office. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API は現在プレビュー段階のため、この`ExtendedPermissions`要素を使用するアドインは、appsource に発行することも、一元展開によって展開することもできません。 |

### <a name="versionoverrides-example"></a>VersionOverrides の例

通常、必須ではありません`<VersionOverrides>`が通常使用される子要素を含む一般的な要素の例を次に示します。

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

VersionOverrides v1.0 と v1.1 の両方のスキーマを実装するためのマニフェストは、次に示す例のようになります。

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
