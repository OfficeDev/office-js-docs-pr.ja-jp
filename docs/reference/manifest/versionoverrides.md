---
title: マニフェスト ファイルの VersionOverrides 要素
description: アドイン マニフェスト (XML) ファイルの VersionOverrides 要素のリファレンス ドキュメントOffice。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 0a70ded82b4603b1ac70698947a4710a4a44b5b6
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555151"
---
# <a name="versionoverrides-element"></a>VersionOverrides 要素

アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。**VersionOverrides** は、マニフェスト内の [OfficeApp](officeapp.md) 要素の子要素です。この要素は、マニフェスト スキーマ v1.1 以降でサポートされていますが、VersionOverrides v1.0 または v1.1 スキーマで定義されています。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **xmlns**       |  はい  |  バージョンオーバーライド スキーマ名前空間。 許容される値は、この `<VersionOverrides>` 要素の **xsi:type** 値と親要素の **xsi:type** 値によって異なります `<OfficeApp>` 。 以下の [「名前空間の値](#namespace-values) 」を参照してください。|
|  **xsi:type**  |  はい  | スキーマのバージョン。現時点では、`VersionOverridesV1_0` および `VersionOverridesV1_1` のみが有効な値になります。 |

### <a name="namespace-values"></a>名前空間の値

次に、親要素の **xsi:type** 値に応じて **、xmlns** 値の必須値を示 `<OfficeApp>` します。

- **TaskPaneApp は** バージョン 1.0 のバージョンオーバーライドのみをサポートしており **、xmlns** は `http://schemas.microsoft.com/office/taskpaneappversionoverrides` ..
- **ContentApp は** バージョン 1.0 のバージョンオーバーライドのみをサポートしており **、xmlns** は `http://schemas.microsoft.com/office/contentappversionoverrides` ..
- **MailApp は** バージョン 1.0 およびバージョン 1.1 のバージョンオーバーライドをサポートしているので **、xmlns** の値はこの `<VersionOverrides>` 要素の **xsi:type** 値によって異なります。
    - **xsi:type** が の場合は `VersionOverridesV1_0` **、 xmlns** を指定する必要があります `http://schemas.microsoft.com/office/mailappversionoverrides` 。
    - **xsi:type** が の場合は `VersionOverridesV1_1` **、 xmlns** を指定する必要があります `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` 。

> [!NOTE]
> 現在、Outlook 2016以降のバージョンオーバーライド v1.1 スキーマと型のみがサポート `VersionOverridesV1_1` されています。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **説明**    |  いいえ   |  アドインについての説明。 これは、マニフェスト内の任意の親部分の `Description` 要素を上書きします。 説明のテキストは、**Resources** 要素の [LongString](resources.md) 要素の子要素に含まれています。 `resid` **Description** 要素の属性は 32 文字以内で、 `id` テキストを含む要素の属性の値に設定 `String` されます。|
|  **Requirements**  |  いいえ   |  アドインに必要な最小の Office.js のセットおよびバージョンを指定します。これは、マニフェストの親部分の `Requirements` 要素を上書きします。|
|  [Hosts](hosts.md)                |  はい  |  Officeアプリケーションのコレクションを指定します。 子の Hosts 要素は、マニフェストの親部分の Hosts 要素をオーバーライドします。  |
|  [Resources](resources.md)    |  はい  | マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。|
|  [EquivalentAddins](equivalentaddins.md)    |  いいえ  | Web アドインと同等のネイティブ (COM/XLL) アドインを指定します。 同等のネイティブ アドインがインストールされている場合、Web アドインはアクティブ化されません。|
|  **VersionOverrides**    |  いいえ  | より新しいスキーマ バージョンでアドイン コマンドを定義します。詳細については、「[複数のバージョンを実装する](#implementing-multiple-versions)」を参照してください。 |
|  [WebApplicationInfo](webapplicationinfo.md)    |  いいえ  | V2.0 Azure Active Directoryなど、セキュリティで保護されたトークン発行者へのアドインの登録に関する詳細を指定します。 |
|  [ExtendedPermissions](extendedpermissions.md) |  いいえ  |  拡張アクセス許可のコレクションを指定します。 |

### <a name="versionoverrides-example"></a>VersionOverrides の例

次に示す一般的な要素の例 `<VersionOverrides>` を示します。

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
