---
title: コンテンツ アドインのマニフェスト ファイル内の VersionOverrides 1.0 要素
description: アドイン マニフェスト (XML) ファイルの VersionOverrides 要素 (コンテンツ) Officeドキュメントを参照してください。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2a9cd431f0e8fb4a7abe49103522e04900d9bcfd
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042189"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-content-add-in"></a>コンテンツ アドインのマニフェスト ファイル内の VersionOverrides 1.0 要素

この要素には、基本マニフェストでサポートされていない機能の情報が含まれています。

> [!NOTE]
> この記事では、要素の属性とバリエーションに関する重要な情報を含む [VersionOverrides](versionoverrides.md)要素の概要を理解している必要があります。

## <a name="child-elements"></a>子要素

次の表は **、VersionOverrides** 要素のバージョン 1.0 にのみ適用され、コンテンツ アドインにのみ適用されます。

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **VersionOverrides**    |  いいえ  | 現在、VersionOverrides 1.0 コンテンツ アドインでは使用できません。 |
|  [WebApplicationInfo](webapplicationinfo.md)    |  いいえ  | セキュリティで保護されたトークン発行者とのアドインの登録に関する詳細 (V2.0 などAzure Active Directory指定します。 |

## <a name="example"></a>例

次に簡単な例を示します。 詳細な例については、アドイン コード サンプルのサンプル アドインOffice[を参照してください](https://github.com/OfficeDev/PnP-OfficeAddins)。

```xml
<OfficeApp ... xsi:type="Content">
...
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/contentappversionoverrides" xsi:type="VersionOverridesV1_0">
        <WebApplicationInfo>
            <Id>$application_GUID here$</Id>
            <Resource>api://localhost:44355/$application_GUID here$</Resource>
            <Scopes>
                <Scope>Files.Read.All</Scope>
                <Scope>profile</Scope>
            </Scopes>
        </WebApplicationInfo>
    </VersionOverrides>
...
</OfficeApp>
```
