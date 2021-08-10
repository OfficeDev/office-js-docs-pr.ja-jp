---
title: マニフェスト ファイルの Authorizations 要素
description: アドインの Web アプリケーションで承認が必要な外部リソースと、必要なアクセス許可を指定します。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 068e6753e2e8e947e5e6e3c0885e7cd006165660862a37346eea114abb81a9b8
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092503"
---
# <a name="authorizations-element"></a>Authorizations 要素

アドインの Web アプリケーションで承認が必要な外部リソースと、必要なアクセス許可を指定します。

**承認は** 、マニフェスト内の [WebApplicationInfo](webapplicationinfo.md) 要素の子要素です。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Authorization](authorization.md)                |  必要     |   アドインの Web アプリケーションで承認が必要な外部リソースと、必要なスコープ (アクセス許可) を識別します。 |

## <a name="example"></a>例

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
