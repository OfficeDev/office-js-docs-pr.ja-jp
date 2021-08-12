---
title: マニフェスト ファイルの承認要素
description: アドインの Web アプリケーションで承認が必要な外部リソースと、必要なアクセス許可を指定します。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: af40a47c4ae30b6d18d3457704487027ff18ac92da2a3ae23cf1afe5c1e9b46a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087713"
---
# <a name="authorization-element"></a>Authorization 要素

アドインの Web アプリケーションで承認が必要な外部リソースと、必要なアクセス許可を指定します。

**承認** は、マニフェストの [Authorizations](authorizations.md) 要素の子要素です。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Resource**  |  はい   |  外部リソースの URL を指定します。|
|  [Scope](scopes.md)                |  必要  |  アドインがリソースに必要とするアクセス許可を指定します。  |

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
