---
title: マニフェスト ファイルの承認要素
description: アドインの Web アプリケーションで承認が必要な外部リソースと、必要なアクセス許可を指定します。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8d3dd31a212a7de00ff4dbf263e8593a8ec2898
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937982"
---
# <a name="authorization-element"></a>Authorization 要素

アドインの Web アプリケーションで承認が必要な外部リソースと、必要なアクセス許可を指定します。

**承認** は、マニフェストの [Authorizations](authorizations.md) 要素の子要素です。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Resource**  |  はい   |  外部リソースの URL を指定します。|
|  [Scope](scopes.md)                |  はい  |  アドインがリソースに必要とするアクセス許可を指定します。  |

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
