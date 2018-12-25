---
title: マニフェスト ファイルの Scopes 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 01d34481b14ac6a9186de07d352b9985dc1695a4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432642"
---
# <a name="scopes-element"></a>Scopes 要素

アドインで必要な Microsoft Graph に対するアクセス許可が含まれます。 Office ストアは、Scopes 要素を使用して同意ダイアログ ボックスを作成します。 ユーザーがストアからアドインをインストールすると、ユーザーの Microsoft Graph のデータに対する指定されたアクセス許可をアドインに付与するように要求されます。

## <a name="child-elements"></a>子要素

|  要素 |  支払期日  |  説明  |
|:-----|:-----|:-----|
|  **Scope**                |  string     |   Microsoft Graph に対するアクセス許可の名前。たとえば、Files.Read.All です。 |

## <a name="example"></a>例

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
