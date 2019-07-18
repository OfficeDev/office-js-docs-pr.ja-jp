---
title: マニフェスト ファイルの Scopes 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cdc9ebeb6fe4167a5ed5e9407f6ecc82d5b8d507
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771787"
---
# <a name="scopes-element"></a>Scopes 要素

アドインで必要な Microsoft Graph に対するアクセス許可が含まれます。 AppSource は、スコープ要素を使用して同意ダイアログボックスを作成します。 ユーザーがストアからアドインをインストールすると、ユーザーの Microsoft Graph のデータに対する指定されたアクセス許可をアドインに付与するように要求されます。

## <a name="child-elements"></a>子要素

|  要素 |  型  |  説明  |
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
