---
title: マニフェスト ファイルの Scopes 要素
description: Scopes 要素には、アドインが外部リソースに接続するために必要なアクセス許可が含まれる。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 05582ae05c13fae8e2272de3fe6111c5ff639f938a817fd0b50ad22e4234d033
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087258"
---
# <a name="scopes-element"></a>Scopes 要素

Microsoft などの外部リソースに対してアドインに必要なアクセス許可がGraph。 Microsoft Graphリソースである場合、AppSource は Scopes 要素を使用して同意ダイアログ ボックスを作成します。 ユーザーがストアからアドインをインストールすると、ユーザーの Microsoft Graph のデータに対する指定されたアクセス許可をアドインに付与するように要求されます。

**スコープは** 、マニフェスト内の [WebApplicationInfo](webapplicationinfo.md) 要素と [Authorization 要素の子](authorization.md) 要素です。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Scope**                |  はい     |   アクセス許可の名前。たとえば、Files.Read.All またはプロファイルです。 |

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
