---
title: マニフェスト ファイルの Scopes 要素
description: Scopes 要素には、アドインが外部リソースに接続するために必要なアクセス許可が含まれる。
ms.date: 10/25/2021
ms.localizationpriority: medium
ms.openlocfilehash: 16e8a19a7aa73efa6aac00c915fde8d2b8647bad
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681536"
---
# <a name="scopes-element"></a>Scopes 要素

Microsoft などの外部リソースに対してアドインに必要なアクセス許可がGraph。 Microsoft Graphリソースである場合、AppSource は Scopes 要素を使用して同意ダイアログ ボックスを作成します。 ユーザーがストアからアドインをインストールすると、ユーザーの Microsoft Graph のデータに対する指定されたアクセス許可をアドインに付与するように要求されます。

**スコープは** 、マニフェスト内の [WebApplicationInfo](webapplicationinfo.md) 要素の子要素です。

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
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc<Resource>
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
