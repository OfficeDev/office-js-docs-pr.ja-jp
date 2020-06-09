---
title: マニフェスト ファイルの Scopes 要素
description: 範囲要素には、アドインが外部リソースに接続するために必要なアクセス許可が含まれています。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: be68033e86de736703d9d1593ad361918d5a147d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612242"
---
# <a name="scopes-element"></a>Scopes 要素

アドインが外部リソース (Microsoft Graph など) に対して必要とするアクセス許可が含まれます。 Microsoft Graph がリソースの場合、AppSource はスコープ要素を使用して同意ダイアログボックスを作成します。 ユーザーがストアからアドインをインストールすると、ユーザーの Microsoft Graph のデータに対する指定されたアクセス許可をアドインに付与するように要求されます。

**スコープ**は、マニフェスト内の[Webapplicationinfo](webapplicationinfo.md)要素と[Authorization](authorization.md)要素の子要素です。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Scope**                |  はい     |   アクセス許可の名前。たとえば、[すべて] または [プロファイル] を参照します。 |

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
