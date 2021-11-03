---
title: マニフェスト ファイルの WebApplicationInfo 要素
description: アドイン マニフェスト (XML) ファイルOffice WebApplicationInfo 要素のリファレンス ドキュメント。
ms.date: 10/25/2021
ms.localizationpriority: medium
ms.openlocfilehash: bb21c584f516fc9e50bdd881a383fb03f01c753c
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681550"
---
# <a name="webapplicationinfo-element"></a>WebApplicationInfo 要素

Office アドインでシングル サインオン (SSO) をサポートします。この要素には、次の両方としてのアドインに関する情報が含まれます。

- クライアント アプリケーションがアクセス許可を必要とする可能性Office OAuth 2.0 リソース。
- Microsoft Graph に対するアクセス許可を必要とする可能性のある OAuth 2.0 *クライアント*。

> [!NOTE]
> シングル サインオン API は現在、Word、Excel、Outlook、およびPowerPoint。 シングル サインオン API の現在のサポート状態に関する詳細は、「[Identity API の要件セット](../requirement-sets/identity-api-requirement-sets.md)」を参照してください。 Outlook アドインで作業している場合は、Microsoft 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

**WebApplicationInfo** は、マニフェスト内の [VersionOverrides](versionoverrides.md) 要素の子要素です。  

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Id**    |  はい   |  Azure Active Directory (Azure AD) v2.0 エンドポイントに登録された、アドインの関連サービスの **アプリケーション ID**。|
|  **Resource**  |  はい   |  Azure Active Directory v2.0 エンドポイントに登録されたアドインの **アプリケーション ID URI** を指定します。|
|  [Scope](scopes.md)                |  はい  |  Microsoft などのリソースに対してアドインに必要なアクセス許可をGraph。  |

## <a name="webapplicationinfo-example"></a>WebApplicationInfo の例

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc</Resource>
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
