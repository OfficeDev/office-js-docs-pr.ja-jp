---
title: マニフェスト ファイルの WebApplicationInfo 要素
description: Office アドインのマニフェスト (XML) ファイルの VersionOverrides 要素の参照ドキュメント。
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: b64739f388720abb872fe2407f7e7d3031b5ec70
ms.sourcegitcommit: 8fdd7369bfd97a273e222a0404e337ba2b8807b0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/05/2020
ms.locfileid: "46573120"
---
# <a name="webapplicationinfo-element"></a>WebApplicationInfo 要素

Office アドインでシングル サインオン (SSO) をサポートします。この要素には、次の両方としてのアドインに関する情報が含まれます。

- Office ホスト アプリケーションでアクセス許可を必要とする可能性のある対象の OAuth 2.0 *リソース*。
- Microsoft Graph に対するアクセス許可を必要とする可能性のある OAuth 2.0 *クライアント*。

> [!NOTE]
> 現在、シングルサインオン API は Word、Excel、Outlook、および PowerPoint でサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[Identity API の要件セット](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)」を参照してください。 Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

**WebApplicationInfo** は、マニフェスト内の [VersionOverrides](versionoverrides.md) 要素の子要素です。  

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Id**    |  はい   |  Azure Active Directory (Azure AD) v2.0 エンドポイントに登録された、アドインの関連サービスの**アプリケーション ID**。|
|  **MsaId**    |  いいえ   |  Msm.live.com に登録されている、アドインの web アプリケーションのクライアント ID。|
|  **Resource**  |  はい   |  Azure Active Directory v2.0 エンドポイントに登録されたアドインの**アプリケーション ID URI** を指定します。|
|  [Scope](scopes.md)                |  はい  |  Microsoft Graph などのリソースに対してアドインが必要とするアクセス許可を指定します。  |
|  [Authorizations](authorizations.md)  |  いいえ   | アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。|

## <a name="webapplicationinfo-example"></a>WebApplicationInfo の例

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
