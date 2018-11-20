# <a name="webapplicationinfo-element"></a>WebApplicationInfo 要素

Office アドインでシングル サインオン (SSO) をサポートします。この要素には、次の両方としてのアドインに関する情報が含まれます。

- Office ホスト アプリケーションでアクセス許可を必要とする可能性のある対象の OAuth 2.0 *リソース*。
- Microsoft Graph に対するアクセス許可を必要とする可能性のある OAuth 2.0 *クライアント*。

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)」を参照してください。 Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

**WebApplicationInfo** は、マニフェスト内の [VersionOverrides](versionoverrides.md) 要素の子要素です。  

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Id**    |  はい   |  Azure Active Directory (Azure AD) v2.0 エンドポイントに登録された、アドインの関連サービスの**アプリケーション ID**。|
|  **Resource**  |  はい   |  Azure Active Directory v2.0 エンドポイントに登録されたアドインの**アプリケーション ID URI** を指定します。|
|  [Scope](scopes.md)                |  いいえ  |  アドインが必要とする Microsoft Graph に対するアクセス許可を指定します。  |

> [!NOTE] 
> 現時点では、アドインのリソースがそのホストと一致している必要があります。 Office は、所有権が証明できない限り、アドインのトークンを要求できません。現在これを行うには、リソースの完全修飾ドメイン名でアドインをホストします。

## <a name="webapplicationinfo-example"></a>WebApplicationInfo の例

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
