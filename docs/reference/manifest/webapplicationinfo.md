# <a name="webapplicationinfo-element"></a><span data-ttu-id="54061-101">WebApplicationInfo 要素</span><span class="sxs-lookup"><span data-stu-id="54061-101">WebApplicationInfo element</span></span>

<span data-ttu-id="54061-102">Office アドインでシングル サインオン (SSO) をサポートします。この要素には、次の両方としてのアドインに関する情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="54061-102">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="54061-103">Office ホスト アプリケーションでアクセス許可を必要とする可能性のある対象の OAuth 2.0 *リソース*。</span><span class="sxs-lookup"><span data-stu-id="54061-103">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="54061-104">Microsoft Graph に対するアクセス許可を必要とする可能性のある OAuth 2.0 *クライアント*。</span><span class="sxs-lookup"><span data-stu-id="54061-104">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

<span data-ttu-id="54061-105">**WebApplicationInfo** は、マニフェスト内の [VersionOverrides](versionoverrides.md) 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="54061-105">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="54061-106">子要素</span><span class="sxs-lookup"><span data-stu-id="54061-106">Child elements</span></span>

|  <span data-ttu-id="54061-107">要素</span><span class="sxs-lookup"><span data-stu-id="54061-107">Element</span></span> |  <span data-ttu-id="54061-108">必須</span><span class="sxs-lookup"><span data-stu-id="54061-108">Required</span></span>  |  <span data-ttu-id="54061-109">説明</span><span class="sxs-lookup"><span data-stu-id="54061-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="54061-110">**ID**</span><span class="sxs-lookup"><span data-stu-id="54061-110">**Id**</span></span>    |  <span data-ttu-id="54061-111">はい</span><span class="sxs-lookup"><span data-stu-id="54061-111">Yes</span></span>   |  <span data-ttu-id="54061-112">Azure Active Directory (Azure AD) v2.0 エンドポイントに登録された、アドインの関連サービスの**アプリケーション ID**。</span><span class="sxs-lookup"><span data-stu-id="54061-112">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="54061-113">**リソース**</span><span class="sxs-lookup"><span data-stu-id="54061-113">**Resource**</span></span>  |  <span data-ttu-id="54061-114">はい</span><span class="sxs-lookup"><span data-stu-id="54061-114">Yes</span></span>   |  <span data-ttu-id="54061-115">Azure Active Directory v2.0 エンドポイントに登録されたアドインの**アプリケーション ID URI** を指定します。</span><span class="sxs-lookup"><span data-stu-id="54061-115">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="54061-116">Scope</span><span class="sxs-lookup"><span data-stu-id="54061-116">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="54061-117">いいえ</span><span class="sxs-lookup"><span data-stu-id="54061-117">No</span></span>  |  <span data-ttu-id="54061-118">アドインが必要とする Microsoft Graph に対するアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="54061-118">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="54061-119">現時点では、アドインのリソースがそのホストと一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="54061-119">Note: Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="54061-120">Office は、所有権が証明できない限り、アドインのトークンを要求できません。現在これを行うには、リソースの完全修飾ドメイン名でアドインをホストします。</span><span class="sxs-lookup"><span data-stu-id="54061-120">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="54061-121">WebApplicationInfo の例</span><span class="sxs-lookup"><span data-stu-id="54061-121">WebApplicationInfo example</span></span>

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
