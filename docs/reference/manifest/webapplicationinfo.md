# <a name="webapplicationinfo-element"></a><span data-ttu-id="ca058-101">WebApplicationInfo 要素</span><span class="sxs-lookup"><span data-stu-id="ca058-101">WebApplicationInfo element</span></span>

<span data-ttu-id="ca058-102">Office アドインでシングル サインオン (SSO) をサポートします。この要素には、次の両方としてのアドインに関する情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="ca058-102">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="ca058-103">Office ホスト アプリケーションでアクセス許可を必要とする可能性のある対象の OAuth 2.0 *リソース*。</span><span class="sxs-lookup"><span data-stu-id="ca058-103">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="ca058-104">Microsoft Graph に対するアクセス許可を必要とする可能性のある OAuth 2.0 *クライアント*。</span><span class="sxs-lookup"><span data-stu-id="ca058-104">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="ca058-105">現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="ca058-105">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="ca058-106">シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca058-106">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span> <span data-ttu-id="ca058-107">Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="ca058-107">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="ca058-108">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca058-108">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="ca058-109">**WebApplicationInfo** は、マニフェスト内の [VersionOverrides](versionoverrides.md) 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="ca058-109">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="ca058-110">子要素</span><span class="sxs-lookup"><span data-stu-id="ca058-110">Child elements</span></span>

|  <span data-ttu-id="ca058-111">要素</span><span class="sxs-lookup"><span data-stu-id="ca058-111">Element</span></span> |  <span data-ttu-id="ca058-112">必須</span><span class="sxs-lookup"><span data-stu-id="ca058-112">Required</span></span>  |  <span data-ttu-id="ca058-113">説明</span><span class="sxs-lookup"><span data-stu-id="ca058-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ca058-114">**Id**</span><span class="sxs-lookup"><span data-stu-id="ca058-114">**Id**</span></span>    |  <span data-ttu-id="ca058-115">はい</span><span class="sxs-lookup"><span data-stu-id="ca058-115">Yes</span></span>   |  <span data-ttu-id="ca058-116">Azure Active Directory (Azure AD) v2.0 エンドポイントに登録された、アドインの関連サービスの**アプリケーション ID**。</span><span class="sxs-lookup"><span data-stu-id="ca058-116">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="ca058-117">**Resource**</span><span class="sxs-lookup"><span data-stu-id="ca058-117">**Resource**</span></span>  |  <span data-ttu-id="ca058-118">はい</span><span class="sxs-lookup"><span data-stu-id="ca058-118">Yes</span></span>   |  <span data-ttu-id="ca058-119">Azure Active Directory v2.0 エンドポイントに登録されたアドインの**アプリケーション ID URI** を指定します。</span><span class="sxs-lookup"><span data-stu-id="ca058-119">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="ca058-120">Scope</span><span class="sxs-lookup"><span data-stu-id="ca058-120">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="ca058-121">いいえ</span><span class="sxs-lookup"><span data-stu-id="ca058-121">No</span></span>  |  <span data-ttu-id="ca058-122">アドインが必要とする Microsoft Graph に対するアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="ca058-122">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="ca058-123">現時点では、アドインのリソースがそのホストと一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca058-123">Note: Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="ca058-124">Office は、所有権が証明できない限り、アドインのトークンを要求できません。現在これを行うには、リソースの完全修飾ドメイン名でアドインをホストします。</span><span class="sxs-lookup"><span data-stu-id="ca058-124">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="ca058-125">WebApplicationInfo の例</span><span class="sxs-lookup"><span data-stu-id="ca058-125">WebApplicationInfo example</span></span>

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
