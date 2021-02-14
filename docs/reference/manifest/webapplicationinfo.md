---
title: マニフェスト ファイルの WebApplicationInfo 要素
description: アドイン マニフェスト (XML) ファイルOffice WebApplicationInfo 要素のリファレンス ドキュメント。
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: 037de49320a6d1a1ca7dce3446b4f4008a2f1331
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234164"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="40eab-103">WebApplicationInfo 要素</span><span class="sxs-lookup"><span data-stu-id="40eab-103">WebApplicationInfo element</span></span>

<span data-ttu-id="40eab-104">Office アドインでシングル サインオン (SSO) をサポートします。この要素には、次の両方としてのアドインに関する情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="40eab-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="40eab-105">クライアント アプリケーションがアクセス許可を必要とする可能性Office OAuth 2.0 リソース。</span><span class="sxs-lookup"><span data-stu-id="40eab-105">An OAuth 2.0 *resource* to which the Office client application might need permissions.</span></span>
- <span data-ttu-id="40eab-106">Microsoft Graph に対するアクセス許可を必要とする可能性のある OAuth 2.0 *クライアント*。</span><span class="sxs-lookup"><span data-stu-id="40eab-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="40eab-107">シングル サインオン API は現在、Word、Excel、Outlook、PowerPoint でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="40eab-107">The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="40eab-108">シングル サインオン API の現在のサポート状態に関する詳細は、「[Identity API の要件セット](../requirement-sets/identity-api-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="40eab-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="40eab-109">Outlook アドインで作業している場合は、Microsoft 365 テナントの先進認証が有効になっていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="40eab-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy.</span></span> <span data-ttu-id="40eab-110">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="40eab-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="40eab-111">**WebApplicationInfo** は、マニフェスト内の [VersionOverrides](versionoverrides.md) 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="40eab-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="40eab-112">子要素</span><span class="sxs-lookup"><span data-stu-id="40eab-112">Child elements</span></span>

|  <span data-ttu-id="40eab-113">要素</span><span class="sxs-lookup"><span data-stu-id="40eab-113">Element</span></span> |  <span data-ttu-id="40eab-114">必須</span><span class="sxs-lookup"><span data-stu-id="40eab-114">Required</span></span>  |  <span data-ttu-id="40eab-115">説明</span><span class="sxs-lookup"><span data-stu-id="40eab-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="40eab-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="40eab-116">**Id**</span></span>    |  <span data-ttu-id="40eab-117">はい</span><span class="sxs-lookup"><span data-stu-id="40eab-117">Yes</span></span>   |  <span data-ttu-id="40eab-118">Azure Active Directory (Azure AD) v2.0 エンドポイントに登録された、アドインの関連サービスの **アプリケーション ID**。</span><span class="sxs-lookup"><span data-stu-id="40eab-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="40eab-119">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="40eab-119">**MsaId**</span></span>    |  <span data-ttu-id="40eab-120">いいえ</span><span class="sxs-lookup"><span data-stu-id="40eab-120">No</span></span>   |  <span data-ttu-id="40eab-121">MSA 用のアドインの Web アプリケーションのクライアント ID を、アドインに登録msm.live.com。</span><span class="sxs-lookup"><span data-stu-id="40eab-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="40eab-122">**Resource**</span><span class="sxs-lookup"><span data-stu-id="40eab-122">**Resource**</span></span>  |  <span data-ttu-id="40eab-123">はい</span><span class="sxs-lookup"><span data-stu-id="40eab-123">Yes</span></span>   |  <span data-ttu-id="40eab-124">Azure Active Directory v2.0 エンドポイントに登録されたアドインの **アプリケーション ID URI** を指定します。</span><span class="sxs-lookup"><span data-stu-id="40eab-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="40eab-125">Scope</span><span class="sxs-lookup"><span data-stu-id="40eab-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="40eab-126">はい</span><span class="sxs-lookup"><span data-stu-id="40eab-126">Yes</span></span>  |  <span data-ttu-id="40eab-127">Microsoft Graph などのリソースに対してアドインに必要なアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="40eab-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="40eab-128">Authorizations</span><span class="sxs-lookup"><span data-stu-id="40eab-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="40eab-129">いいえ</span><span class="sxs-lookup"><span data-stu-id="40eab-129">No</span></span>   | <span data-ttu-id="40eab-130">アドインの Web アプリケーションが承認する必要がある外部リソースと必要なアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="40eab-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="40eab-131">WebApplicationInfo の例</span><span class="sxs-lookup"><span data-stu-id="40eab-131">WebApplicationInfo example</span></span>

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
