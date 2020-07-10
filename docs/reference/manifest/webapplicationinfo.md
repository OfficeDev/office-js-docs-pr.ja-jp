---
title: マニフェスト ファイルの WebApplicationInfo 要素
description: Office アドインのマニフェスト (XML) ファイルの VersionOverrides 要素の参照ドキュメント。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 5be75c6e202e40d60961a1b930ef43e583dee240
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094408"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="87666-103">WebApplicationInfo 要素</span><span class="sxs-lookup"><span data-stu-id="87666-103">WebApplicationInfo element</span></span>

<span data-ttu-id="87666-104">Office アドインでシングル サインオン (SSO) をサポートします。この要素には、次の両方としてのアドインに関する情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="87666-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="87666-105">Office ホスト アプリケーションでアクセス許可を必要とする可能性のある対象の OAuth 2.0 *リソース*。</span><span class="sxs-lookup"><span data-stu-id="87666-105">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="87666-106">Microsoft Graph に対するアクセス許可を必要とする可能性のある OAuth 2.0 *クライアント*。</span><span class="sxs-lookup"><span data-stu-id="87666-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="87666-107">現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="87666-107">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="87666-108">シングル サインオン API の現在のサポート状態に関する詳細は、「[Identity API の要件セット](../requirement-sets/identity-api-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="87666-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="87666-109">Outlook アドインで作業している場合は、Microsoft 365 テナントの先進認証を有効にしてください。</span><span class="sxs-lookup"><span data-stu-id="87666-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy.</span></span> <span data-ttu-id="87666-110">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="87666-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="87666-111">**WebApplicationInfo** は、マニフェスト内の [VersionOverrides](versionoverrides.md) 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="87666-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="87666-112">子要素</span><span class="sxs-lookup"><span data-stu-id="87666-112">Child elements</span></span>

|  <span data-ttu-id="87666-113">要素</span><span class="sxs-lookup"><span data-stu-id="87666-113">Element</span></span> |  <span data-ttu-id="87666-114">必須</span><span class="sxs-lookup"><span data-stu-id="87666-114">Required</span></span>  |  <span data-ttu-id="87666-115">説明</span><span class="sxs-lookup"><span data-stu-id="87666-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="87666-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="87666-116">**Id**</span></span>    |  <span data-ttu-id="87666-117">はい</span><span class="sxs-lookup"><span data-stu-id="87666-117">Yes</span></span>   |  <span data-ttu-id="87666-118">Azure Active Directory (Azure AD) v2.0 エンドポイントに登録された、アドインの関連サービスの**アプリケーション ID**。</span><span class="sxs-lookup"><span data-stu-id="87666-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="87666-119">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="87666-119">**MsaId**</span></span>    |  <span data-ttu-id="87666-120">いいえ</span><span class="sxs-lookup"><span data-stu-id="87666-120">No</span></span>   |  <span data-ttu-id="87666-121">Msm.live.com に登録されている、アドインの web アプリケーションのクライアント ID。</span><span class="sxs-lookup"><span data-stu-id="87666-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="87666-122">**Resource**</span><span class="sxs-lookup"><span data-stu-id="87666-122">**Resource**</span></span>  |  <span data-ttu-id="87666-123">はい</span><span class="sxs-lookup"><span data-stu-id="87666-123">Yes</span></span>   |  <span data-ttu-id="87666-124">Azure Active Directory v2.0 エンドポイントに登録されたアドインの**アプリケーション ID URI** を指定します。</span><span class="sxs-lookup"><span data-stu-id="87666-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="87666-125">Scope</span><span class="sxs-lookup"><span data-stu-id="87666-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="87666-126">はい</span><span class="sxs-lookup"><span data-stu-id="87666-126">Yes</span></span>  |  <span data-ttu-id="87666-127">Microsoft Graph などのリソースに対してアドインが必要とするアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="87666-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="87666-128">Authorizations</span><span class="sxs-lookup"><span data-stu-id="87666-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="87666-129">いいえ</span><span class="sxs-lookup"><span data-stu-id="87666-129">No</span></span>   | <span data-ttu-id="87666-130">アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="87666-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="87666-131">WebApplicationInfo の例</span><span class="sxs-lookup"><span data-stu-id="87666-131">WebApplicationInfo example</span></span>

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
