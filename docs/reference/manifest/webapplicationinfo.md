---
title: マニフェスト ファイルの WebApplicationInfo 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 1328dc40e98c321c9c4b7d3d692da8c8bdd29492
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389199"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="aaaa9-102">WebApplicationInfo 要素</span><span class="sxs-lookup"><span data-stu-id="aaaa9-102">WebApplicationInfo element</span></span>

<span data-ttu-id="aaaa9-103">Office アドインでシングル サインオン (SSO) をサポートします。この要素には、次の両方としてのアドインに関する情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-103">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="aaaa9-104">Office ホスト アプリケーションでアクセス許可を必要とする可能性のある対象の OAuth 2.0 *リソース*。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-104">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="aaaa9-105">Microsoft Graph に対するアクセス許可を必要とする可能性のある OAuth 2.0 *クライアント*。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-105">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="aaaa9-106">現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-106">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="aaaa9-107">シングル サインオン API の現在のサポート状態に関する詳細は、「[Identity API の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-107">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="aaaa9-108">Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="aaaa9-109">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-109">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="aaaa9-110">**WebApplicationInfo** は、マニフェスト内の [VersionOverrides](versionoverrides.md) 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-110">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="aaaa9-111">子要素</span><span class="sxs-lookup"><span data-stu-id="aaaa9-111">Child elements</span></span>

|  <span data-ttu-id="aaaa9-112">要素</span><span class="sxs-lookup"><span data-stu-id="aaaa9-112">Element</span></span> |  <span data-ttu-id="aaaa9-113">必須</span><span class="sxs-lookup"><span data-stu-id="aaaa9-113">Required</span></span>  |  <span data-ttu-id="aaaa9-114">説明</span><span class="sxs-lookup"><span data-stu-id="aaaa9-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="aaaa9-115">**Id**</span><span class="sxs-lookup"><span data-stu-id="aaaa9-115">**Id**</span></span>    |  <span data-ttu-id="aaaa9-116">はい</span><span class="sxs-lookup"><span data-stu-id="aaaa9-116">Yes</span></span>   |  <span data-ttu-id="aaaa9-117">Azure Active Directory (Azure AD) v2.0 エンドポイントに登録された、アドインの関連サービスの**アプリケーション ID**。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-117">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="aaaa9-118">**Resource**</span><span class="sxs-lookup"><span data-stu-id="aaaa9-118">**Resource**</span></span>  |  <span data-ttu-id="aaaa9-119">はい</span><span class="sxs-lookup"><span data-stu-id="aaaa9-119">Yes</span></span>   |  <span data-ttu-id="aaaa9-120">Azure Active Directory v2.0 エンドポイントに登録されたアドインの**アプリケーション ID URI** を指定します。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-120">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="aaaa9-121">Scope</span><span class="sxs-lookup"><span data-stu-id="aaaa9-121">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="aaaa9-122">いいえ</span><span class="sxs-lookup"><span data-stu-id="aaaa9-122">No</span></span>  |  <span data-ttu-id="aaaa9-123">アドインが必要とする Microsoft Graph に対するアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-123">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="aaaa9-124">現時点では、アドインのリソースがそのホストと一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-124">Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="aaaa9-125">Office は、所有権が証明できない限り、アドインのトークンを要求できません。現在これを行うには、リソースの完全修飾ドメイン名でアドインをホストします。</span><span class="sxs-lookup"><span data-stu-id="aaaa9-125">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="aaaa9-126">WebApplicationInfo の例</span><span class="sxs-lookup"><span data-stu-id="aaaa9-126">WebApplicationInfo example</span></span>

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
